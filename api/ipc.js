import XLSX from 'xlsx';

const XLS_URL = 'https://www.indec.gob.ar/ftp/cuadros/economia/sh_ipc_aperturas.xls';

const MONTHS = {
  ene: 1,
  feb: 2,
  mar: 3,
  abr: 4,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  ago: 8,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dic: 12,
  dec: 12,
};

function normalizeText(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function excelDateToISO(serial) {
  if (typeof serial !== 'number' || !Number.isFinite(serial)) return null;

  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const date = new Date(utcValue * 1000);

  if (Number.isNaN(date.getTime())) return null;

  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  return `${y}-${m}-01`;
}

function parseMonthCell(cell) {
  if (cell == null || cell === '') return null;

  if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
    const y = cell.getFullYear();
    const m = String(cell.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}-01`;
  }

  if (typeof cell === 'number') {
    return excelDateToISO(cell);
  }

  const raw = String(cell).trim();

  const match = raw.match(/^([A-Za-zÁÉÍÓÚáéíóú]{3})[-\/\s]?(\d{2,4})$/);
  if (!match) return null;

  const mon = normalizeText(match[1]).slice(0, 3);
  const month = MONTHS[mon];
  if (!month) return null;

  let year = Number(match[2]);
  if (year < 100) year += year >= 70 ? 1900 : 2000;

  return `${year}-${String(month).padStart(2, '0')}-01`;
}

function isNumeric(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function findHeaderRow(rows) {
  let best = null;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;

    let monthCount = 0;
    for (let c = 0; c < row.length; c++) {
      if (parseMonthCell(row[c])) monthCount++;
    }

    if (monthCount >= 12) {
      if (!best || monthCount > best.monthCount) {
        best = { rowIndex: r, monthCount };
      }
    }
  }

  return best?.rowIndex ?? -1;
}

function extractMonthColumns(headerRow) {
  const cols = [];

  for (let c = 0; c < headerRow.length; c++) {
    const iso = parseMonthCell(headerRow[c]);
    if (iso) {
      cols.push({ colIndex: c, iso });
    }
  }

  return cols;
}

function findNivelGeneralRow(rows, headerRowIndex) {
  const candidates = [];

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!Array.isArray(row) || row.length === 0) continue;

    const firstCell = normalizeText(row[0]);
    if (firstCell === 'nivel general') {
      const numericCount = row.filter(isNumeric).length;
      candidates.push({ rowIndex: r, numericCount });
    }
  }

  if (candidates.length === 0) return -1;

  candidates.sort((a, b) => b.numericCount - a.numericCount || a.rowIndex - b.rowIndex);
  return candidates[0].rowIndex;
}

function buildSeries(rows) {
  const headerRowIndex = findHeaderRow(rows);
  if (headerRowIndex === -1) {
    throw new Error('No encontré la fila de meses en la planilla de INDEC');
  }

  const headerRow = rows[headerRowIndex];
  const monthCols = extractMonthColumns(headerRow);

  if (monthCols.length < 2) {
    throw new Error('No encontré suficientes columnas de meses');
  }

  const nivelRowIndex = findNivelGeneralRow(rows, headerRowIndex);
  if (nivelRowIndex === -1) {
    throw new Error('No encontré la fila "Nivel general"');
  }

  const nivelRow = rows[nivelRowIndex];

  const series = monthCols
    .map(({ colIndex, iso }) => {
      const value = nivelRow[colIndex];
      if (!isNumeric(value)) return null;

      return [iso, value / 100];
    })
    .filter(Boolean);

  if (series.length < 2) {
    throw new Error('No pude extraer suficientes datos de IPC');
  }

  series.sort((a, b) => a[0].localeCompare(b[0]));
  return series;
}

export default async function handler(req, res) {
  try {
    const response = await fetch(XLS_URL, {
      headers: {
        'User-Agent': 'ipc-proxy',
        Accept: '*/*',
      },
    });

    if (!response.ok) {
      return res.status(response.status).json({
        error: 'No se pudo descargar la planilla oficial de INDEC',
      });
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, {
      header: 1,
      raw: true,
      defval: null,
    });

    const series = buildSeries(rows);
    const desc = [...series].reverse().slice(0, 6);

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 's-maxage=21600, stale-while-revalidate=86400');

    return res.status(200).json({
      source: 'INDEC',
      source_url: XLS_URL,
      updated_at: new Date().toISOString(),
      data: desc,
    });
  } catch (error) {
    return res.status(500).json({
      error: 'Error interno del proxy',
      detail: error.message,
    });
  }
}
