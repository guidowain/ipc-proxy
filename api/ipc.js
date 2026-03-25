export default async function handler(req, res) {
  const url = 'https://apis.datos.gob.ar/series/api/series/?ids=145.3_INGNACUAL_DICI_M_38&limit=4&sort=desc&representation_mode=percent_change&format=json';
  const data = await fetch(url).then(r => r.json());
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.json(data);
}
