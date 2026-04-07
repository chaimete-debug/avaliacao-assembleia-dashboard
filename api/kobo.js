// api/kobo.js — Vercel serverless function
// Proxies requests to KoboToolbox, adding auth and CORS headers

const ASSET = 'aQgBUEPn9wyVBAbWJPUbSQ';
const TOKEN = '4c6c1683faa457dfc69c8fc743d3374d8b4fd800';
const BASE  = 'https://eu.kobotoolbox.org/api/v2';

export default async function handler(req, res) {
  // Allow browser requests from any origin
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const all = [];
    let url = `${BASE}/assets/${ASSET}/data/?format=json&limit=30000`;

    while (url) {
      const resp = await fetch(url, {
        headers: { Authorization: `Token ${TOKEN}` },
      });

      if (!resp.ok) {
        return res.status(resp.status).json({
          error: `KoboToolbox responded with ${resp.status}: ${resp.statusText}`,
        });
      }

      const json = await resp.json();
      (json.results || []).forEach(r => all.push(normalise(r)));
      url = json.next || null;
    }

    return res.status(200).json({ count: all.length, results: all });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}

function normalise(r) {
  const out = {};
  Object.entries(r).forEach(([k, v]) => {
    const clean = k.startsWith('_') ? k : k.split('/').pop();
    out[clean] = v === null || v === undefined ? '' : String(v).trim();
  });
  return out;
}
