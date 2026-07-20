// ═══════════════════════════════════════════════════════════════════════════
// NEXTA — /api/ors-proxy
// ═══════════════════════════════════════════════════════════════════════════
// Vercel Serverless Function (Node.js) — repassa a requisição de rota pra
// OpenRouteService, mas segurando a chave de API AQUI (variável de ambiente
// do servidor), nunca exposta no bundle do navegador.
//
// Antes: a chave (ORS_API_KEY) ficava literalmente escrita no código-fonte
// de nexta-frota-dashboard.js — qualquer pessoa que abrisse o DevTools
// conseguia copiar e usar a cota diária de 2.000 requisições da empresa.
//
// Configuração necessária na Vercel (Project Settings → Environment
// Variables): criar ORS_API_KEY com o valor da chave da OpenRouteService.
// Recomendado GERAR UMA CHAVE NOVA em openrouteservice.org (a antiga já
// esteve exposta publicamente no bundle, então deve ser considerada
// comprometida — trocar em vez de reaproveitar).
// ═══════════════════════════════════════════════════════════════════════════

module.exports = async function handler(req, res) {
  // CORS: o front-end é servido pelo mesmo domínio da Vercel, mas libera
  // explicitamente pra facilitar teste local/preview em domínios diferentes.
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Método não permitido — use POST.' });
    return;
  }

  const apiKey = process.env.ORS_API_KEY;
  if (!apiKey) {
    console.error('[ors-proxy] ORS_API_KEY não configurada nas variáveis de ambiente da Vercel.');
    res.status(500).json({ error: 'Serviço de rota não configurado no servidor (ORS_API_KEY ausente).' });
    return;
  }

  // Corpo esperado (o mesmo formato que o front já montava direto pra ORS):
  // { coordinates: [[lon,lat],[lon,lat]], options: { vehicle_type: 'hgv', ... } }
  const body = req.body || {};
  if (!Array.isArray(body.coordinates) || body.coordinates.length < 2) {
    res.status(400).json({ error: 'Corpo inválido — esperado { coordinates: [[lon,lat],[lon,lat]] }.' });
    return;
  }

  try {
    const orsRes = await fetch('https://api.openrouteservice.org/v2/directions/driving-hgv/geojson', {
      method: 'POST',
      headers: {
        Authorization: apiKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });
    const data = await orsRes.json().catch(() => null);
    res.status(orsRes.status).json(data);
  } catch (e) {
    console.error('[ors-proxy] falha ao contatar a OpenRouteService:', e);
    res.status(502).json({ error: 'Falha ao contatar a OpenRouteService: ' + e.message });
  }
};
