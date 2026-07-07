// ═══════════════════════════════════════════════════════════════════════════
// NEXTA — DETECTOR DE PEDÁGIOS E CÁLCULO DE CUSTOS
// ═══════════════════════════════════════════════════════════════════════════
// Detecta pedágios na rota usando geolocalização e distância de pontos
// Integra valores ao custo final da viagem com tarifas variáveis por eixo

// ─── Base de Dados de Pedágios Brasileiros ─────────────────────────────────
// Formato: { nome, lat, lon, tarifas: { 2: valor, 3: valor, 5: valor, 6: valor }, rodovia, regiao, modelo }
// Tarifas por número de eixos — obterTarifaPedagio() calcula proporcionalmente
// pra qualquer eixo mesmo se só o valor de 2 eixos estiver cadastrado.
// modelo: 'praca' (padrão, cabine física com vale-pedágio) ou 'free_flow'
// (pórtico eletrônico sem cabine — cobrança automática por tag/placa, sem
// vale-pedágio físico). Se "modelo" não for informado, assume-se 'praca'.
const PEDAGIOS_BR = [
  // INTERIOR DE SP - Rodovia Anhanguera
  { nome: 'Pedágio Jundiaí', lat: -23.1881, lon: -46.8786, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Anhanguera', regiao: 'SP' },
  { nome: 'Pedágio Campinas (Sul)', lat: -22.9062, lon: -47.0583, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'Anhanguera', regiao: 'SP' },
  { nome: 'Pedágio Campinas (Norte)', lat: -22.8856, lon: -47.0628, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'Anhanguera', regiao: 'SP' },
  
  // INTERIOR DE SP - Rodovia Imigrantes / Cônego Domênico
  { nome: 'Pedágio Limeira', lat: -22.5644, lon: -47.4090, tarifas: { 2: 24.60, 3: 36.90, 5: 61.50, 6: 73.80 }, rodovia: 'Imigrantes', regiao: 'SP' },
  { nome: 'Pedágio Americana', lat: -22.7408, lon: -47.3261, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Imigrantes', regiao: 'SP' },
  { nome: 'Pedágio Paulínia', lat: -22.7589, lon: -47.1539, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'Cônego Domênico', regiao: 'SP' },
  
  // INTERIOR DE SP - Rodovia Castelo Branco
  { nome: 'Pedágio Barueri', lat: -23.5090, lon: -46.8667, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Castelo Branco', regiao: 'SP' },
  { nome: 'Pedágio Osasco', lat: -23.5291, lon: -46.7892, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Castelo Branco', regiao: 'SP' },
  { nome: 'Pedágio Itapetininga', lat: -23.5873, lon: -48.0486, tarifas: { 2: 34.50, 3: 51.75, 5: 86.25, 6: 103.50 }, rodovia: 'Castelo Branco', regiao: 'SP' },
  
  // INTERIOR DE SP - Rodovia Raposo Tavares
  { nome: 'Pedágio Mairinque', lat: -23.5291, lon: -47.3389, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Raposo Tavares', regiao: 'SP' },
  { nome: 'Pedágio Salto', lat: -23.1962, lon: -47.3089, tarifas: { 2: 24.60, 3: 36.90, 5: 61.50, 6: 73.80 }, rodovia: 'Raposo Tavares', regiao: 'SP' },
  
  // INTERIOR DE SP - Rodovia Washington Luiz
  { nome: 'Pedágio Ribeirão Preto', lat: -21.1949, lon: -47.8078, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'Washington Luiz', regiao: 'SP' },
  { nome: 'Pedágio Araraquara', lat: -22.0140, lon: -48.1738, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Washington Luiz', regiao: 'SP' },
  
  // MINAS GERAIS - Rodovia Fernão Dias
  { nome: 'Pedágio Bragança Paulista', lat: -22.9574, lon: -46.5306, tarifas: { 2: 24.60, 3: 36.90, 5: 61.50, 6: 73.80 }, rodovia: 'Fernão Dias', regiao: 'SP/MG' },
  { nome: 'Pedágio Entre Rios', lat: -22.5708, lon: -46.3281, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'Fernão Dias', regiao: 'MG' },
  { nome: 'Pedágio Três Corações', lat: -21.7879, lon: -45.2574, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'Fernão Dias', regiao: 'MG' },
  
  // MINAS GERAIS - Outras rodovias
  { nome: 'Pedágio Juiz de Fora', lat: -21.7626, lon: -43.3551, tarifas: { 2: 26.80, 3: 40.20, 5: 67.00, 6: 80.40 }, rodovia: 'BR-116', regiao: 'MG' },
  { nome: 'Pedágio Belo Horizonte', lat: -19.8661, lon: -43.9758, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'BR-381', regiao: 'MG' },
  
  // RIO DE JANEIRO
  { nome: 'Pedágio Volta Redonda', lat: -22.5215, lon: -44.0687, tarifas: { 2: 26.80, 3: 40.20, 5: 67.00, 6: 80.40 }, rodovia: 'BR-116', regiao: 'RJ' },
  { nome: 'Pedágio Dutra', lat: -22.4650, lon: -44.1987, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'BR-116', regiao: 'RJ' },
  
  // GOIÁS
  { nome: 'Pedágio Goiânia', lat: -15.7867, lon: -48.0055, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'BR-040', regiao: 'GO' },
  { nome: 'Pedágio Anápolis', lat: -15.9436, lon: -48.9564, tarifas: { 2: 28.70, 3: 43.05, 5: 71.75, 6: 86.10 }, rodovia: 'BR-153', regiao: 'GO' },
  
  // BAHIA
  { nome: 'Pedágio Feira de Santana', lat: -12.2652, lon: -38.9667, tarifas: { 2: 34.50, 3: 51.75, 5: 86.25, 6: 103.50 }, rodovia: 'BR-116', regiao: 'BA' },
  { nome: 'Pedágio Salvador', lat: -12.9789, lon: -38.5106, tarifas: { 2: 31.50, 3: 47.25, 5: 78.75, 6: 94.50 }, rodovia: 'BR-116', regiao: 'BA' },
  
  // CEARÁ
  { nome: 'Pedágio Fortaleza', lat: -3.7319, lon: -38.5433, tarifas: { 2: 36.00, 3: 54.00, 5: 90.00, 6: 108.00 }, rodovia: 'BR-116', regiao: 'CE' },
  
  // PERNAMBUCO
  { nome: 'Pedágio Recife', lat: -8.0476, lon: -34.8770, tarifas: { 2: 34.50, 3: 51.75, 5: 86.25, 6: 103.50 }, rodovia: 'BR-101', regiao: 'PE' },

  // LITORAL SP - Pórticos Free Flow (Concessionária Novo Litoral / CNL)
  // Sem cabine física — cobrança automática por tag ou placa (30 dias p/ pagar).
  // ATENÇÃO: coordenada abaixo é uma ESTIMATIVA a partir do km da rodovia
  // (não achei o lat/lon oficial do pórtico em fonte pública). Se o alerta não
  // disparar na rota certa (ou disparar num lugar errado), ajuste lat/lon aqui
  // comparando com o ponto real no mapa da viagem.
  { nome: 'Pórtico Free Flow Bertioga (P03)', lat: -23.8320, lon: -46.1500, tarifas: { 2: 6.95 }, rodovia: 'SP-098 (Mogi-Bertioga, km 92)', regiao: 'SP', modelo: 'free_flow' },
];

// ─── Haversine: Distância entre dois pontos (lat/lon) ──────────────────────
function distanciaHaversine(lat1, lon1, lat2, lon2) {
  const R = 6371; // Raio da Terra em km
  const toRad = (deg) => (deg * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

// ─── Obter tarifa de um pedágio para número de eixos ───────────────────────
// O pedágio brasileiro é cobrado de forma proporcional ao número de eixos:
// cada eixo acima do 2º custa o mesmo valor unitário (tarifas[2] / 2).
// Por isso, em vez de só aceitar os eixos fixos que vieram cadastrados
// (2, 3, 5, 6), calculamos a tarifa para QUALQUER quantidade de eixos
// (4, 7, 9...) com base nesse valor unitário — assim veículos com 4 eixos
// (comuns na frota) não caem incorretamente na tarifa de 2 eixos.
function obterTarifaPedagio(pedagio, eixos = 2) {
  if (!pedagio || !pedagio.tarifas) return 0;
  const eixosNum = parseInt(eixos, 10);
  if (!eixosNum || eixosNum < 2) return pedagio.tarifas[2] || 0;
  // Se a tarifa exata pra esse número de eixos já estiver cadastrada, usa ela.
  if (pedagio.tarifas[eixosNum] != null) return pedagio.tarifas[eixosNum];
  // Caso contrário, deriva proporcionalmente a partir da tarifa de 2 eixos.
  const base2 = pedagio.tarifas[2];
  if (base2 == null) return 0;
  const tarifaUnitaria = base2 / 2;
  return +(tarifaUnitaria * eixosNum).toFixed(2);
}

// ─── Distância mínima entre um ponto e um SEGMENTO (não só as pontas) ──────
// Antes a detecção só comparava o pedágio aos 2 pontos extremos de cada
// trecho (origem/destino) — então um pedágio que fica no MEIO do caminho,
// mas longe dos dois extremos, nunca era visto (comum quando o endereço
// exato do cliente fica um pouco afastado da rodovia principal).
// Aqui amostramos pontos ao longo da linha reta entre p1 e p2 e pegamos a
// menor distância — captura pedágios "no caminho", não só nas pontas.
// A quantidade de amostras se adapta ao tamanho do segmento: quando a função
// recebe o traçado REAL (centenas de pontos já bem próximos uns dos outros,
// ~50-100m de distância), não faz sentido gastar amostras extras num segmento
// de 80m — 2 bastam. Quando recebe só as paradas (linha reta de dezenas de
// km, usada apenas como fallback antes do trajeto real carregar), amostra a
// cada 150m — mais fino do que antes, já que o raio de detecção caiu de 3km
// pra 300m e uma amostragem grossa poderia "pular" por cima do ponto exato
// onde a reta passa mais perto do pedágio.
function _distanciaPontoSegmento(lat1, lon1, lat2, lon2, latP, lonP) {
  const comprimentoKm = distanciaHaversine(lat1, lon1, lat2, lon2);
  const amostras = Math.max(2, Math.min(80, Math.ceil(comprimentoKm / 0.15)));
  let menor = Infinity;
  for (let i = 0; i <= amostras; i++) {
    const t = i / amostras;
    const lat = lat1 + (lat2 - lat1) * t;
    const lon = lon1 + (lon2 - lon1) * t;
    const d = distanciaHaversine(lat, lon, latP, lonP);
    if (d < menor) menor = d;
  }
  return menor;
}

// ─── Detecta pedágios próximos à rota ───────────────────────────────────────
// Raio de detecção configurável via `raioKm` (padrão 300m). Antes era 3 km
// fixo, o que gerava falsos positivos — bastava a rota passar em algum
// bairro nas redondezas do pedágio (sem necessariamente usar aquele trecho
// da rodovia) pra ele ser acusado. Com o traçado REAL vindo do OSRM (ver
// obterPontosRotaComCoords / _mvRoutePoints), um raio pequeno é o certo: se
// o trajeto de verdade passa pelo pedágio, os pontos da polyline ficam a
// poucos metros dele; se o trajeto vai por outra via (mesmo perto), a
// distância mínima sobe bem acima de 300m e o pedágio não é mais acusado à
// toa.
// IMPORTANTE: 300m só faz sentido quando `paradas` é o traçado real (polyline
// do OSRM). Para chamadas que usam só a LINHA RETA entre pontos (fallback
// antes do trajeto real carregar, ou relatórios que nunca buscam o trajeto
// real, como o relatório de custo de frete), 300m é curto demais — a reta
// entre origem/destino se afasta facilmente mais de 300m do ponto onde a
// estrada real (que curva) passa perto do pedágio, e o pedágio deixa de ser
// detectado. Quem só tem a linha reta deve passar um `raioKm` maior
// (ex.: 3) explicitamente.
// ATENÇÃO: se um pedágio real deixar de ser detectado com o traçado REAL, o
// mais provável é a coordenada cadastrada em PEDAGIOS_BR estar um pouco
// deslocada da posição exata da praça/pórtico — ajuste lat/lon da entrada em
// vez de aumentar esse raio de volta.
function detectarPedagiosNaRota(paradas, eixos = 2, raioKm = 0.3) {
  if (!paradas || paradas.length < 2) return [];
  
  const pedagiosEncontrados = [];
  const RAIO_DETECCAO_KM = raioKm;
  
  // Itera sobre todas as paradas consecutivas
  for (let i = 0; i < paradas.length - 1; i++) {
    const p1 = paradas[i];
    const p2 = paradas[i + 1];
    
    if (!p1 || !p2 || p1.lat == null || p1.lon == null || p2.lat == null || p2.lon == null) continue;
    
    // Para cada pedágio, verifica se está próximo ao TRAJETO (não só nas pontas)
    PEDAGIOS_BR.forEach((ped) => {
      const distMinima = _distanciaPontoSegmento(p1.lat, p1.lon, p2.lat, p2.lon, ped.lat, ped.lon);
      
      if (distMinima <= RAIO_DETECCAO_KM) {
        // Verifica se já não foi detectado
        const jaExiste = pedagiosEncontrados.find(
          (p) => p.nome === ped.nome && Math.abs(p.distanciaKm - distMinima) < 0.5
        );
        
        if (!jaExiste) {
          // Calcula valor com base em eixos
          const valor = obterTarifaPedagio(ped, eixos);
          pedagiosEncontrados.push({
            nome: ped.nome,
            lat: ped.lat,
            lon: ped.lon,
            tarifas: ped.tarifas,
            valor: valor,
            rodovia: ped.rodovia,
            regiao: ped.regiao,
            distanciaKm: distMinima,
            eixos: eixos,
            modelo: ped.modelo || 'praca',
          });
        }
      }
    });
  }
  
  // Remove duplicatas por nome e ordena por distância
  const unicos = Array.from(new Map(pedagiosEncontrados.map(p => [p.nome, p])).values());
  return unicos.sort((a, b) => a.distanciaKm - b.distanciaKm);
}

// ─── Calcula custo total de pedágios para uma viagem ───────────────────────
function calcularCustoPedagios(pedagios, quantidadeVeiculos = 1) {
  return pedagios.reduce((total, ped) => total + (ped.valor * quantidadeVeiculos), 0);
}

// ─── Gera HTML com alerta de pedágios ──────────────────────────────────────
function renderizarAlertaPedagios(pedagios, custoPedagios, eixos = 2) {
  if (!pedagios || pedagios.length === 0) return '';

  const temPraca = pedagios.some(p => (p.modelo || 'praca') === 'praca');
  const temFreeFlow = pedagios.some(p => p.modelo === 'free_flow');

  const listaPedagios = pedagios
    .map((ped) => {
      const tarifaAtual = obterTarifaPedagio(ped, eixos);
      const isFreeFlow = ped.modelo === 'free_flow';
      const tagModelo = isFreeFlow
        ? `<span style="background:#DBEAFE;color:#1E40AF;font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;margin-left:4px;">FREE FLOW</span>`
        : '';
      return `
        <div style="display:flex;align-items:center;justify-content:space-between;padding:6px 8px;background:rgba(220,38,38,0.08);border-radius:6px;font-size:11px;">
          <span>
            <strong>${ped.nome}</strong>${tagModelo} • ${ped.rodovia}
            <span style="color:#6B7280;font-size:10px;margin-left:4px;">(${ped.regiao} • ${eixos} eixos)</span>
          </span>
          <span style="color:#DC2626;font-weight:700;white-space:nowrap;">R$ ${tarifaAtual.toFixed(2)}</span>
        </div>
      `;
    })
    .join('');

  // Texto do rodapé muda conforme o que foi detectado: praça física exige
  // vale-pedágio antecipado; pórtico free flow é cobrado depois, sem vale físico.
  let rodape;
  if (temPraca && temFreeFlow) {
    rodape = `
      <div style="font-size:10px;color:#92400E;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;">Vale-Pedágio + Pórtico(s) Free Flow</div>
      <div style="font-size:13px;color:#DC2626;font-weight:700;margin-top:2px;">
        💰 R$ ${custoPedagios.toFixed(2)}
        <span style="font-size:11px;font-weight:400;color:#7F1D1D;">
          — Lance o Vale-Pedágio (praças físicas) no Sem Parar; os pórticos Free Flow são cobrados depois, por tag/placa
        </span>
      </div>`;
  } else if (temFreeFlow) {
    rodape = `
      <div style="font-size:10px;color:#1E40AF;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;">Pórtico(s) Free Flow — sem cabine</div>
      <div style="font-size:13px;color:#DC2626;font-weight:700;margin-top:2px;">
        💰 R$ ${custoPedagios.toFixed(2)}
        <span style="font-size:11px;font-weight:400;color:#7F1D1D;">
          — Cobrado automaticamente pela tag ou pela placa (não precisa comprar vale-pedágio pra isso, mas o valor deve entrar no custo da viagem)
        </span>
      </div>`;
  } else {
    rodape = `
      <div style="font-size:10px;color:#92400E;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;">Vale-Pedágio Necessário</div>
      <div style="font-size:13px;color:#DC2626;font-weight:700;margin-top:2px;">
        💰 R$ ${custoPedagios.toFixed(2)}
        <span style="font-size:11px;font-weight:400;color:#7F1D1D;">
          — Inclua Vale-Pedágio no Sistema Sem Parar
        </span>
      </div>`;
  }

  return `
    <div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:8px;padding:12px;margin-top:8px;">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
        <span style="font-size:16px;">⚠️</span>
        <span style="font-weight:700;color:#7F1D1D;font-size:12px;">PEDÁGIO DETECTADO NA ROTA</span>
      </div>
      <div style="display:grid;gap:4px;margin-bottom:8px;max-height:120px;overflow-y:auto;">
        ${listaPedagios}
      </div>
      <div style="padding:8px;background:#FEF9F3;border-radius:6px;border-left:3px solid #D97706;">
        ${rodape}
      </div>
    </div>
  `;
}

// ─── Marca pedágios no mapa com marcadores ─────────────────────────────────
function adicionarMarcadorPedagiosNoMapa(mapa, pedagios) {
  if (!mapa || !pedagios || pedagios.length === 0) return [];
  
  const marcadores = [];
  
  pedagios.forEach((ped) => {
    // Cria ícone customizado para pedágio
    const iconHtml = `
      <div style="
        width: 32px;
        height: 32px;
        background: #DC2626;
        border: 2px solid white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        cursor: pointer;
      ">💰</div>
    `;
    
    const popup = `
      <div style="font-size:12px;min-width:200px;">
        <div style="font-weight:700;color:#DC2626;margin-bottom:4px;">${ped.nome}</div>
        <div style="color:#666;font-size:11px;margin-bottom:8px;">
          <strong>Rodovia:</strong> ${ped.rodovia}<br/>
          <strong>Região:</strong> ${ped.regiao}
        </div>
        <div style="background:#FEF2F2;padding:6px;border-radius:4px;text-align:center;">
          <div style="font-size:10px;color:#666;margin-bottom:2px;">Tarifa (${ped.eixos} eixos)</div>
          <div style="font-size:14px;font-weight:700;color:#DC2626;">R$ ${ped.valor.toFixed(2)}</div>
        </div>
      </div>
    `;
    
    try {
      if (window.L && window.L.marker) {
        const marker = window.L.marker([ped.lat, ped.lon], {
          icon: window.L.divIcon({
            html: iconHtml,
            className: 'custom-icon-pedagio',
            iconSize: [32, 32],
            iconAnchor: [16, 16],
            popupAnchor: [0, -16],
          }),
        }).addTo(mapa).bindPopup(popup);
        marcadores.push(marker);
      }
    } catch (e) {
      console.warn('[Pedágios] Erro ao adicionar marcador:', e);
    }
  });
  
  return marcadores;
}

// ─── Exporta funções globais ───────────────────────────────────────────────
window.detectarPedagiosNaRota = detectarPedagiosNaRota;
window.calcularCustoPedagios = calcularCustoPedagios;
window.renderizarAlertaPedagios = renderizarAlertaPedagios;
window.adicionarMarcadorPedagiosNoMapa = adicionarMarcadorPedagiosNoMapa;
window.distanciaHaversine = distanciaHaversine;
window.obterTarifaPedagio = obterTarifaPedagio;
