// ═══════════════════════════════════════════════════════════════════════════
// NEXTA — MAPA DE VIAGEM + DASHBOARD DE OPERAÇÕES
// ═══════════════════════════════════════════════════════════════════════════
// Terceiro <script> não-module, carregado após Leaflet/html2canvas (tags
// <script src> imediatamente acima no documento). Continuação do mesmo
// escopo global do Roteirizador (script anterior) — usa variáveis e funções
// definidas ali (pedidos, veiculos, ultimoResultado, showToast, etc).
//
// ÍNDICE
//   MAPA DE VIAGEM .............. waypoints arrastáveis + exportação de imagem do mapa
//   DASHBOARD DE OPERAÇÕES — IIFE isolada, com seu próprio armazenamento:
//     Armazenamento ................. snapshots de roteirizações salvas em localStorage
//     Salvar roteirização atual ..... window.dashSalvarAtual()
//     Popular select de meses ....... window.dashPopularMeses()
//     Extrair dados agregados ....... agrega snapshots de um ou mais meses
//     Renderizar Dashboard .......... monta a tela com KPIs e gráficos
//     Gráfico de barras horizontal ... gráfico inline (sem libs externas)
//     Gráfico Km vs Volume ........... gráfico de dispersão
//     Gráfico de ocupação (canvas) .... gráfico de ocupação por viagem
//     Mapa Histórico .................. mapa Leaflet com rotas agregadas
//     Carregar por mês selecionado .... window.dashCarregarMes()
//     Carregar todos os períodos ...... window.dashCarregarTodos()
//     Hook: popular meses .............. dispara ao abrir a aba do dashboard
// ═══════════════════════════════════════════════════════════════════════════
/* ══════════════════════════════════════════════════════
   MAPA DE VIAGEM — waypoints arrastáveis + exportação
══════════════════════════════════════════════════════ */
// IMPORTANTE: var (não let/const). abrirModalViagem() — que inicia o mapa de
// viagem — está definida no <script> do Roteirizador, um escopo de script
// SEPARADO deste. Ela lê/escreve estas variáveis sem prefixo "window." (ex.:
// "_mvWaypoints = [...]"), o que só funciona se elas forem propriedades reais
// de window — daí "var" em vez de "let/const" aqui. Com let/const, essa escrita
// criava uma variável global *desconectada* (non-strict) ou lançava
// ReferenceError, e o mapa de viagem nunca via os dados que abrirModalViagem
// pensava ter preenchido — causa raiz de "traçado não aparece"/"km não atualiza".
var _mvWaypoints = [];       // [{lat,lon,marker,tipo,nome}] pontos originais (origem, paradas, retorno)
var _mvPolylines = [];       // polylines desenhadas (uma por segmento)
var _mvUserWaypoints = [];   // waypoints intermediários inseridos pelo usuário arrastando a linha
var _mvDistSpan = null;
var _mvDragMarker = null;    // marcador fantasma que aparece ao hover na linha
var _mvRenderToken = 0;      // token incremental — evita que uma chamada antiga de mvDesenharRota
                              // sobrescreva o resultado de uma chamada mais recente (race condition)
var _mvCanvasRenderer = null; // renderer Canvas dedicado para a polyline (evita bug de SVG no html2canvas)
// ─── State para salvar rota ajustada na viagem ────────────────────────────
// Salva os waypoints do usuário por chave "vId||iV" para persistir durante a sessão
var _mvRotasSalvas = {};
function _mvChaveViagem() {
  return `${window._mvVeiculoId || ''}||${window._mvIdxViagem ?? ''}`;
}
// ─── Garantir mapa ──────────────────────────────────────────────────────────
function garantirMapaViagem() {
  if (!mapaViagem) {
    mapaViagem = L.map('mv-mapa', { zoomControl: true });
    L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap &copy; CARTO',
      crossOrigin: true // permite que o html2canvas leia os tiles sem "tainted canvas" (CORS)
    }).addTo(mapaViagem);
    // Renderer Canvas dedicado para a linha da rota. O Leaflet, por padrão, desenha
    // polylines em SVG — e o html2canvas tem um bug documentado e conhecido onde o
    // SVG do Leaflet fica deslocado/cortado na imagem exportada (_updateSvgViewport
    // não é interpretado corretamente). Usar Canvas para esta camada evita o problema
    // por completo, já que <canvas> é capturado de forma direta e fiel pelo html2canvas.
    _mvCanvasRenderer = null; // não usado (polylines usam SVG padrão do Leaflet)
  }
  if (camadaViagem) camadaViagem.remove();
  camadaViagem = L.layerGroup().addTo(mapaViagem);
  _mvPolylines = [];
  _mvUserWaypoints = [];
  _mvWaypoints = [];
  const badge = document.getElementById('mv-vias-badge');
  if (badge) badge.style.display = 'none';
  setTimeout(() => mapaViagem.invalidateSize(), 50);
}
// ─── Desenhar marcador de parada (fixo, não arrastável) ────────────────────
function mvAdicionarWaypointArrastavel(p, idx) {
  const isOrigem  = p.tipo === 'origem';
  const isRetorno = p.tipo === 'retorno';
  const iconHtml = `<div style="
    width:24px;height:24px;border-radius:50%;
    background:${isOrigem ? '#111827' : isRetorno ? '#6B7280' : '#4F46E5'};
    color:#fff;font-size:11px;font-weight:700;
    display:flex;align-items:center;justify-content:center;
    border:2.5px solid #fff;
    box-shadow:0 2px 8px rgba(0,0,0,0.4);
    font-family:Inter,sans-serif;
    pointer-events:auto;
  ">${isOrigem ? '⬤' : isRetorno ? '↩' : String(idx)}</div>`;
  const icon = L.divIcon({ className: '', html: iconHtml, iconSize: [24,24], iconAnchor: [12,12] });
  // Pontos de parada NÃO são arrastáveis — apenas as vias intermediárias o são
  const marker = L.marker([p.lat, p.lon], { icon, draggable: false }).addTo(camadaViagem);
  if (p._popup) marker.bindPopup(p._popup, { maxWidth: 300 });
  _mvWaypoints[idx].marker = marker;
}
// ─── Desenhar rota completa com waypoints intermediários ───────────────────
// Monta a sequência: wp[0] → [vias entre 0 e 1] → wp[1] → [vias entre 1 e 2] → wp[2] ...
async function mvDesenharRota() {
  // Token desta execução — se outra chamada de mvDesenharRota for disparada
  // antes desta terminar (ex.: clique seguido de arraste rápido), a execução
  // mais antiga detecta que não é mais a "atual" e descarta seu resultado,
  // evitando que o km e o traçado fiquem presos num estado anterior.
  const myToken = ++_mvRenderToken;
  let obsoleta = false; // true = outra chamada mais recente assumiu — não recalcular km aqui
  try {
    // Monta sequência completa intercalando vias do usuário
    const sequencia = _mvMontarSequencia();
    const novasPolylines = [];
    for (let i = 0; i < sequencia.length - 1; i++) {
      try {
        const { coords } = await osrmFetchSegmento(sequencia[i], sequencia[i+1]);
        if (myToken !== _mvRenderToken) {
          novasPolylines.forEach(pl => { try { camadaViagem.removeLayer(pl); } catch(e){} });
          obsoleta = true;
          return; // execução obsoleta — abandona e limpa o que já tinha desenhado
        }
        const pl = L.polyline(coords, {
          color: '#00A499', weight: 5, opacity: 0.85,
          lineCap: 'round', lineJoin: 'round'
        }).addTo(camadaViagem);
        // Torna cada segmento clicável para inserir nova via
        _mvBindPolylineClick(pl, i, sequencia);
        novasPolylines.push(pl);
      } catch(e) {
        if (myToken !== _mvRenderToken) {
          novasPolylines.forEach(pl => { try { camadaViagem.removeLayer(pl); } catch(e2){} });
          obsoleta = true;
          return; // execução obsoleta — abandona e limpa o que já tinha desenhado
        }
        const pl = L.polyline(
          [[sequencia[i].lat, sequencia[i].lon],[sequencia[i+1].lat, sequencia[i+1].lon]],
          { color: '#00A499', weight: 4, opacity: 0.5, dashArray: '6,4' }
        ).addTo(camadaViagem);
        novasPolylines.push(pl);
      }
    }
    if (myToken !== _mvRenderToken) {
      // Outra chamada mais recente já está em andamento — descarta o que foi desenhado aqui
      novasPolylines.forEach(pl => { try { camadaViagem.removeLayer(pl); } catch(e){} });
      obsoleta = true;
      return;
    }
    // Só agora, com a certeza de que esta é a execução mais recente, remove as
    // polylines antigas e assume as novas.
    _mvPolylines.forEach(pl => { try { camadaViagem.removeLayer(pl); } catch(e){} });
    _mvPolylines = novasPolylines;
    // Re-desenhar marcadores de via por cima das polylines
    _mvUserWaypoints.forEach(uw => {
      if (uw.marker) { try { camadaViagem.removeLayer(uw.marker); } catch(e){} }
      _mvDesenharMarcadorVia(uw);
    });
    // Re-trazer marcadores de parada para frente
    _mvWaypoints.forEach(w => { if (w.marker) w.marker.bringToFront(); });
  } catch (e) {
    // Qualquer erro inesperado não previsto pelos try/catch internos acima
    // (ex.: falha ao desenhar um marcador) cai aqui. Não é re-lançado: graças
    // ao finally abaixo, o km ainda será recalculado com o que for possível,
    // em vez de deixar o badge "preso" sem nunca atualizar.
    console.error('mvDesenharRota: erro inesperado', e);
  } finally {
    // Recalcula o km SEMPRE que esta não for uma execução obsoleta — mesmo que
    // algo tenha falhado acima. Isso garante que o usuário sempre veja o km
    // atualizado (ou ao menos "calculando…" seguido de um valor), nunca um
    // badge ausente para sempre.
    if (!obsoleta) {
      await mvRecalcularDistancia(myToken);
    }
  }
}
// ─── Monta sequência intercalando vias entre waypoints fixos ──────────────
function _mvMontarSequencia() {
  const seq = [];
  for (let i = 0; i < _mvWaypoints.length; i++) {
    seq.push(_mvWaypoints[i]);
    // Adicionar vias do usuário que pertencem ao segmento i → i+1
    const viasSegmento = _mvUserWaypoints
      .filter(uw => uw.segmento === i)
      .sort((a, b) => a.ordem - b.ordem);
    viasSegmento.forEach(v => seq.push(v));
  }
  return seq;
}
// ─── Bind clique na polyline para inserir via ──────────────────────────────
function _mvBindPolylineClick(pl, segIdx, sequencia) {
  // Cursor crosshair ao hover
  pl.on('mouseover', () => { document.getElementById('mv-mapa').style.cursor = 'crosshair'; });
  pl.on('mouseout',  () => { document.getElementById('mv-mapa').style.cursor = 'default'; });
  pl.on('click', (e) => {
    L.DomEvent.stopPropagation(e);
    // Descobrir em qual segmento fixo estamos
    // segIdx é o índice na sequência completa; mapear para segmento fixo
    const segFixo = _mvMapearSegmentoFixo(segIdx, sequencia);
    const ordem = _mvUserWaypoints.filter(uw => uw.segmento === segFixo).length;
    const via = {
      lat: e.latlng.lat,
      lon: e.latlng.lng,
      segmento: segFixo,
      ordem,
      marker: null
    };
    _mvUserWaypoints.push(via);
    _mvAtualizarBadgeVias();
    mvDesenharRota();
  });
}
// ─── Mapear índice da sequência completa para segmento fixo ──────────────
function _mvMapearSegmentoFixo(seqIdx, sequencia) {
  // Conta quantos waypoints fixos passamos até seqIdx
  let fixos = 0;
  for (let i = 0; i <= seqIdx && i < sequencia.length; i++) {
    if (_mvWaypoints.includes(sequencia[i])) fixos++;
  }
  return Math.max(0, fixos - 1);
}
// ─── Desenhar marcador de via (arrastável, removível) ─────────────────────
function _mvDesenharMarcadorVia(via) {
  const iconHtml = `<div style="
    width:18px;height:18px;
    background:#4F46E5;border:2.5px solid #fff;
    border-radius:4px;transform:rotate(45deg);
    box-shadow:0 2px 6px rgba(0,0,0,0.4);
    cursor:grab;
  "></div>`;
  const icon = L.divIcon({ className: '', html: iconHtml, iconSize:[18,18], iconAnchor:[9,9] });
  const marker = L.marker([via.lat, via.lon], { icon, draggable: true, zIndexOffset: 1000 })
    .addTo(camadaViagem);
  marker.bindTooltip('Arraste para ajustar · Clique para remover', { direction:'top', offset:[0,-12] });
  marker.on('drag', (e) => {
    via.lat = e.latlng.lat;
    via.lon = e.latlng.lng;
  });
  marker.on('dragend', () => {
    mvDesenharRota();
  });
  marker.on('click', (e) => {
    L.DomEvent.stopPropagation(e);
    _mvUserWaypoints = _mvUserWaypoints.filter(uw => uw !== via);
    try { camadaViagem.removeLayer(marker); } catch(e2){}
    _mvAtualizarBadgeVias();
    mvDesenharRota();
  });
  via.marker = marker;
}
// ─── Badge de vias ─────────────────────────────────────────────────────────
function _mvAtualizarBadgeVias() {
  const badge = document.getElementById('mv-vias-badge');
  if (!badge) return;
  const n = _mvUserWaypoints.length;
  if (n === 0) { badge.style.display = 'none'; return; }
  badge.style.display = 'inline-block';
  badge.textContent = `${n} via(s)`;
}
// ─── Recalcular distância total ────────────────────────────────────────────
async function mvRecalcularDistancia(token = null) {
  const myToken = token != null ? token : ++_mvRenderToken;
  const el = document.getElementById('mv-dist-total');
  if (!el) return;
  el.style.display = 'inline-block';
  el.textContent = '📏 calculando…';
  const seq = _mvMontarSequencia();
  let totalKm = 0;
  for (let i = 0; i < seq.length - 1; i++) {
    try {
      const res = await osrmFetchSegmento(seq[i], seq[i+1]);
      totalKm += res.distKm;
    } catch(e) {}
  }
  if (myToken !== _mvRenderToken) return; // outra atualização mais recente já está em curso
  // Atualiza o texto exibido primeiro — é a parte mais importante para o usuário
  // e não deve depender do sucesso das etapas de persistência abaixo.
  el.textContent = `📏 ${totalKm.toFixed(1)} km`;
  try {
    // Salvar km e vias na viagem do resultado (para que o card de resultados atualize)
    const chave = _mvChaveViagem();
    _mvRotasSalvas[chave] = {
      userWaypoints: _mvUserWaypoints.map(uw => ({ lat: uw.lat, lon: uw.lon, segmento: uw.segmento, ordem: uw.ordem })),
      kmAjustado: totalKm
    };
    // Atualizar distanciaKm na viagem do ultimoResultado se estiver disponível
    const vId = window._mvVeiculoId;
    const iV  = window._mvIdxViagem;
    if (vId != null && iV != null && ultimoResultado?.[vId]?.[iV]) {
      ultimoResultado[vId][iV]._kmAjustado = totalKm;
      ultimoResultado[vId][iV]._userWaypoints = _mvRotasSalvas[chave].userWaypoints;
    }
  } catch (e) {
    console.error('mvRecalcularDistancia: erro ao persistir km/vias', e);
  }
}
// ─── Resetar para rota original ────────────────────────────────────────────
async function mvResetarRota() {
  _mvUserWaypoints.forEach(uw => { try { camadaViagem.removeLayer(uw.marker); } catch(e){} });
  _mvUserWaypoints = [];
  _mvAtualizarBadgeVias();
  // Limpar rota salva para esta viagem
  const chave = _mvChaveViagem();
  delete _mvRotasSalvas[chave];
  if (window._mvVeiculoId != null && window._mvIdxViagem != null && ultimoResultado?.[window._mvVeiculoId]?.[window._mvIdxViagem]) {
    delete ultimoResultado[window._mvVeiculoId][window._mvIdxViagem]._kmAjustado;
    delete ultimoResultado[window._mvVeiculoId][window._mvIdxViagem]._userWaypoints;
  }
  await mvDesenharRota();
}
// ─── Compartilhar via Google Maps ──────────────────────────────────────────
function mvCompartilharGoogleMaps() {
  const seq = _mvMontarSequencia();
  if (!seq || seq.length < 2) { showToast('Nenhuma rota para compartilhar.', false); return; }
  const origem  = seq[0];
  const destino = seq[seq.length - 1];
  const mids = seq.slice(1, -1);
  const waypointStr = mids.map(p => `${p.lat.toFixed(6)},${p.lon.toFixed(6)}`).join('|');
  const url = `https://www.google.com/maps/dir/?api=1`
    + `&origin=${origem.lat.toFixed(6)},${origem.lon.toFixed(6)}`
    + `&destination=${destino.lat.toFixed(6)},${destino.lon.toFixed(6)}`
    + `&travelmode=driving`
    + (waypointStr ? `&waypoints=${encodeURIComponent(waypointStr)}` : '');
  // Copiar para clipboard
  navigator.clipboard.writeText(url).then(() => {
    showToast('Link copiado! Cole no WhatsApp para enviar ao motorista ✓');
  }).catch(() => {
    window.open(url, '_blank');
  });
}
// ─── Exportar mapa ─────────────────────────────────────────────────────────
function mvExportarMapa() {
  if (!mapaViagem) return;
  const mapaEl = document.getElementById('mv-mapa');
  if (!mapaEl) return;
  if (typeof html2canvas === 'undefined') {
    mvCompartilharGoogleMaps();
    showToast('Para salvar como imagem, use a impressão do navegador (Ctrl+P → Salvar como PDF).', true);
    return;
  }
  showToast('Gerando imagem…', true);
  let respondeu = false;
  const watchdog = setTimeout(() => {
    if (!respondeu) {
      respondeu = true;
      showToast('Não foi possível gerar a imagem (tempo esgotado). Tente novamente ou use Ctrl+P → Salvar como PDF.', false);
    }
  }, 15000);
  // Esconde controles de zoom e atribuição para imagem limpa
  const zoomCtrl = mapaEl.querySelector('.leaflet-control-zoom');
  const zoomOrig = zoomCtrl ? zoomCtrl.style.display : null;
  if (zoomCtrl) zoomCtrl.style.display = 'none';
  // ── Estratégia: html2canvas captura tiles + marcadores (DOM/divIcon).
  //    A polyline SVG do Leaflet é problemática no html2canvas (SVG enorme
  //    com viewBox mal recalculado após pan). Solução: após capturar o DOM,
  //    redesenhar a polyline diretamente no canvas resultado usando a projeção
  //    geográfica do Leaflet — assim o traçado fica perfeitamente alinhado
  //    com os tiles capturados.
  html2canvas(mapaEl, {
    useCORS: true,
    allowTaint: false,
    backgroundColor: '#ffffff',
    scale: 2,
    logging: false,
    // No onclone: ocultar o SVG das polylines (será redesenhado manualmente)
    // e corrigir translate3d → translate nos panes de marcadores/tiles
    onclone: (clonedDoc) => {
      const clonedMapa = clonedDoc.getElementById('mv-mapa');
      if (!clonedMapa) return;
      // Oculta o pane SVG (polylines) — será re-desenhado manualmente
      const overlayPane = clonedMapa.querySelector('.leaflet-overlay-pane');
      if (overlayPane) overlayPane.style.display = 'none';
      // Corrige translate3d → translate em todos os panes (tiles, marcadores)
      // O leaflet-map-pane tem offset enorme — zeramos e compensamos nos filhos
      const mapPane = clonedMapa.querySelector('.leaflet-map-pane');
      function parseXY(t) {
        if (!t) return { x: 0, y: 0 };
        const m = t.match(/translate3d\(\s*([^,]+),\s*([^,]+)/) ||
                  t.match(/translate\(\s*([^,]+),\s*([^)]+)/);
        return m ? { x: parseFloat(m[1]) || 0, y: parseFloat(m[2]) || 0 } : { x: 0, y: 0 };
      }
      if (mapPane) {
        const off = parseXY(mapPane.style.transform);
        mapPane.style.transform = 'translate(0px,0px)';
        mapPane.style.webkitTransform = 'translate(0px,0px)';
        Array.from(mapPane.children).forEach(child => {
          if (child === overlayPane) return; // já oculto
          const c = parseXY(child.style.transform);
          const nx = c.x + off.x, ny = c.y + off.y;
          child.style.transform = `translate(${nx}px,${ny}px)`;
          child.style.webkitTransform = `translate(${nx}px,${ny}px)`;
          // Corrige translate3d residuais em netos (ex: tiles individuais)
          child.querySelectorAll('[style*="translate3d"]').forEach(el => {
            el.style.transform = (el.style.transform || '').replace(
              /translate3d\(([^,]+),([^,]+),[^)]+\)/, 'translate($1,$2)');
          });
        });
      } else {
        clonedMapa.querySelectorAll('[style*="translate3d"]').forEach(el => {
          el.style.transform = (el.style.transform || '').replace(
            /translate3d\(([^,]+),([^,]+),[^)]+\)/, 'translate($1,$2)');
        });
      }
    }
  }).then(canvas => {
    if (respondeu) return;
    respondeu = true;
    clearTimeout(watchdog);
    if (zoomCtrl) zoomCtrl.style.display = zoomOrig;
    try {
      const scale = 2; // deve coincidir com o scale do html2canvas
      const ctx = canvas.getContext('2d');
      const mapRect = mapaEl.getBoundingClientRect();
      // ── Redesenha cada polyline da rota usando projeção Leaflet → pixel ──
      // _mvPolylines contém os objetos L.polyline ativos no mapa.
      if (Array.isArray(_mvPolylines) && _mvPolylines.length > 0) {
        _mvPolylines.forEach(pl => {
          if (!pl || !pl.getLatLngs) return;
          let latlngs = pl.getLatLngs();
          // Achata arrays aninhados (multipolyline)
          if (latlngs.length && Array.isArray(latlngs[0])) latlngs = latlngs.flat();
          if (latlngs.length < 2) return;
          const opts = pl.options || {};
          const cor = opts.color || '#00A499';
          const peso = (opts.weight || 5) * scale;
          ctx.beginPath();
          ctx.strokeStyle = cor;
          ctx.lineWidth = peso;
          ctx.lineCap = opts.lineCap || 'round';
          ctx.lineJoin = opts.lineJoin || 'round';
          ctx.globalAlpha = opts.opacity !== undefined ? opts.opacity : 0.85;
          if (opts.dashArray) {
            const parts = String(opts.dashArray).split(/[\s,]+/).map(n => parseFloat(n) * scale);
            ctx.setLineDash(parts);
          } else {
            ctx.setLineDash([]);
          }
          latlngs.forEach((ll, i) => {
            // containerPointToLatLng inverso: latLngToContainerPoint
            const pt = mapaViagem.latLngToContainerPoint(ll);
            const px = pt.x * scale;
            const py = pt.y * scale;
            if (i === 0) ctx.moveTo(px, py);
            else ctx.lineTo(px, py);
          });
          ctx.stroke();
        });
        ctx.globalAlpha = 1;
        ctx.setLineDash([]);
      }
      const titulo = (document.getElementById('mv-titulo')?.textContent || 'Mapa_Viagem')
        .replace(/[^a-zA-Z0-9]/g, '_');
      const a = document.createElement('a');
      a.download = titulo + '.png';
      a.href = canvas.toDataURL('image/png');
      a.click();
      showToast('Mapa exportado! ✅');
    } catch (e2) {
      console.error('mvExportarMapa canvas draw:', e2);
      showToast('Não foi possível salvar a imagem (restrição do navegador). Use Ctrl+P → Salvar como PDF.', false);
    }
  }).catch(err => {
    if (respondeu) return;
    respondeu = true;
    clearTimeout(watchdog);
    if (zoomCtrl) zoomCtrl.style.display = zoomOrig;
    console.error('mvExportarMapa html2canvas:', err);
    showToast('Erro ao gerar imagem. Tente Ctrl+P → Salvar como PDF.', false);
  });
}
/* ── OSRM Road Routing ──────────────────────────────────────────
   Busca trajeto real por estradas via OSRM público (OpenStreetMap).
   Recebe array de {lat, lon}, devolve array de [lat, lng] do trajeto.
   Faz chamadas sequenciais par-a-par para suportar múltiplos waypoints.
─────────────────────────────────────────────────────────────────*/
/* ── OSRM Routing ────────────────────────────────────────────────
   Perfil: truck (caminhão) com fallback para car se indisponível.
   Retorna array de distâncias acumuladas por ponto (em km) via
   window._osrmDistancias[layer._leaflet_id] para uso nos popups.
──────────────────────────────────────────────────────────────── */
// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD DE OPERAÇÕES
// ═══════════════════════════════════════════════════════════════════════════
(function() {
// ── Armazenamento ──────────────────────────────────────────────────────────
// Snapshots guardados em localStorage: { [chaveYYYYMM]: [snapshot, ...] }
const DASH_KEY = 'nexta_dash_v1';
function dashGetStore() {
  try { return JSON.parse(localStorage.getItem(DASH_KEY) || '{}'); } catch { return {}; }
}
function dashSetStore(store) {
  try { localStorage.setItem(DASH_KEY, JSON.stringify(store)); } catch(e) {
    showToast('Armazenamento cheio. Remova períodos antigos.', false);
  }
}
function dashChave(snapshot) {
  const d = new Date(snapshot.savedAt || snapshot.datasEntrega?.[0] || Date.now());
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}`;
}
function dashChaveLabel(k) {
  const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return `${meses[parseInt(k.slice(4))-1]}/${k.slice(0,4)}`;
}
// ── Lê snapshots do histórico em disco (dirHandleHistorico) ─────────────────
async function dashLerHistoricoDisco() {
  console.log('[DASH] dashLerHistoricoDisco iniciado. dirHandleHistorico=', window.dirHandleHistorico);
  if (!window.dirHandleHistorico) {
    console.warn('[DASH] dirHandleHistorico não definido — pasta do histórico não selecionada.');
    return {};
  }
  let permOk = false;
  try { permOk = (await window.dirHandleHistorico.queryPermission({ mode: 'read' })) === 'granted'; } catch(e) { console.warn('[DASH] queryPermission erro:', e); }
  if (!permOk) {
    try { permOk = (await window.dirHandleHistorico.requestPermission({ mode: 'read' })) === 'granted'; } catch(e) { console.warn('[DASH] requestPermission erro:', e); }
  }
  console.log('[DASH] permissão leitura:', permOk);
  if (!permOk) return {};
  const store = {};
  let total = 0, aceitos = 0, rejeitados = 0;
  for await (const [name, handle] of window.dirHandleHistorico.entries()) {
    if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
    total++;
    try {
      const text = await (await handle.getFile()).text();
      const data = JSON.parse(text);
      console.log(`[DASH] arquivo: ${name} | versao=${data.versao} | savedAt=${data.savedAt} | temResultado=${!!data.resultado} | datasEntrega=`, data.datasEntrega);
      if (!data.versao || !data.savedAt) {
        console.warn(`[DASH] rejeitado (sem versao/savedAt): ${name}`);
        rejeitados++;
        continue;
      }
      const dataRef = (() => {
        const de = data.datasEntrega && data.datasEntrega[0];
        if (de) {
          const pts = de.split('/');
          if (pts.length === 3) {
            const d = new Date(parseInt(pts[2]), parseInt(pts[1])-1, parseInt(pts[0]));
            if (!isNaN(d.getTime())) return d;
          }
          const d2 = new Date(de);
          if (!isNaN(d2.getTime())) return d2;
        }
        return new Date(data.savedAt);
      })();
      const chave = `${dataRef.getFullYear()}${String(dataRef.getMonth()+1).padStart(2,'0')}`;
      console.log(`[DASH] aceito: ${name} → chave ${chave}`);
      aceitos++;
      const snap = {
        versao:       data.versao,
        savedAt:      data.savedAt,
        salvoPor:     data.salvoPor || 'Desconhecido',
        datasEntrega: data.datasEntrega || [],
        resumo:       data.resumo || {},
        pedidos:      data.pedidos   || [],
        terminais:    data.terminais || [],
        veiculos:     data.veiculos  || [],
        resultado:    data.resultado || {},
        controleTempo: data.controleTempo || {},
        _fonte:       'disco',
        _filename:    name,
      };
      if (!store[chave]) store[chave] = [];
      store[chave].push(snap);
    } catch(e) {
      console.error(`[DASH] erro ao ler ${name}:`, e);
      rejeitados++;
    }
  }
  console.log(`[DASH] total arquivos json: ${total} | aceitos: ${aceitos} | rejeitados: ${rejeitados}`);
  console.log('[DASH] chaves geradas:', Object.keys(store));
  return store;
}
// ── Mescla localStorage + disco, deduplicando por savedAt ────────────────
async function dashGetStoreMerged() {
  const local = dashGetStore();
  const disco  = await dashLerHistoricoDisco();
  const merged = {};
  // Indexa todos os savedAt já vindos do localStorage para deduplicar
  const vistosLocal = new Set();
  for (const snaps of Object.values(local)) {
    for (const s of snaps) if (s.savedAt) vistosLocal.add(s.savedAt);
  }
  // Copia localStorage
  for (const [k, snaps] of Object.entries(local)) {
    merged[k] = [...snaps];
  }
  // Adiciona do disco apenas se não já existe no localStorage (mesmo savedAt)
  for (const [k, snaps] of Object.entries(disco)) {
    if (!merged[k]) merged[k] = [];
    for (const s of snaps) {
      if (!s.savedAt || !vistosLocal.has(s.savedAt)) {
        merged[k].push(s);
      }
    }
  }
  // Ordena cada chave por data decrescente
  for (const k of Object.keys(merged)) {
    merged[k].sort((a, b) => (b.savedAt || '').localeCompare(a.savedAt || ''));
  }
  return merged;
}
// ── Salvar roteirização atual no Dashboard ─────────────────────────────────
window.dashSalvarAtual = function() {
  if (!ultimoResultado) { showToast('Execute a otimização primeiro.', false); return; }
  const snapshot = {
    versao: 1,
    savedAt: new Date().toISOString(),
    datasEntrega: [...new Set((pedidos||[]).map(p=>p.dataEntregaLogistica).filter(Boolean))],
    pedidos:    JSON.parse(JSON.stringify(pedidos||[])),
    terminais:  JSON.parse(JSON.stringify(terminaisCad||[])),
    veiculos:   JSON.parse(JSON.stringify(veiculos||[])),
    resultado:  JSON.parse(JSON.stringify(ultimoResultado||{})),
  };
  const chave = dashChave(snapshot);
  const store = dashGetStore();
  if (!store[chave]) store[chave] = [];
  store[chave].push(snapshot);
  dashSetStore(store);
  dashPopularMeses();
  showToast(`Roteirização salva no Dashboard — ${dashChaveLabel(chave)} ✓`);
};
// ── Popular select de meses ────────────────────────────────────────────────
window.dashPopularMeses = async function() {
  const sel = document.getElementById('dash-mes-sel');
  if (!sel) return;
  const store = await dashGetStoreMerged();
  const chaves = Object.keys(store).sort().reverse();
  sel.innerHTML = '<option value="">Selecionar</option>';
  chaves.forEach(k => {
    const opt = document.createElement('option');
    opt.value = k; opt.textContent = dashChaveLabel(k) + ` (${store[k].length} rot.)`;
    sel.appendChild(opt);
  });
};
// ── Extrair dados agregados de snapshots ───────────────────────────────────
// ── Chave de unificação de cliente ────────────────────────────────────────
// Prioridade: codigoSAP > nome normalizado (sem acentos, sem sufixo jurídico, maiúsculo)
function dashChaveCliente(ped) {
  const sap = (ped.codigoSAP || ped.codSAP || ped.sap || '').toString().trim();
  if (sap) return 'SAP:' + sap;
  // Fallback: normaliza o nome removendo acentos, sufixos jurídicos e espaços extras
  const nome = (ped.cliente || ped.nomeCliente || ped.nome || '?').toString();
  return 'NM:' + nome
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
    .toUpperCase()
    .replace(/\b(LTDA|EIRELI|ME|EPP|SA|S\.A\.|COMERCIO|COMERCIAL|INDUSTRIA|IND|COM)\b\.?/g, '')
    .replace(/[^A-Z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
// ── Nome canônico preferido (mais curto = menos abreviado) ─────────────────
function dashNomeCanônico(atual, novo) {
  if (!atual) return novo;
  // Prefere o nome mais longo (mais completo), sem truncamentos
  return novo.length > atual.length ? novo : atual;
}
function dashAgregar(snapshots) {
  const clientes = {};   // key=nome: {entregas, volume, km, lat, lon, cidade, capTotal}
  const viagens_ocup = []; // {label, ocup}
  let totalViagens = 0, totalKm = 0, totalVol = 0, totalCap = 0;
  // veiculos_escalados: lista de {id, snapIdx, capV, viagensIds} para cálculo de ocupação filtrada
  const veiculos_escalados = [];
  const rotasMap = [];   // para o mapa: [{termLat,termLon,paradas:[{lat,lon,nome}]}]
  snapshots.forEach((snap, snapIdx) => {
    const res = snap.resultado || {};
    const vecs = snap.veiculos || [];
    const terms = snap.terminais || [];
    vecs.forEach(v => {
      const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
      if (!viagens.length) return;
      // Capacidade do veículo escalado: conta 1x por veículo por snapshot
      const capV_esc = v.capacidade || v.capacidadeTotal || 0;
      if (capV_esc > 0) {
        totalCap += capV_esc;
        veiculos_escalados.push({ snapIdx, vid: v.id, capV: capV_esc });
      }
      // Terminal lat/lon
      const term = terms.find(t => t.nome === v.terminal);
      const tLat = term?.lat, tLon = term?.lon;
      const capV = v.capacidade || v.capacidadeTotal || 0;
      viagens.forEach((vi, iV) => {
        totalViagens++;
        const rotaPontos = { termLat: tLat, termLon: tLon, placa: v.placa, paradas: [] };
        let volViagem = 0;
        vi.paradas.forEach(par => {
          const ped = par.pedido || {};
          const nome = ped.cliente || ped.nomeCliente || par.nome || '?';
          const cidade = ped.cidade || '-';
          const vol = par.volumeTotal || 0;
          const coords = latLonEfetivo ? latLonEfetivo(ped) : { lat: par.lat, lon: par.lon };
          const lat = coords?.lat || par.lat;
          const lon = coords?.lon || par.lon;
          // Usa distanciaKm da parada (calculado na roteirização com fator de tortuosidade 1.3x)
          // em vez de Haversine direto terminal→cliente, que ignora a sequência de paradas.
          // Para clientes com múltiplas entregas, acumula e divide pela contagem → km médio.
          const km = par.distanciaKm > 0 ? par.distanciaKm
            : ((tLat && tLon && lat && lon) ? haversine(tLat, tLon, lat, lon) : 0);
          const chave = dashChaveCliente(ped);
          if (!clientes[chave]) clientes[chave] = { nome, cidade, entregas:0, volume:0, km:0, kmTotal:0, lat, lon, capTotal:0, viagensIds: new Set() };
          clientes[chave].nome = dashNomeCanônico(clientes[chave].nome, nome);
          clientes[chave].entregas++;
          clientes[chave].volume += vol;
          // capTotal por cliente: acumula a capacidade do veículo UMA VEZ por viagem
          // (não por parada), para que volume/capTotal dê a ocupação real dessa viagem.
          const _viagemId = v.id + '_' + iV;
          if (capV > 0 && !clientes[chave].viagensIds.has(_viagemId)) {
            clientes[chave].capTotal += capV;
            clientes[chave].viagensIds.add(_viagemId);
          }
          clientes[chave].kmTotal = (clientes[chave].kmTotal || 0) + km;
          clientes[chave].km = clientes[chave].kmTotal / clientes[chave].entregas; // km médio por entrega
          totalVol += vol;
          volViagem += vol;
          totalKm += km;
          rotaPontos.paradas.push({ lat, lon, nome, vol });
        });
        rotasMap.push(rotaPontos);
        if (capV > 0) {
          const ocup = Math.round((volViagem / capV) * 100);
          viagens_ocup.push({ label: `${v.placa} V${iV+1}`, ocup, snapIdx, vid: v.id, iV });
        }
      });
    });
  });
  const totalVec = totalCap > 0 ? Math.round((totalVol / totalCap) * 100) : 0;
  const totalEntregas = Object.values(clientes).reduce((s,c)=>s+c.entregas, 0);
  // Ocupação por cliente: volume do cliente / capacidade acumulada dos veículos que o atenderam
  const clientes_ocup = Object.values(clientes)
    .filter(c => c.capTotal > 0)
    .map(c => ({
      nome: c.nome,
      ocup: Math.min(100, Math.round((c.volume / c.capTotal) * 100)),
      volMedio: parseFloat((c.volume / c.entregas).toFixed(1))
    }))
    .sort((a, b) => b.ocup - a.ocup);
  return {
    clientes: Object.values(clientes).sort((a,b)=>b.volume-a.volume),
    viagens_ocup,
    clientes_ocup,
    veiculos_escalados, // [{snapIdx, vid, capV}] para cálculo de ocupação filtrada
    totalViagens,
    totalEntregas,
    totalVol: parseFloat(totalVol.toFixed(1)),
    totalKm: Math.round(totalKm),
    totalOcup: totalVec,
    totalClientes: Object.keys(clientes).length,
    rotasMap,
  };
}
// ── Renderizar Dashboard ───────────────────────────────────────────────────
// ── Filtro de clientes ────────────────────────────────────────────────────
let _dashClientesSelecionados = null; // null = todos; Set = filtro ativo
let _dashSnapshotsAtivos = [];        // snapshots atualmente carregados
let _dashTodosClientes   = [];        // lista completa de clientes do período

function dashToggleFiltroClientes() {
  const panel = document.getElementById('dash-cli-panel');
  if (!panel) return;
  const visible = panel.style.display !== 'none';
  panel.style.display = visible ? 'none' : 'flex';
  if (!visible) {
    document.getElementById('dash-cli-search').value = '';
    dashFiltrarListaClientes('');
  }
}

function _dashCheckSVG() {
  return '<svg width="11" height="9" viewBox="0 0 11 9"><polyline points="1,4.5 4,7.5 10,1" stroke="#000" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>';
}

function _dashAtualizarBoxVisual(box, checked) {
  box.style.borderColor = checked ? 'var(--pet-green,#b5e51d)' : '#bbb';
  box.style.background  = checked ? 'var(--pet-green,#b5e51d)' : 'transparent';
  box.innerHTML         = checked ? _dashCheckSVG() : '';
}

function dashPopularListaClientes() {
  const list = document.getElementById('dash-cli-list');
  if (!list) return;
  list.innerHTML = _dashTodosClientes.map(nome => {
    const sel     = !_dashClientesSelecionados || _dashClientesSelecionados.has(nome);
    const nomeSafe = nome.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
    const chk      = sel ? _dashCheckSVG() : '';
    const bc       = sel ? 'var(--pet-green,#b5e51d)' : '#bbb';
    const bg       = sel ? 'var(--pet-green,#b5e51d)' : 'transparent';
    return `<div data-cli="${nomeSafe}" data-checked="${sel ? '1' : '0'}"
      style="display:flex;align-items:center;gap:10px;padding:8px 14px;cursor:pointer;border-radius:6px;margin:0 4px;user-select:none;">
      <span class="dash-cb-box" style="flex-shrink:0;width:20px;height:20px;border-radius:5px;border:2px solid ${bc};background:${bg};display:flex;align-items:center;justify-content:center;transition:all .12s;">${chk}</span>
      <span style="font-size:12px;color:var(--text,#111);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;">${nome}</span>
    </div>`;
  }).join('');

  list.querySelectorAll('div[data-cli]').forEach(row => {
    row.addEventListener('mouseenter', () => row.style.background = 'rgba(0,0,0,0.04)');
    row.addEventListener('mouseleave', () => row.style.background = '');
    row.addEventListener('click', () => {
      const checked = row.dataset.checked !== '1';
      row.dataset.checked = checked ? '1' : '0';
      _dashAtualizarBoxVisual(row.querySelector('.dash-cb-box'), checked);
    });
  });
}

function dashSelecionarTodosClientesVisual(sel) {
  const list = document.getElementById('dash-cli-list');
  if (!list) return;
  list.querySelectorAll('div[data-cli]').forEach(row => {
    // Respeita busca ativa — só afeta itens visíveis
    if (row.style.display === 'none') return;
    row.dataset.checked = sel ? '1' : '0';
    _dashAtualizarBoxVisual(row.querySelector('.dash-cb-box'), sel);
  });
}

function dashFiltrarListaClientes(busca) {
  const list = document.getElementById('dash-cli-list');
  if (!list) return;
  const b = (busca || '').toLowerCase();
  list.querySelectorAll('div[data-cli]').forEach(row => {
    row.style.display = row.dataset.cli.toLowerCase().includes(b) ? '' : 'none';
  });
}

function dashSelecionarTodosClientes(sel) {
  dashSelecionarTodosClientesVisual(sel);
}

function dashAplicarFiltroClientes() {
  const list = document.getElementById('dash-cli-list');
  if (!list) return;
  const selecionados = new Set();
  // dataset.cli pode ter entidades HTML escapadas (&amp; &quot;) — decodifica antes de comparar com c.nome
  const _dec = s => s.replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>');
  list.querySelectorAll('div[data-checked="1"]').forEach(row => selecionados.add(_dec(row.dataset.cli)));
  // Se todos selecionados = sem filtro ativo
  _dashClientesSelecionados = selecionados.size === _dashTodosClientes.length ? null : selecionados;
  // Atualiza badge
  const badge = document.getElementById('dash-cli-badge');
  if (badge) {
    if (_dashClientesSelecionados) {
      badge.textContent = selecionados.size;
      badge.style.display = '';
    } else {
      badge.style.display = 'none';
    }
  }
  document.getElementById('dash-cli-panel').style.display = 'none';
  dashRenderComFiltro();
}

function dashRenderComFiltro() {
  dashRender(_dashSnapshotsAtivos);
}

// Fecha painel ao clicar fora
document.addEventListener('click', function(e) {
  const panel = document.getElementById('dash-cli-panel');
  const btn   = document.getElementById('dash-cli-btn');
  if (panel && panel.style.display !== 'none' && !panel.contains(e.target) && !btn?.contains(e.target)) {
    panel.style.display = 'none';
  }
});

window.dashToggleFiltroClientes   = dashToggleFiltroClientes;
window.dashFiltrarListaClientes    = dashFiltrarListaClientes;
window.dashSelecionarTodosClientes = dashSelecionarTodosClientes;
window.dashAplicarFiltroClientes   = dashAplicarFiltroClientes;

function dashRender(snapshots) {
  _dashSnapshotsAtivos = snapshots || [];
  if (!snapshots || !snapshots.length) {
    document.querySelectorAll('#dk-viagens,#dk-entregas,#dk-volume,#dk-ocup,#dk-km,#dk-clientes')
      .forEach(el => { if(el) el.textContent = '-'; });
    document.getElementById('dash-tabela-cli-body').innerHTML =
      '<tr><td colspan="6" style="color:var(--text-3);text-align:center;padding:32px;">Nenhum dado para este período</td></tr>';
    return;
  }
  const d = dashAgregar(snapshots);
  // Atualiza lista global de clientes para o filtro
  // Só reinicia a lista visual se não houver filtro ativo (evita resetar seleção do usuário)
  const _novaListaClientes = d.clientes.map(c => c.nome).sort();
  const _listaIgual = _novaListaClientes.length === _dashTodosClientes.length &&
    _novaListaClientes.every((n,i) => n === _dashTodosClientes[i]);
  if (!_listaIgual) {
    _dashTodosClientes = _novaListaClientes;
    // Se havia filtro ativo, mantém apenas os clientes que ainda existem
    if (_dashClientesSelecionados) {
      const novosNomes = new Set(_dashTodosClientes);
      const filtroAtualizado = new Set([..._dashClientesSelecionados].filter(n => novosNomes.has(n)));
      _dashClientesSelecionados = filtroAtualizado.size === _dashTodosClientes.length ? null : filtroAtualizado;
    }
  }
  dashPopularListaClientes();
  // Aplica filtro se ativo
  const clientesFiltrados = _dashClientesSelecionados
    ? d.clientes.filter(c => _dashClientesSelecionados.has(c.nome))
    : d.clientes;
  const ocupFiltrados = _dashClientesSelecionados
    ? d.clientes_ocup.filter(c => _dashClientesSelecionados.has(c.nome))
    : d.clientes_ocup;
  // KPIs — calculados sobre os clientes filtrados (respeita filtro de cliente ativo)
  const set = (id, v) => { const el=document.getElementById(id); if(el) el.textContent=v; };
  // Recalcula totais a partir dos clientes filtrados
  const _kpiEntregas = clientesFiltrados.reduce((s,c) => s+c.entregas, 0);
  const _kpiVol      = clientesFiltrados.reduce((s,c) => s+c.volume,   0);
  const _kpiKm       = clientesFiltrados.reduce((s,c) => s+(c.kmTotal||c.km*c.entregas||0), 0);
  // Ocupação = volume total de pedidos / capacidade total da frota escalada.
  // Sem filtro: d.totalOcup (calculado em dashAgregar: totalVol / totalCap por veículo).
  // Com filtro de cliente: volume dos clientes filtrados / cap dos veículos que os atenderam.
  let _kpiOcup = d.totalOcup;
  if (_dashClientesSelecionados) {
    let _filtVol = 0, _filtCap = 0;
    const _nomesF = _dashClientesSelecionados;
    const _veicsAtenderam = new Set(); // evita contar o mesmo veículo duas vezes
    _dashSnapshotsAtivos.forEach((snap, sIdx) => {
      const res = snap.resultado || {}, vecs = snap.veiculos || [];
      vecs.forEach(v => {
        const capV = v.capacidade || v.capacidadeTotal || 0;
        const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
        const atendeCliente = viagens.some(vi =>
          vi.paradas.some(par => {
            const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
            return _nomesF.has(n);
          })
        );
        if (!atendeCliente) return;
        // Capacidade: conta 1x por veículo
        const _vKey = sIdx + '_' + v.id;
        if (capV > 0 && !_veicsAtenderam.has(_vKey)) {
          _filtCap += capV;
          _veicsAtenderam.add(_vKey);
        }
        // Volume: apenas paradas dos clientes filtrados
        viagens.forEach(vi => vi.paradas.forEach(par => {
          const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
          if (_nomesF.has(n)) _filtVol += par.volumeTotal || 0;
        }));
      });
    });
    _kpiOcup = _filtCap > 0 ? Math.round((_filtVol / _filtCap) * 100) : d.totalOcup;
  }
  // Viagens: conta apenas viagens que atendem ao menos um cliente filtrado
  const _nomesFilter = _dashClientesSelecionados;
  let _kpiViagens = d.totalViagens;
  if (_nomesFilter) {
    _kpiViagens = 0;
    _dashSnapshotsAtivos.forEach(snap => {
      const res  = snap.resultado || {};
      const vecs = snap.veiculos  || [];
      vecs.forEach(v => {
        (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length).forEach(vi => {
          const temCliente = vi.paradas.some(par => {
            const nome = (par.pedido||{}).cliente || (par.pedido||{}).nomeCliente || par.nome || '';
            return _nomesFilter.has(nome);
          });
          if (temCliente) _kpiViagens++;
        });
      });
    });
  }
  set('dk-viagens',  _kpiViagens);
  set('dk-entregas', _kpiEntregas);
  set('dk-volume',   parseFloat(_kpiVol.toFixed(1)) + ' m³');
  set('dk-ocup',     _kpiOcup + '%');
  set('dk-km',       Math.round(_kpiKm).toLocaleString('pt-BR') + ' km');
  set('dk-clientes', clientesFiltrados.length);
  // Gráfico de barras: volume por cliente
  dashBarChart('dash-chart-vol', clientesFiltrados.slice(0,30), c=>c.volume.toFixed(1),
    '#f0be40', 'm³', c=>c.nome);
  // Gráfico de barras: entregas por cliente
  dashBarChart('dash-chart-ent', clientesFiltrados.slice(0,30).sort((a,b)=>b.entregas-a.entregas),
    c=>c.entregas, '#70a8f0', 'ent.', c=>c.nome);
  // Gráfico Km vs Volume
  dashKmVolChart('dash-chart-km', clientesFiltrados.slice(0,30));
  // Gráfico de ocupação por cliente
  dashOcupClienteChart('dash-chart-ocup', ocupFiltrados.slice(0,30));
  // Mapa
  dashRenderMapa(d.rotasMap);
  // Tabela
  const tbody = document.getElementById('dash-tabela-cli-body');
  if (tbody) {
    tbody.innerHTML = clientesFiltrados.map((c, i) => {
      const kmVol = c.km > 0 ? (c.volume / c.km).toFixed(2) : '-';
      const bg = i % 2 === 0 ? '#FFFFFF' : '#F9FAFB';
      return `<tr style="background:${bg};border-bottom:1px solid #E5E7EB;">
        <td style="padding:10px 14px;font-size:13px;font-weight:500;color:#111827;">${c.nome}</td>
        <td style="padding:10px 14px;font-size:12px;color:#6B7280;">${c.cidade}</td>
        <td style="padding:10px 14px;font-size:13px;color:#111827;text-align:center;">${c.entregas}</td>
        <td style="padding:10px 14px;font-size:13px;color:#111827;text-align:center;">${c.volume.toFixed(1)}</td>
        <td style="padding:10px 14px;font-size:13px;color:#111827;text-align:center;">${c.km > 0 ? c.km.toFixed(0)+' km' : '-'}</td>
        <td style="padding:10px 14px;font-size:13px;color:#111827;text-align:center;">${kmVol !== '-' ? kmVol+' m³/km' : '-'}</td>
      </tr>`;
    }).join('');
  }
}
// ── Gráfico de barras horizontal inline ───────────────────────────────────
function dashBarChart(containerId, itens, valFn, cor, sufixo, labelFn) {
  const el = document.getElementById(containerId);
  if (!el) return;
  if (!itens.length) { el.innerHTML = '<div style="color:var(--text-3);font-size:12px;padding:12px">Sem dados</div>'; return; }
  const maxV = Math.max(...itens.map(i => parseFloat(valFn(i)) || 0), 1);
  el.innerHTML = itens.map(item => {
    const v = parseFloat(valFn(item)) || 0;
    const pct = Math.round((v / maxV) * 100);
    const label = labelFn(item);
    const display = Number.isInteger(v) ? v : parseFloat(v).toFixed(1);
    return `<div style="margin-bottom:12px;">
      <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;">
        <span style="font-size:12px;font-weight:700;color:#000000;" title="${label}">${label}</span>
        <span style="font-size:13px;font-weight:700;color:#000000;white-space:nowrap;margin-left:10px;">${display} ${sufixo}</span>
      </div>
      <div style="height:10px;background:rgba(0,0,0,0.08);border-radius:99px;overflow:hidden;">
        <div style="width:${pct}%;height:100%;background:${cor};border-radius:99px;transition:width .5s ease;"></div>
      </div>
    </div>`;
  }).join('');
}
// ── Gráfico Km vs Volume ───────────────────────────────────────────────────
function dashKmVolChart(containerId, clientes) {
  const el = document.getElementById(containerId);
  if (!el) return;
  // Mostra só clientes com km calculado, ordenados por km decrescente
  const itens = clientes.filter(c => c.km > 0).sort((a, b) => b.km - a.km).slice(0, 30);
  if (!itens.length) { el.innerHTML = '<div style="color:#888;font-size:12px;padding:12px">Sem dados de distância — abra o mapa de cada viagem para calcular.</div>'; return; }
  const maxKm  = Math.max(...itens.map(c => c.km  || 0), 1);
  const maxVol = Math.max(...itens.map(c => c.volume || 0), 1);
  el.innerHTML = itens.map(c => {
    const pctKm  = c.km > 0 ? Math.round((c.km  / maxKm)  * 100) : 0;
    const pctVol = Math.round((c.volume / maxVol) * 100);
    const kmLabel = c.km > 0 ? c.km.toFixed(0) + ' km' : '— km';
    return `<div style="margin-bottom:12px;">
      <div style="font-size:12px;font-weight:700;color:#000000;margin-bottom:5px;">${c.nome}</div>
      <div style="display:flex;gap:4px;align-items:center;">
        <span style="font-size:10px;font-weight:700;color:#000000;width:22px;text-align:right;">km</span>
        <div style="flex:1;height:8px;background:rgba(0,0,0,0.08);border-radius:99px;overflow:hidden;">
          <div style="width:${pctKm}%;height:100%;background:#6ee04a;border-radius:99px;"></div>
        </div>
        <span style="font-size:10px;font-weight:700;color:#000000;width:52px;">${kmLabel}</span>
      </div>
      <div style="display:flex;gap:4px;align-items:center;margin-top:4px;">
        <span style="font-size:10px;font-weight:700;color:#000000;width:22px;text-align:right;">m³</span>
        <div style="flex:1;height:8px;background:rgba(0,0,0,0.08);border-radius:99px;overflow:hidden;">
          <div style="width:${pctVol}%;height:100%;background:#f0be40;border-radius:99px;"></div>
        </div>
        <span style="font-size:10px;font-weight:700;color:#000000;width:52px;">${c.volume.toFixed(1)} m³</span>
      </div>
    </div>`;
  }).join('');
}
// ── Gráfico de ocupação por cliente (barras horizontais HTML) ──────────────
function dashOcupClienteChart(containerId, itens) {
  const el = document.getElementById(containerId);
  if (!el) return;
  if (!itens.length) {
    el.innerHTML = '<div style="color:#888;font-size:12px;padding:12px">Sem dados de ocupação</div>';
    return;
  }
  el.innerHTML = itens.map(function(item) {
    const pct = Math.min(Math.round(item.ocup), 100);
    const cor = pct >= 90 ? '#4caf50' : pct >= 60 ? '#f0be40' : '#f06060';
    const vol = item.volMedio != null ? item.volMedio + ' m³/ent.' : '';
    return '<div style="margin-bottom:14px;">'
      + '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px;gap:8px;">'
      + '<span style="font-size:12px;font-weight:700;color:var(--text,#111);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + item.nome + '</span>'
      + '<span style="display:flex;gap:12px;align-items:center;white-space:nowrap;flex-shrink:0;">'
      + (vol ? '<span style="font-size:11px;font-weight:500;color:var(--text-3,#777);">' + vol + '</span>' : '')
      + '<span style="font-size:13px;font-weight:800;color:' + cor + ';min-width:36px;text-align:right;">' + pct + '%</span>'
      + '</span>'
      + '</div>'
      + '<div style="position:relative;height:14px;background:rgba(0,0,0,0.07);border-radius:99px;overflow:hidden;">'
      + '<div style="width:' + pct + '%;height:100%;background:' + cor + ';border-radius:99px;transition:width .6s ease;"></div>'
      + '</div>'
      + '</div>';
  }).join('');
}
// ── Gráfico de ocupação por viagem (canvas) — mantido para compatibilidade ──
function dashOcupChart(canvasId, itens) {
  dashOcupClienteChart(canvasId, itens.map(function(i){ return { nome: i.label, ocup: i.ocup }; }));
}
// ── Mapa Histórico ─────────────────────────────────────────────────────────
let _dashMap = null;
let _dashMapLayers = [];
function dashRenderMapa(rotasMap) {
  if (typeof L === 'undefined') return;
  const el = document.getElementById('dash-mapa');
  if (!el) return;
  // Inicializar mapa uma só vez
  if (!_dashMap) {
    _dashMap = L.map('dash-mapa', { zoomControl: true, attributionControl: false });
    L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
      maxZoom: 18,
      attribution: '&copy; OpenStreetMap &copy; CARTO'
    }).addTo(_dashMap);
  }
  // Limpar camadas anteriores
  _dashMapLayers.forEach(l => _dashMap.removeLayer(l));
  _dashMapLayers = [];
  if (!rotasMap.length) return;
  const bounds = [];
  const terminaisVistos = new Set();
  const clientesVistos = {};
  rotasMap.forEach((rota, ri) => {
    const cor = `hsl(${(ri * 47) % 360},70%,55%)`;
    // Terminal
    if (rota.termLat && rota.termLon) {
      const tKey = `${rota.termLat.toFixed(4)},${rota.termLon.toFixed(4)}`;
      if (!terminaisVistos.has(tKey)) {
        terminaisVistos.add(tKey);
        const icon = L.divIcon({
          className: '',
          html: `<div style="width:14px;height:14px;background:#00A499;border:2px solid #fff;border-radius:50%;box-shadow:0 0 6px #00A499;"></div>`,
          iconSize: [14,14], iconAnchor: [7,7]
        });
        const m = L.marker([rota.termLat, rota.termLon], { icon })
          .bindPopup(`<b>Terminal</b><br><small>${rota.placa}</small>`);
        m.addTo(_dashMap);
        _dashMapLayers.push(m);
        bounds.push([rota.termLat, rota.termLon]);
      }
    }
    // Paradas e linhas
    let prev = rota.termLat ? [rota.termLat, rota.termLon] : null;
    rota.paradas.forEach(par => {
      if (!par.lat || !par.lon) return;
      bounds.push([par.lat, par.lon]);
      // Linha do terminal até o cliente
      if (prev) {
        const line = L.polyline([prev, [par.lat, par.lon]], {
          color: cor, weight: 1.5, opacity: 0.45, dashArray: '4 4'
        }).addTo(_dashMap);
        _dashMapLayers.push(line);
      }
      prev = [par.lat, par.lon];
      // Marcador cliente (agrupa múltiplas visitas)
      const cKey = `${par.lat.toFixed(4)},${par.lon.toFixed(4)}`;
      if (!clientesVistos[cKey]) {
        clientesVistos[cKey] = { lat: par.lat, lon: par.lon, nome: par.nome, visitas: 0, vol: 0 };
      }
      clientesVistos[cKey].visitas++;
      clientesVistos[cKey].vol += par.vol || 0;
    });
  });
  // Renderizar marcadores de clientes
  Object.values(clientesVistos).forEach(c => {
    const r = Math.min(12, 6 + c.visitas * 1.5);
    const icon = L.divIcon({
      className: '',
      html: `<div style="width:${r*2}px;height:${r*2}px;background:#4F46E5;border:1.5px solid rgba(255,255,255,0.6);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:8px;color:#fff;font-weight:700;">${c.visitas>1?c.visitas:''}</div>`,
      iconSize: [r*2, r*2], iconAnchor: [r, r]
    });
    const m = L.marker([c.lat, c.lon], { icon })
      .bindPopup(`<b>${c.nome}</b><br>${c.visitas} visita(s)<br>${c.vol.toFixed(1)} m³`);
    m.addTo(_dashMap);
    _dashMapLayers.push(m);
  });
  // Ajustar zoom
  if (bounds.length) {
    try { _dashMap.fitBounds(bounds, { padding: [30, 30], maxZoom: 12 }); } catch(e) {}
  }
  // Forçar resize após render
  setTimeout(() => _dashMap && _dashMap.invalidateSize(), 200);
}
// ── Sincronizar: relê o histórico em disco e atualiza tudo ────────────────
window.dashSincronizar = async function() {
  const btn = document.getElementById('dash-sync-btn');
  if (btn) { btn.textContent = '⏳ Sincronizando...'; btn.disabled = true; }
  try {
    // Se a pasta do histórico ainda não foi selecionada, pede ao usuário
    if (!window.dirHandleHistorico) {
      if (!window.showDirectoryPicker) {
        showToast('Seu navegador não suporta acesso a pastas. Use Chrome ou Edge.', false);
        return;
      }
      try {
        const handle = await window.showDirectoryPicker({ mode: 'read' });
        window.dirHandleHistorico = handle;
        // Persiste no mesmo IndexedDB usado pela aba "Histórico" — assim a
        // pasta escolhida aqui também é lembrada em sessões futuras (sujeito
        // à reconfirmação de permissão do navegador).
        if (typeof _histSalvarHandle === 'function') {
          try { await _histSalvarHandle(handle); } catch(e) {}
        }
      } catch(e) {
        if (e.name !== 'AbortError') showToast('Pasta não selecionada.', false);
        return;
      }
    }
    await dashPopularMeses();
    const sel = document.getElementById('dash-mes-sel');
    if (sel && sel.value) {
      await window.dashCarregarMes(sel.value);
    } else {
      await window.dashCarregarTodos();
    }
    showToast('Dashboard sincronizado com o histórico ✅');
  } catch(e) {
    showToast('Erro ao sincronizar: ' + e.message, false);
  } finally {
    if (btn) { btn.textContent = '🔄 Sincronizar'; btn.disabled = false; }
  }
};
// ── Carregar por mês selecionado ───────────────────────────────────────────
window.dashCarregarMes = async function(chave) {
  if (!chave) return;
  const store = await dashGetStoreMerged();
  dashRender(store[chave] || []);
};
// ── Carregar todos os períodos ─────────────────────────────────────────────
window.dashCarregarTodos = async function() {
  const store = await dashGetStoreMerged();
  const todos = Object.values(store).flat();
  dashRender(todos);
};
// ── Hook: popular meses quando abre a aba ─────────────────────────────────
const _origShowTab = window.showTab;
window.showTab = function(tab) {
  if (_origShowTab) _origShowTab(tab);
  if (tab === 'dashboard_rot') {
    dashPopularMeses();
    // Invalidar mapa se já existir
    setTimeout(() => { if (_dashMap) _dashMap.invalidateSize(); }, 300);
  }
};
// ── Refresh do dashboard ao excluir entrada do histórico ──────────────────
// Intercepta excluirEntradaHistorico para re-popular meses e re-renderizar
const _origExcluir = window.excluirEntradaHistorico;
window.excluirEntradaHistorico = async function(filename, btn) {
  if (_origExcluir) await _origExcluir(filename, btn);
  // Atualiza o select de meses com os dados restantes
  await dashPopularMeses();
  // Se há dados carregados no dashboard, recarrega automaticamente
  const selMes = document.getElementById('dash-mes-sel');
  if (selMes && selMes.value) {
    await window.dashCarregarMes(selMes.value);
  }
};
// Popular na inicialização — aguarda primeiro a restauração do handle salvo
// no IndexedDB (window.recuperarHandleHistoricoPromise, definida no script do
// Roteirizador) para não tentar ler o disco antes do handle estar disponível.
(async () => {
  try {
    if (window.recuperarHandleHistoricoPromise) {
      await window.recuperarHandleHistoricoPromise;
    }
  } catch (e) { /* segue mesmo se a restauração falhar */ }
  dashPopularMeses();

// ═══════════════════════════════════════════════════════════════════════════
// EXPORTAÇÃO EXCEL — Dashboard
// Gera um .xlsx fiel ao dashboard: KPIs, tabela por cliente, detalhamento
// por viagem/parada e ocupação por veículo — usando SheetJS (já carregado).
// ═══════════════════════════════════════════════════════════════════════════
window.dashExportarExcel = async function dashExportarExcel() {
  const snapshots = _dashSnapshotsAtivos;
  if (!snapshots || !snapshots.length) {
    alert('Nenhum dado carregado. Selecione um mês ou clique em "Todos os períodos" primeiro.');
    return;
  }
  // Respeita filtro de clientes ativo
  const filtroAtivo = _dashClientesSelecionados; // null = todos, Set = filtro

  // Verifica SheetJS
  if (typeof XLSX === 'undefined') {
    alert('Biblioteca SheetJS não encontrada. Verifique a importação no index.html.');
    return;
  }

  const btn = document.getElementById('dash-export-btn');
  const orig = btn?.textContent;
  if (btn) { btn.textContent = '⏳ Gerando...'; btn.disabled = true; }

  try {
    const d = dashAgregar(snapshots);
    const wb = XLSX.utils.book_new();

    // ── Helpers ─────────────────────────────────────────────────────────────
    const pct = v => (typeof v === 'number' ? v + '%' : v);
    const num = (v, dec=1) => typeof v === 'number' ? parseFloat(v.toFixed(dec)) : v;

    // Estilo de cabeçalho compartilhado (SheetJS Community não suporta estilos,
    // mas estruturamos para que apps Pro / xlsx-style possam aplicar facilmente)
    function addSheet(name, rows) {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // Largura automática por coluna
      const colWidths = rows.reduce((acc, row) => {
        row.forEach((cell, i) => {
          const len = cell !== null && cell !== undefined ? String(cell).length : 0;
          acc[i] = Math.max(acc[i] || 8, Math.min(len + 2, 60));
        });
        return acc;
      }, []);
      ws['!cols'] = colWidths.map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    }

    // ── Aba 1: Resumo KPIs ──────────────────────────────────────────────────
    // Aplica filtro de clientes aos dados agregados
    const cliExport = filtroAtivo
      ? d.clientes.filter(c => filtroAtivo.has(c.nome))
      : d.clientes;
    const ocupExport = filtroAtivo
      ? d.clientes_ocup.filter(c => filtroAtivo.has(c.nome))
      : d.clientes_ocup;
    // Recalcula KPIs filtrados para o Excel
    const _exVol      = cliExport.reduce((s,c) => s+c.volume, 0);
    const _exKm       = cliExport.reduce((s,c) => s+(c.kmTotal||c.km*c.entregas||0), 0);
    const _exEntregas = cliExport.reduce((s,c) => s+c.entregas, 0);
    // Ocupação do export = volume pedidos filtrados / cap frota que atendeu esses clientes
    let _exOcup = d.totalOcup;
    if (filtroAtivo) {
      let _exFiltVol = 0, _exFiltCap = 0;
      const _exVeicsAtenderam = new Set();
      snapshots.forEach((snap, sIdx) => {
        const res = snap.resultado || {}, vecs = snap.veiculos || [];
        vecs.forEach(v => {
          const capV = v.capacidade || v.capacidadeTotal || 0;
          const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
          const atende = viagens.some(vi =>
            vi.paradas.some(par => {
              const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
              return filtroAtivo.has(n);
            })
          );
          if (!atende) return;
          const _vKey = sIdx + '_' + v.id;
          if (capV > 0 && !_exVeicsAtenderam.has(_vKey)) {
            _exFiltCap += capV;
            _exVeicsAtenderam.add(_vKey);
          }
          viagens.forEach(vi => vi.paradas.forEach(par => {
            const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
            if (filtroAtivo.has(n)) _exFiltVol += par.volumeTotal || 0;
          }));
        });
      });
      _exOcup = _exFiltCap > 0 ? Math.round((_exFiltVol / _exFiltCap) * 100) : d.totalOcup;
    }
    let _exViagens = d.totalViagens;
    if (filtroAtivo) {
      _exViagens = 0;
      snapshots.forEach(snap => {
        const res = snap.resultado||{}; const vecs = snap.veiculos||[];
        vecs.forEach(v => {
          (res[v.id]||[]).filter(vi=>!vi._vazio&&(vi.paradas||[]).length).forEach(vi => {
            if (vi.paradas.some(par => {
              const n=(par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
              return filtroAtivo.has(n);
            })) _exViagens++;
          });
        });
      });
    }
    const filtroLabel = filtroAtivo
      ? `Clientes filtrados: ${[...filtroAtivo].join(', ')}`
      : 'Todos os clientes';

    addSheet('Resumo', [
      ['DASHBOARD NEXTA — RESUMO DO PERÍODO'],
      ['Gerado em', new Date().toLocaleString('pt-BR')],
      ['Período(s)', snapshots.map(s => s.chave || '').filter(Boolean).join(', ') || 'Todos'],
      ['Filtro de Clientes', filtroLabel],
      [],
      ['INDICADOR', 'VALOR'],
      ['Total de Viagens',     _exViagens],
      ['Total de Entregas',    _exEntregas],
      ['Volume Total (m³)',    num(_exVol)],
      ['Ocupação Média (%)',   _exOcup],
      ['Km Total',             Math.round(_exKm)],
      ['Clientes Atendidos',   cliExport.length],
    ]);

    // ── Aba 2: Por Cliente ──────────────────────────────────────────────────
    const cliRows = [
      ['CLIENTE', 'CIDADE', 'ENTREGAS', 'VOLUME (m³)', 'KM MÉDIO', 'VOL/KM (m³/km)', 'OCUP. MÉDIA (%)'],
    ];
    // Mescla dados de clientes com ocupação — respeita filtro
    const ocupMap = {};
    ocupExport.forEach(c => { ocupMap[c.nome] = c.ocup; });
    cliExport.forEach(c => {
      const kmVol = c.km > 0 ? num(c.volume / c.km, 2) : '';
      cliRows.push([
        c.nome,
        c.cidade,
        c.entregas,
        num(c.volume),
        c.km > 0 ? num(c.km, 0) : '',
        kmVol,
        ocupMap[c.nome] ?? '',
      ]);
    });
    // Linha de totais
    cliRows.push([]);
    cliRows.push([
      'TOTAL', '',
      _exEntregas,
      num(_exVol),
      num(_exKm > 0 && _exEntregas > 0 ? _exKm / _exEntregas : 0, 0),
      '',
      _exOcup,
    ]);
    addSheet('Por Cliente', cliRows);

    // ── Aba 3: Por Viagem (detalhamento completo) ───────────────────────────
    const viaRows = [
      ['DATA CARGA', 'PLACA', 'TRANSPORTADORA', 'TERMINAL', 'TIPO', 'CAP. (m³)',
       'MOTORISTA', 'Nº VIAGEM', 'PARADA', 'CLIENTE', 'CIDADE', 'ORDER SAP',
       'VOLUME PARADA (m³)', 'CPT', 'PRODUTO', 'KM PARADA', 'OCUP. VIAGEM (%)'],
    ];
    snapshots.forEach(snap => {
      const res   = snap.resultado  || {};
      const vecs  = snap.veiculos   || [];
      const terms = snap.terminais  || [];
      const datasEntrega = snap.datasEntrega || [];
      const dataLabel = datasEntrega[0] || '';
      vecs.forEach(v => {
        const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
        const capV = v.capacidade || 0;
        viagens.forEach((vi, iV) => {
          let volViagem = 0;
          vi.paradas.forEach(par => { volViagem += par.volumeTotal || 0; });
          const ocupV = capV > 0 ? Math.round((volViagem / capV) * 100) : '';
          const term = terms.find(t => t.nome === v.terminal);

          vi.paradas.forEach((par, iP) => {
            const ped   = par.pedido || {};
            // Pula parada se filtro de cliente ativo e este cliente não está no filtro
            const _nomeParada = ped.cliente || ped.nomeCliente || par.nome || '';
            if (filtroAtivo && !filtroAtivo.has(_nomeParada)) return;
            const prods = (par.produtosSelecionados || par.produtos || []);
            const prodLabel = prods.map(p =>
              [p.codigoProduto || p.codigo || '', p.descricao || p.produto || ''].filter(Boolean).join(' - ')
            ).join(' | ') || '';
            const cptLabel = prods.map(p => p.cpt || p.compartimento || '').filter(Boolean).join(',') || '';

            viaRows.push([
              dataLabel,
              v.placa || '',
              v.transportadora || '',
              v.terminal || '',
              v.tipo || '',
              capV || '',
              v.motoristaDiurno || v.motorista || '',
              iV + 1,
              iP + 1,
              ped.cliente || ped.nomeCliente || par.nome || '',
              ped.cidade || '',
              ped.codigoSAP || ped.codSAP || '',
              num(par.volumeTotal || 0),
              cptLabel,
              prodLabel,
              par.distanciaKm > 0 ? num(par.distanciaKm, 0) : '',
              ocupV,
            ]);
          });
        });
      });
    });
    addSheet('Por Viagem', viaRows);

    // ── Aba 4: Ocupação por Veículo ─────────────────────────────────────────
    const ocupRows = [
      ['PLACA', 'Nº VIAGEM', 'OCUPAÇÃO (%)'],
    ];
    // Usa ocupação filtrada quando há filtro de clientes ativo
    // (viagens_ocup não tem info de cliente — usa d.viagens_ocup completo quando sem filtro)
    const _ocupExportFinal = ocupExport.length ? ocupExport : d.clientes_ocup;
    // Para ocupação por veículo mantemos d.viagens_ocup (por viagem, não por cliente)
    d.viagens_ocup.forEach(vo => {
      const parts = vo.label.split(' V');
      ocupRows.push([parts[0] || vo.label, parts[1] ? parseInt(parts[1]) : '', vo.ocup]);
    });
    addSheet('Ocupação Veículos', ocupRows);

    // ── Gera arquivo ────────────────────────────────────────────────────────
    const agora = new Date();
    const p2 = n => String(n).padStart(2, '0');
    const fname = `Dashboard_Nexta_${agora.getFullYear()}${p2(agora.getMonth()+1)}${p2(agora.getDate())}_${p2(agora.getHours())}${p2(agora.getMinutes())}.xlsx`;
    XLSX.writeFile(wb, fname);

    if (btn) { btn.textContent = '✅ Exportado!'; setTimeout(() => { btn.textContent = orig; btn.disabled = false; }, 2000); }
  } catch(e) {
    console.error('dashExportarExcel:', e);
    alert('Erro ao gerar Excel: ' + e.message);
    if (btn) { btn.textContent = orig; btn.disabled = false; }
  }
};

})();
})(); // fim IIFE dashboard
async function osrmFetchSegmento(a, b) {
  // Tenta perfil truck primeiro, cai para driving se falhar
  const perfis = ['truck', 'driving'];
  for (const perfil of perfis) {
    try {
      const url = `https://router.project-osrm.org/route/v1/${perfil}/` +
        `${a.lon},${a.lat};${b.lon},${b.lat}` +
        `?overview=full&geometries=geojson&steps=false`;
      const res = await fetch(url);
      if (!res.ok) continue;
      const data = await res.json();
      if (data.code === 'Ok' && data.routes?.[0]) {
        return {
          coords:   data.routes[0].geometry.coordinates.map(c => [c[1], c[0]]),
          distKm:   data.routes[0].distance / 1000,
        };
      }
    } catch(e) { /* tenta próximo perfil */ }
  }
  // Fallback: linha reta com distância haversine
  const R = 6371;
  const dLat = (b.lat - a.lat) * Math.PI / 180;
  const dLon = (b.lon - a.lon) * Math.PI / 180;
  const ha = Math.sin(dLat/2)**2 + Math.cos(a.lat*Math.PI/180) * Math.cos(b.lat*Math.PI/180) * Math.sin(dLon/2)**2;
  const distKm = R * 2 * Math.atan2(Math.sqrt(ha), Math.sqrt(1-ha));
  return { coords: [[a.lat, a.lon], [b.lat, b.lon]], distKm };
}
async function osrmRoute(pontos, layer, cor, peso, opacidade) {
  peso      = peso      || 4;
  opacidade = opacidade || 0.88;
  if (!pontos || pontos.length < 2) return;
  // Distância acumulada por ponto (índice 0 = origem = 0 km)
  const distAcum = [0];
  let totalKm = 0;
  // Busca segmentos em paralelo (máx 4 simultâneos) para não travar
  const BATCH = 4;
  const segmentos = new Array(pontos.length - 1);
  for (let start = 0; start < pontos.length - 1; start += BATCH) {
    const end = Math.min(start + BATCH, pontos.length - 1);
    const batch = [];
    for (let i = start; i < end; i++) {
      batch.push(osrmFetchSegmento(pontos[i], pontos[i + 1]).then(r => { segmentos[i] = r; }));
    }
    await Promise.all(batch);
    // Desenha cada segmento conforme fica pronto
    for (let i = start; i < end; i++) {
      totalKm += segmentos[i].distKm;
      distAcum.push(totalKm);
      L.polyline(segmentos[i].coords, { color: cor, weight: peso, opacity: opacidade }).addTo(layer);
    }
  }
  // Expõe as distâncias acumuladas indexadas pelo array de pontos
  // Usa o índice do layer como chave
  if (!layer._nexta_distAcum) layer._nexta_distAcum = [];
  layer._nexta_distAcum = distAcum;
  return distAcum;
}
