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
var _mvRoutePoints = [];     // [{lat,lon}] pontos REAIS do traçado (extraídos das polylines do OSRM),
                              // em ordem, do início ao fim da viagem — usado pra detectar pedágio
                              // ao longo do trajeto de verdade, não só numa linha reta entre paradas.
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
  _mvRoutePoints = [];
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
    // Extrai os pontos REAIS do traçado (em ordem) pra detecção de pedágio
    // ao longo do trajeto de verdade — não só entre as pontas das paradas.
    try {
      _mvRoutePoints = novasPolylines
        .flatMap(pl => pl.getLatLngs())
        .map(ll => ({ lat: ll.lat, lon: ll.lng }));
    } catch (e) {
      _mvRoutePoints = [];
      console.warn('[mvDesenharRota] não foi possível extrair pontos do traçado real:', e);
    }
    // Re-desenhar marcadores de via por cima das polylines
    _mvUserWaypoints.forEach(uw => {
      if (uw.marker) { try { camadaViagem.removeLayer(uw.marker); } catch(e){} }
      _mvDesenharMarcadorVia(uw);
    });
    // Traçado real agora disponível em _mvRoutePoints — recalcula o custo/pedágio
    // com o caminho de verdade em vez da aproximação por linha reta, e atualiza
    // isso toda vez que a rota é redesenhada (abertura do modal OU arrasto manual).
    if (typeof renderCustoMapaViagem === 'function') {
      try { renderCustoMapaViagem(); } catch (e) { console.warn('[mvDesenharRota] falha ao atualizar custo com traçado real:', e); }
    }
    // Defensivo: se por alguma condição de corrida (modal reaberto rápido, drag
    // simultâneo) w.marker ainda não for um L.Marker de verdade, isso não pode
    // interromper o resto da função — daí o typeof + try/catch por item.
    _mvWaypoints.forEach(w => {
      if (w.marker && typeof w.marker.bringToFront === 'function') {
        try { w.marker.bringToFront(); } catch (e) { console.warn('[mvDesenharRota] falha ao trazer marcador para frente:', e); }
      }
    });
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
    if (typeof window.renderCustoMapaViagem === 'function') window.renderCustoMapaViagem();
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
      const file = await handle.getFile();
      const text = await file.text();
      const data = JSON.parse(text);
      console.log(`[DASH] arquivo: ${name} | versao=${data.versao} | savedAt=${data.savedAt} | temResultado=${!!data.resultado} | datasEntrega=`, data.datasEntrega);
      // Arquivo já substituído por uma revisão posterior (correção salva por
      // cima) não é mais a versão vigente daquela programação — a versão
      // nova (que aponta pra este via revisaoDe) já está em outro arquivo e
      // será contada normalmente. Contar os dois ao mesmo tempo inflava
      // viagens/km/volume/entregas do Dashboard a cada correção, igual à
      // regra que a aba Histórico já aplica (linha ~7786: vigentes = entries
      // sem substituidoPor).
      if (data.substituidoPor) {
        console.warn(`[DASH] rejeitado (substituído por revisão posterior: ${data.substituidoPor}): ${name}`);
        rejeitados++;
        continue;
      }
      // Antes, arquivo sem "versao" OU sem "savedAt" era descartado inteiro —
      // igual ao Frete, agora só exige QUE HAJA DADOS DE ROTEIRIZAÇÃO
      // (resultado ou pedidos), e preenche savedAt que falte com a data de
      // modificação do arquivo (mesmo fallback já usado no Frete), em vez de
      // simplesmente jogar a roteirização inteira fora. Isso fazia o
      // Dashboard mostrar menos viagens/km/volume que o Frete pro "mesmo"
      // período, porque roteirizações salvas antes de o campo "versao" achar
      // uma casa no JSON (ou qualquer salvamento sem savedAt) sumiam aqui
      // mas continuavam contando lá.
      if (!data.resultado && !data.pedidos && !data.versao) {
        console.warn(`[DASH] rejeitado (sem resultado/pedidos/versao): ${name}`);
        rejeitados++;
        continue;
      }
      if (!data.savedAt) data.savedAt = new Date(file.lastModified).toISOString();
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
// ── Fonte dos dados do Dashboard ──────────────────────────────────────────
// Antes mesclava a pasta compartilhada em disco com um "estoque" à parte em
// localStorage (alimentado por dashSalvarAtual, uma função que ficou órfã —
// não está mais ligada a nenhum botão da tela). Isso fazia o Dashboard somar
// roteirizações que só existiam no navegador de quem um dia usou aquele
// botão, e que nunca apareceram na pasta compartilhada — daí os totais do
// Dashboard ficarem sistematicamente maiores que os da aba Histórico, que
// sempre leu só a pasta. Agora as duas telas leem exatamente a mesma fonte
// (a pasta em disco, com o mesmo filtro de "só vigente, nunca substituída"),
// então os números batem.
async function dashGetStoreMerged() {
  return await dashLerHistoricoDisco();
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
// ── Normaliza nome de cliente pra comparação (remove acentos, sufixos
// jurídicos, pontuação e espaços extras) — extraído de dashChaveCliente()
// pra poder ser reaproveitado na busca de Segmento (ver dashClienteSegmento).
function dashNormalizarNomeCliente(nome) {
  return (nome || '?').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
    .toUpperCase()
    .replace(/\b(LTDA|EIRELI|ME|EPP|SA|S\.A\.|COMERCIO|COMERCIAL|INDUSTRIA|IND|COM)\b\.?/g, '')
    .replace(/[^A-Z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
function dashChaveCliente(ped) {
  const sap = (ped.codigoSAP || ped.codSAP || ped.sap || '').toString().trim();
  if (sap) return 'SAP:' + sap;
  // Fallback: normaliza o nome removendo acentos, sufixos jurídicos e espaços extras
  const nome = (ped.cliente || ped.nomeCliente || ped.nome || '?').toString();
  return 'NM:' + dashNormalizarNomeCliente(nome);
}
// ── Nome canônico preferido (mais curto = menos abreviado) ─────────────────
function dashNomeCanônico(atual, novo) {
  if (!atual) return novo;
  // Prefere o nome mais longo (mais completo), sem truncamentos
  return novo.length > atual.length ? novo : atual;
}
// Retorna a cidade da operação (terminal) de uma viagem — usa terminalOrigem
// da própria viagem (mais preciso) com fallback pro terminal cadastrado no
// veículo. Terminal já tem campo `cidade` próprio (import Excel), separado
// do nome da base (ex.: nome "Betim POTENCIAL Nexta" → cidade "Betim").
function dashCidadeOperacaoViagem(vi, v, terms) {
  const nomeTerm = vi.terminalOrigem || v.terminal || '';
  const term = terms.find(t => t.nome === nomeTerm);
  return term?.cidade || '(sem cidade)';
}
function dashAgregar(snapshots, cidadesFiltro = null) {
  const clientes = {};   // key=nome: {entregas, volume, km, lat, lon, cidade, capTotal}
  const operacoes = {};  // key=cidade da operação: {cidade, volume, capTotal, viagens} — pro gráfico Ocupação vs Volume por Operação
  const viagens_ocup = []; // {label, ocup}
  let totalViagens = 0, totalKm = 0, totalVol = 0, totalCap = 0;
  let totalViagensComPedagio = 0; // pra "% de Rotas Pedagiadas" e "Tempo Gasto com Lançamento de Pedágios"
  // veiculos_escalados: lista de {id, snapIdx, capV, viagensIds} para cálculo de ocupação filtrada
  const veiculos_escalados = [];
  const rotasMap = [];   // para o mapa: [{termLat,termLon,paradas:[{lat,lon,nome}]}]
  // Entradas cruas por viagem, usadas pelo Ranking de Transportadoras
  // (dashAgregarTransportadoras) — evita rodar esse loop duas vezes.
  const entradasTransportadora = [];
  // "placa normalizada + dia" de todo veículo que teve viagem de verdade
  // naquele dia (respeitando o filtro de cidade) — usado pelo indicador de
  // Ociosidade pra saber se um veículo marcado "Disponível" no Painel de
  // Disponibilidade (fonte de verdade de quem foi ESCALADO no dia — ver
  // dashCarregarOciosidade) foi realmente utilizado ou ficou parado.
  const diasComViagemPorPlaca = new Set();
  // placa normalizada -> cidade — construído a partir do MESMO critério usado
  // pra filtrar tudo o mais no Dashboard (terminal do veículo, via
  // dashCidadeOperacaoViagem), não do campo "Operação" cadastrado
  // separadamente na tela de Placas. Os dois podiam divergir (nomes
  // diferentes, cadastro desatualizado), fazendo a Ociosidade nunca bater
  // com nenhuma cidade do filtro e sempre voltar 0/0 quando filtrado —
  // usando essa mesma fonte, ela fica automaticamente consistente com
  // Jornada/Ranking/Viagens/etc.
  // Guardado POR DIA (placa+data), não só por placa: um mesmo veículo pode
  // operar de terminais/cidades diferentes em dias diferentes (troca de
  // operação), e usar só a "primeira cidade vista" pra placa inteira fazia a
  // Ociosidade comparar com a cidade errada nesses casos — sempre dando 0/0
  // ao filtrar uma operação que na verdade tinha viagens daquela placa.
  const placaCidadePorDia = new Map();
  const placaCidade = new Map(); // fallback: última cidade vista pra placa, sem dia (cobre disponibilidade sem viagem correspondente naquele dia)
  snapshots.forEach((snap, snapIdx) => {
    const res = snap.resultado || {};
    const vecs = snap.veiculos || [];
    const terms = snap.terminais || [];
    const mesKeySnap = (snap.savedAt || '').slice(0,7);
    const dataSnap = (snap.savedAt || '').slice(0,10);
    vecs.forEach(v => {
      const pNormMap = (v.placa || '').trim().toUpperCase();
      const viagensTodas = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
      if (pNormMap) {
        // Cidade de CADA viagem de fato realizada (vi.terminalOrigem||v.terminal,
        // o mesmo critério usado pra filtrar Viagens/Entregas/etc.) — guardada
        // como conjunto, pois um veículo pode ter rodado de terminais/cidades
        // diferentes no MESMO dia. Usar só o terminal padrão do veículo (v.terminal)
        // aqui, ignorando o terminal real de cada viagem, fazia esse mapa
        // divergir da cidade que realmente contou pro filtro em cada viagem.
        const cidadesDoDia = viagensTodas.length
          ? viagensTodas.map(vi => dashCidadeOperacaoViagem(vi, v, terms))
          : [dashCidadeOperacaoViagem({}, v, terms)]; // sem viagem no dia: usa terminal padrão do veículo como melhor palpite
        cidadesDoDia.forEach(c => {
          if (dataSnap) {
            const chaveDia = pNormMap + '__' + dataSnap;
            if (!placaCidadePorDia.has(chaveDia)) placaCidadePorDia.set(chaveDia, new Set());
            placaCidadePorDia.get(chaveDia).add(c);
          }
          if (!placaCidade.has(pNormMap)) placaCidade.set(pNormMap, new Set());
          placaCidade.get(pNormMap).add(c);
        });
      }
      // Jornada (Consumo de Jornada) mede só quem teve viagem de verdade —
      // veículo sem nenhuma viagem no dia simplesmente não entra nessa conta.
      if (!viagensTodas.length) return;
      // Filtro de cidade da operação: aplica ANTES de qualquer acumulação, pra
      // que viagens, entregas, ocupação, km — tudo — reflita só a(s) cidade(s)
      // selecionada(s). null = sem filtro (todas as cidades).
      const viagens = cidadesFiltro
        ? viagensTodas.filter(vi => cidadesFiltro.has(dashCidadeOperacaoViagem(vi, v, terms)))
        : viagensTodas;
      if (!viagens.length) return;
      diasComViagemPorPlaca.add((v.placa || '').trim().toUpperCase() + '__' + dataSnap);
      // capV é acumulado por VIAGEM realizada (dentro do loop de viagens abaixo)
      const capV = v.capacidade || v.capacidadeTotal || 0;
      viagens.forEach((vi, iV) => {
        totalViagens++;
        // Terminal DESTA viagem específica — prioriza o terminal do PEDIDO
        // (vi.paradas[x].pedido.terminal), não o do veículo (v.terminal).
        // Desde que passou a ser possível um veículo carregar em mais de uma
        // base da mesma cidade, v.terminal ficou vazio de propósito (o
        // veículo não tem mais UM terminal fixo) — usar só v.terminal fazia
        // a detecção de pedágio (e o ponto de origem no mapa) nunca achar
        // NENHUM terminal pra NENHUMA viagem, silenciosamente. Mantém
        // v.terminal como fallback só pra arquivos antigos, de antes dessa
        // mudança, que ainda tinham o campo preenchido no veículo.
        const termNomeViagem = vi.terminalOrigem || vi.paradas?.find(p => p.pedido?.terminal)?.pedido?.terminal || v.terminal;
        const term = terms.find(t => t.nome === termNomeViagem);
        const tLat = term?.lat, tLon = term?.lon;
        const rotaPontos = { termLat: tLat, termLon: tLon, placa: v.placa, paradas: [] };
        // Pontos pra detecção de pedágio (% de Rotas Pedagiadas / Tempo Gasto
        // com Lançamento de Pedágios): terminal de origem + cada parada na
        // ordem da viagem — mesmo formato que detectarPedagiosNaRota espera.
        // Usa a LINHA RETA entre os pontos (não o trajeto real via OSRM, que
        // exigiria uma chamada de rede por viagem — inviável para centenas
        // de viagens históricas de uma vez), com raioKm=3 igual ao fallback
        // já usado pelo relatório de Frete pro mesmo cenário (ver comentário
        // em detectarPedagiosNaRota, em pedagios-utils.js). Isso é uma
        // ESTIMATIVA: pode errar pra mais (raio largo pega pedágio perto mas
        // fora do trajeto real) ou pra menos (trajeto real com curva que a
        // reta não capta) — serve como indicador de tendência, não como
        // conferência financeira exata (essa já existe no Roteirizador/Frete,
        // com o trajeto real).
        const pontosPedagio = (tLat != null && tLon != null && !isNaN(parseFloat(tLat)) && !isNaN(parseFloat(tLon)))
          ? [{ lat: parseFloat(tLat), lon: parseFloat(tLon) }]
          : [];
        let volViagem = 0;
        let kmIdaViagem = 0; // soma do km "base" (sem duplicar ida+volta) — usado no Ranking de Transportadoras
        // Km real da viagem — fonte única: NextaKm.obterKmViagem() (ver
        // assets/km-utils.js). Se o usuário ajustou a rota manualmente no
        // mapa, ou se o Dashboard já rodou o recálculo em lote,
        // vi._kmAjustado guarda a distância real via rota — essa é a fonte
        // mais fiel disponível. Quando não há isso ainda, mantém o cálculo
        // padrão (Haversine sequencial entre paradas, sem fator de inflação).
        const _kmInfoViagem = NextaKm.obterKmViagem(vi);
        const _kmAjustadoViagem = _kmInfoViagem.real ? _kmInfoViagem.km : null;
        let _somaKmOriginalViagem = null;
        if (_kmAjustadoViagem != null) {
          _somaKmOriginalViagem = NextaKm.estimativaLinhaReta(vi);
        }
        vi.paradas.forEach(par => {
          const ped = par.pedido || {};
          const nome = ped.cliente || ped.nomeCliente || par.nome || '?';
          const cidade = ped.cidade || '-';
          const vol = par.volumeTotal || 0;
          const coords = latLonEfetivo ? latLonEfetivo(ped) : { lat: par.lat, lon: par.lon };
          const lat = coords?.lat || par.lat;
          const lon = coords?.lon || par.lon;
          // Km base da parada: distanciaKm calculado na roteirização (Haversine
          // sequencial entre paradas), com fallback para Haversine terminal→cliente.
          const kmBase = par.distanciaKm > 0 ? par.distanciaKm
            : ((tLat && tLon && lat && lon) ? haversine(tLat, tLon, lat, lon) : 0);
          // Se a viagem foi ajustada manualmente no mapa, rateia o km real (OSRM)
          // entre as paradas na mesma proporção do km original de cada uma —
          // preserva o km médio por cliente mesmo com a rota redesenhada.
          const km = (_kmAjustadoViagem != null)
            ? (_somaKmOriginalViagem > 0
                ? _kmAjustadoViagem * (kmBase / _somaKmOriginalViagem)
                : _kmAjustadoViagem / vi.paradas.length)
            : kmBase;
          const chave = dashChaveCliente(ped);
          if (!clientes[chave]) clientes[chave] = { nome, cidade, entregas:0, volume:0, km:0, kmTotal:0, lat, lon, capTotal:0, viagensIds: new Set(), codigoSAP: '' };
          clientes[chave].nome = dashNomeCanônico(clientes[chave].nome, nome);
          if (!clientes[chave].codigoSAP && ped.codigoSAP) clientes[chave].codigoSAP = String(ped.codigoSAP).trim();
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
          kmIdaViagem += km;
          rotaPontos.paradas.push({ lat, lon, nome, vol });
          // Mesma trava de coordenada usada no km real — centralizada em
          // NextaKm.coordenadaValida (ver assets/km-utils.js).
          if (NextaKm.coordenadaValida(lat, lon)) {
            pontosPedagio.push({ lat: parseFloat(lat), lon: parseFloat(lon) });
          }
        });
        rotasMap.push(rotaPontos);
        // Detecção de pedágio desta viagem (ver comentário acima de
        // pontosPedagio) — conta 1 viagem "pedagiada" se pelo menos 1 praça
        // for detectada no trajeto, independente de quantas praças tenha.
        if (pontosPedagio.length >= 2) {
          const _eixosV = v.eixos || 2;
          const _pedagiosDetectados = (typeof detectarPedagiosNaRota === 'function')
            ? detectarPedagiosNaRota(pontosPedagio, _eixosV, 3)
            : [];
          if (_pedagiosDetectados.length > 0) totalViagensComPedagio++;
        }
        // Data REAL da entrega desta viagem (não a data em que a
        // roteirização foi SALVA) — uma sessão/arquivo pode cobrir vários
        // dias de entrega de uma vez (comum, ver datasEntrega), então usar
        // "dataSnap" (savedAt) aqui juntava jornada de dias diferentes no
        // mesmo balde, inflando o "Estouro de Jornada" pra números
        // impossíveis (ex.: 90h somadas de 8 dias diferentes, todos com o
        // mesmo savedAt, aparecendo como se fosse 1 dia só).
        const _dataEntregaBr = vi.paradas?.[0]?.pedido?.dataEntregaLogistica || vi.paradas?.[0]?.pedido?.dataEntrega || '';
        const _mtDataEntrega = String(_dataEntregaBr).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        const dataViagemReal = _mtDataEntrega
          ? `${_mtDataEntrega[3]}-${_mtDataEntrega[2].padStart(2,'0')}-${_mtDataEntrega[1].padStart(2,'0')}`
          : dataSnap;
        // Entrada crua desta viagem para o Ranking de Transportadoras — cálculo
        // de custo (contrato) é feito depois, em dashAgregarTransportadoras,
        // porque precisa do total de viagens do mês (rateio do valor fixo).
        entradasTransportadora.push({
          transportadora: v.transportadora || '(sem transportadora)',
          placa: v.placa,
          data: dataViagemReal,
          mesKey: mesKeySnap,
          termOrigem: vi.terminalOrigem || v.terminal || '',
          destinos: Array.from(new Set(vi.paradas.map(p => (p.pedido?.cidade || p.pedido?.cliente || '')))).join(', '),
          kmIda: kmIdaViagem,
          volume: volViagem,
          jornadaDispMin: Number(v.jornadaMin) || (typeof duracaoJornadaMin === 'function' ? duracaoJornadaMin(v.jornadaInicio || '06:00', v.jornadaFim || '18:00') : 720) || 720,
          jornadaUsadaMin: Number(vi.tempoConsumidoMin) || vi.paradas.reduce((s,p,idx) =>
            s + (idx === 0 ? (p.tempoCarregamentoMin || 0) : 0)
              + (p.waitAfterLoadingMin || 0) + (p.deslocCarregadoMin || 0)
              // Espera overnight (motorista descansando de madrugada até o
              // cliente abrir) NÃO conta como jornada trabalhada — mesma
              // isenção já usada no cálculo "ao vivo" da tela (renderResultado)
              // e no Herrlog. Sem isso, um trecho overnight de 8-12h de
              // espera inflava a jornada usada pra números impossíveis
              // (90h+ num dia), mesmo depois da correção da data.
              + (p.overnight ? 0 : (p.tempoEsperaRestricaoMin || 0))
              + (p.tempoDescargaMin || 0) + (p.deslocVazioMin || 0), 0),
        });
        // Acumula por OPERAÇÃO (cidade do terminal desta viagem) — volume e
        // capacidade, uma vez por viagem (mesma regra usada pra ocupação de
        // cliente/frota acima), pro gráfico "Ocupação vs Volume por Operação".
        // Já sai filtrado corretamente porque `viagens` acima já respeitou
        // cidadesFiltro antes de chegar aqui.
        {
          const cidadeOp = dashCidadeOperacaoViagem(vi, v, terms);
          if (!operacoes[cidadeOp]) operacoes[cidadeOp] = { cidade: cidadeOp, volume: 0, capTotal: 0, viagens: 0 };
          operacoes[cidadeOp].volume += volViagem;
          operacoes[cidadeOp].viagens += 1;
          if (capV > 0) operacoes[cidadeOp].capTotal += capV;
        }
        if (capV > 0) {
          const ocup = Math.round((volViagem / capV) * 100);
          viagens_ocup.push({ label: `${v.placa} V${iV+1}`, ocup, snapIdx, vid: v.id, iV });
          // Capacidade acumulada por viagem realizada: veículo de 35m³ que fez 2 viagens = 70m³
          totalCap += capV;
          veiculos_escalados.push({ snapIdx, vid: v.id, capV, iV });
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
  // Ocupação por operação (cidade): volume da operação / capacidade acumulada
  // dos veículos que operaram a partir dela — mesma lógica de clientes_ocup,
  // só que agrupado por cidade em vez de por cliente.
  const operacoes_ocup = Object.values(operacoes)
    .filter(o => o.capTotal > 0)
    .map(o => ({
      nome: o.cidade,
      ocup: Math.min(100, Math.round((o.volume / o.capTotal) * 100)),
      volume: parseFloat(o.volume.toFixed(1)),
      viagens: o.viagens,
    }))
    .sort((a, b) => b.volume - a.volume);
  return {
    clientes: Object.values(clientes).sort((a,b)=>b.volume-a.volume),
    viagens_ocup,
    clientes_ocup,
    operacoes_ocup,
    veiculos_escalados, // [{snapIdx, vid, capV}] para cálculo de ocupação filtrada
    totalViagens,
    totalViagensComPedagio,
    totalEntregas,
    totalVol: parseFloat(totalVol.toFixed(1)),
    totalKm: Math.round(totalKm),
    totalOcup: totalVec,
    totalClientes: Object.keys(clientes).length,
    rotasMap,
    entradasTransportadora,
    diasComViagemPorPlaca,
    placaCidade,
    placaCidadePorDia,
  };
}
// ── Ranking de Transportadoras ───────────────────────────────────────────────
// Reaproveita os contratos já cadastrados na aba Frete (freteCarregarContratos/
// freteCarregarSpot) e replica as MESMAS fórmulas de custo usadas lá
// (kmEfetivo + custoViagem), pra o valor pago aqui bater com o que aparece
// no relatório de Frete. Se uma placa não tiver contrato cadastrado, o custo
// dela entra como 0 e ela é marcada `semContrato: true` — aparece no ranking
// por volume/km normalmente, só o valor pago fica indefinido (mostrado como
// "—", nunca inventado).
function _dashKmEfetivoRanking(entry, contrato) {
  const modo = (contrato && contrato.kmModo) || 'ida_volta';
  return modo === 'ida' ? entry.kmIda : entry.kmIda * 2;
}
function _dashCustoViagemRanking(entry, contrato, nViagensMesPlaca, spots) {
  const nViagMes = nViagensMesPlaca || 1;
  const fixo = parseFloat(contrato.fixo) || 0;
  const fixoRateado = fixo / Math.max(nViagMes, 1);
  const km = _dashKmEfetivoRanking(entry, contrato);
  if (contrato.tipo === 'fixo_km') return fixoRateado + (parseFloat(contrato.km)||0) * km;
  if (contrato.tipo === 'fixo_m3') return fixoRateado + (parseFloat(contrato.m3)||0) * entry.volume;
  // Diária = obrigação por dia corrido do veículo Dedicado (viajou ou não) —
  // mesma regra aplicada no Frete. entry._diariaFlatHoje marca qual entrada
  // do dia leva a cobrança cheia (evita duplicar em dias com 2+ viagens).
  if (contrato.tipo === 'diaria')  return (parseFloat(contrato.diaria) || 0) * (entry._diariaFlatHoje ? 1 : 0);
  if (contrato.tipo === 'diaria_km') return ((parseFloat(contrato.diaria)||0) * (entry._diariaFlatHoje ? 1 : 0)) + (parseFloat(contrato.km)||0) * km;
  if (contrato.tipo === 'spot') {
    const sp = (spots || []).find(s =>
      entry.termOrigem.toLowerCase().includes((s.origem||'').toLowerCase()) &&
      entry.destinos.toLowerCase().includes((s.destino||'').toLowerCase()) &&
      (!s.transportadora || s.transportadora === contrato.transportadora)
    );
    return sp ? (parseFloat(sp.valor)||0) * entry.volume : 0;
  }
  return 0;
}
// Completa os dias sem viagem de placas com contrato "diaria"/"diaria_km"
// com uma entrada vazia (0 km, 0 m³) só pra carregar a cobrança da diária
// daquele dia — mesma lógica do Frete. Período de referência: min/max data
// já presente nas entradas carregadas (reflete o mês/período que o usuário
// escolheu no filtro do Dashboard).
function _dashCompletarDiasParados(entradas, contratos) {
  const normPlaca = (typeof _freteNormPlaca === 'function') ? _freteNormPlaca : (p => (p||'').toString().trim().toUpperCase());
  const placasDiaria = new Set();
  contratos.forEach(c => { if ((c.tipo === 'diaria' || c.tipo === 'diaria_km') && c.placa) placasDiaria.add(normPlaca(c.placa)); });
  if (!placasDiaria.size) return entradas;
  // Período de referência: usa o intervalo dos snapshots REALMENTE carregados
  // (mês selecionado / todos os períodos) — não só os dias com viagem, senão
  // dias parados no fim do mês (sem NENHUMA viagem de ninguém naquele dia)
  // ficariam de fora da cobrança por engano.
  const datasSnap = (typeof _dashSnapshotsAtivos !== 'undefined' ? _dashSnapshotsAtivos : [])
    .map(s => (s.savedAt || '').slice(0,10)).filter(Boolean).sort();
  if (!datasSnap.length) return entradas;
  const dIni = new Date(datasSnap[0] + 'T00:00:00');
  const dFim = new Date(datasSnap[datasSnap.length-1] + 'T00:00:00');
  if (isNaN(dIni) || isNaN(dFim)) return entradas;

  const porPlaca = {};
  entradas.forEach(e => { (porPlaca[e.placa] ||= []).push(e); });
  // Garante que TODA placa com contrato de diária entra no loop abaixo, mesmo
  // que nunca tenha feito nenhuma viagem no período carregado (senão ela nem
  // aparece em `entradas` e ficaria de fora da cobrança e do próprio ranking).
  contratos.forEach(c => {
    if ((c.tipo !== 'diaria' && c.tipo !== 'diaria_km') || !c.placa) return;
    const placaNorm = normPlaca(c.placa);
    const jaExiste = Object.keys(porPlaca).some(p => normPlaca(p) === placaNorm);
    if (!jaExiste) porPlaca[c.placa] = [];
  });
  const extras = [];
  Object.keys(porPlaca).forEach(placa => {
    if (!placasDiaria.has(normPlaca(placa))) return;
    const porDia = {};
    porPlaca[placa].forEach(e => { (porDia[e.data] ||= []).push(e); });
    const vCad = (typeof veiculos !== 'undefined') ? veiculos.find(v => normPlaca(v.placa) === normPlaca(placa)) : null;
    const contratoDaPlaca = contratos.find(c => normPlaca(c.placa) === normPlaca(placa));
    const transp = (porPlaca[placa][0] && porPlaca[placa][0].transportadora) || (contratoDaPlaca && contratoDaPlaca.transportadora) || (vCad && vCad.transportadora) || '(sem transportadora)';
    const mesKeyOf = d => d.slice(0,7);
    for (let d = new Date(dIni); d <= dFim; d.setDate(d.getDate()+1)) {
      const diaStr = d.toISOString().slice(0,10);
      const entradasHoje = porDia[diaStr];
      if (entradasHoje && entradasHoje.length) {
        entradasHoje.forEach((e, idx) => { e._diariaFlatHoje = (idx === 0); });
      } else {
        extras.push({
          transportadora: transp, placa, data: diaStr, mesKey: mesKeyOf(diaStr),
          termOrigem: '', destinos: '(sem viagem — diária)', kmIda: 0, volume: 0,
          jornadaDispMin: 720, jornadaUsadaMin: 0, fatorJornada: 0,
          _semViagem: true, _diariaFlatHoje: true,
        });
      }
    }
  });
  return entradas.concat(extras);
}
function dashAgregarTransportadoras(entradasTransportadora) {
  const contratos = (typeof freteCarregarContratos === 'function') ? freteCarregarContratos() : [];
  const spots     = (typeof freteCarregarSpot === 'function') ? freteCarregarSpot() : [];
  const normPlaca = (typeof _freteNormPlaca === 'function') ? _freteNormPlaca : (p => (p||'').toString().trim().toUpperCase());
  // Índice placa normalizada → contrato, montado UMA VEZ (O(m)) em vez de
  // varrer a lista de contratos pra cada viagem (O(n×m) — era o gargalo que
  // travava o dashboard em qualquer filtro com muitos dados).
  const contratoPorPlaca = new Map();
  contratos.forEach(c => { if (c.placa) contratoPorPlaca.set(normPlaca(c.placa), c); });
  // Completa dias parados de veículos com diária ANTES de tudo — precisa
  // rodar cedo pra esses dias entrarem na contagem de nViagensPorPlacaMes
  // (que, note, exclui _semViagem logo abaixo) e no fatorJornada.
  const entradasCompletas = _dashCompletarDiasParados(entradasTransportadora, contratos);
  // fatorJornada por entrada (precisa antes do custo, igual ao Frete)
  entradasCompletas.forEach(e => {
    e.fatorJornada = e.jornadaDispMin > 0 ? Math.min(1, Math.max(0, e.jornadaUsadaMin / e.jornadaDispMin)) : 0;
  });
  // Nº de viagens por placa+mês, pra ratear o valor fixo mensal (igual ao
  // Frete) — exclui os dias parados sintéticos, senão dilui o rateio do fixo.
  const nViagensPorPlacaMes = {};
  entradasCompletas.forEach(e => {
    if (e._semViagem) return;
    const k = e.placa + '__' + e.mesKey;
    nViagensPorPlacaMes[k] = (nViagensPorPlacaMes[k] || 0) + 1;
  });
  const porTransportadora = {};
  entradasCompletas.forEach(e => {
    const contrato = contratoPorPlaca.get(normPlaca(e.placa));
    const nMes = nViagensPorPlacaMes[e.placa + '__' + e.mesKey] || 1;
    const custo = contrato ? _dashCustoViagemRanking(e, contrato, nMes, spots) : 0;
    const key = e.transportadora;
    if (!porTransportadora[key]) porTransportadora[key] = { transportadora: key, volume: 0, km: 0, viagens: 0, custo: 0, temContrato: false, placas: new Set() };
    porTransportadora[key].volume += e.volume;
    porTransportadora[key].km += _dashKmEfetivoRanking(e, contrato || { kmModo: 'ida_volta' });
    if (!e._semViagem) porTransportadora[key].viagens += 1; // dia parado (diária) não conta como viagem
    porTransportadora[key].placas.add(e.placa);
    if (contrato) { porTransportadora[key].custo += custo; porTransportadora[key].temContrato = true; }
  });
  return Object.values(porTransportadora)
    .map(t => ({ ...t, nPlacas: t.placas.size, placas: undefined }))
    .sort((a, b) => b.volume - a.volume);
}
// ── Consumo de Jornada por Transportadora ────────────────────────────────────
// Jornada DISPONÍVEL de um veículo é uma capacidade DIÁRIA (ex.: 12h/dia) —
// não pode ser somada uma vez por VIAGEM, senão um veículo que fez 3 viagens
// no mesmo dia contaria 3x a jornada disponível dele naquele dia. Por isso
// agrupa primeiro por veículo+dia (chave placa+data), pega a jornada
// disponível UMA VEZ por essa chave, e soma a jornada USADA de todas as
// viagens daquele veículo naquele dia (essa sim soma normalmente, porque
// cada viagem realmente consome tempo adicional da jornada do dia).
function dashAgregarJornada(entradasTransportadora) {
  const porVeiculoDia = new Map(); // chave: placa+data
  (entradasTransportadora || []).forEach(e => {
    if (e._semViagem) return; // dia parado sintético (só diária, do Frete) — sem jornada real pra contar
    const key = e.placa + '__' + e.data;
    if (!porVeiculoDia.has(key)) {
      porVeiculoDia.set(key, { placa: e.placa, data: e.data, transportadora: e.transportadora, dispMin: e.jornadaDispMin || 0, usadoMin: 0 });
    }
    porVeiculoDia.get(key).usadoMin += e.jornadaUsadaMin || 0;
  });
  const porTransportadora = {};
  let totalDispMin = 0, totalUsadoMin = 0;
  // ── Estouro de jornada ─────────────────────────────────────────────────
  // O agregado por transportadora acima (usadoMin/dispMin somados de TODOS
  // os veículos) já pinta em âmbar quando passa de 100%, mas isso é uma
  // MÉDIA — 1 veículo estourando feio num dia com 19 outros dentro do
  // normal quase não move o agregado. Aqui conta CADA veículo/dia
  // individualmente: estourou (usadoMin > dispMin) ou não, sem diluir.
  let diasComEstouro = 0, minutosEstouroTotal = 0;
  const estourosPorTransportadora = {};
  const estourosDetalhe = [];
  porVeiculoDia.forEach(reg => {
    totalDispMin  += reg.dispMin;
    totalUsadoMin += reg.usadoMin;
    const key = reg.transportadora;
    if (!porTransportadora[key]) porTransportadora[key] = { transportadora: key, dispMin: 0, usadoMin: 0, veiculosDia: 0 };
    porTransportadora[key].dispMin  += reg.dispMin;
    porTransportadora[key].usadoMin += reg.usadoMin;
    porTransportadora[key].veiculosDia += 1;

    if (!estourosPorTransportadora[key]) estourosPorTransportadora[key] = { transportadora: key, diasComEstouro: 0, minutosEstouroTotal: 0, veiculosDia: 0 };
    estourosPorTransportadora[key].veiculosDia += 1;
    if (reg.dispMin > 0 && reg.usadoMin > reg.dispMin) {
      const estouroMin = reg.usadoMin - reg.dispMin;
      diasComEstouro++;
      minutosEstouroTotal += estouroMin;
      estourosPorTransportadora[key].diasComEstouro += 1;
      estourosPorTransportadora[key].minutosEstouroTotal += estouroMin;
      estourosDetalhe.push({
        placa: reg.placa,
        transportadora: reg.transportadora,
        data: reg.data,
        dispMin: reg.dispMin,
        usadoMin: reg.usadoMin,
        estouroMin,
        estouroPct: Math.round((reg.usadoMin / reg.dispMin) * 100),
      });
    }
  });
  const porTransportadoraArr = Object.values(porTransportadora)
    .map(t => ({ ...t, pct: t.dispMin > 0 ? Math.round((t.usadoMin / t.dispMin) * 100) : 0 }))
    .sort((a, b) => b.usadoMin - a.usadoMin);
  const totalVeiculoDia = porVeiculoDia.size;
  const estourosPorTransportadoraArr = Object.values(estourosPorTransportadora)
    .filter(t => t.diasComEstouro > 0)
    .map(t => ({ ...t, pctDiasComEstouro: t.veiculosDia > 0 ? Math.round((t.diasComEstouro / t.veiculosDia) * 100) : 0 }))
    .sort((a, b) => b.diasComEstouro - a.diasComEstouro);
  estourosDetalhe.sort((a, b) => b.estouroMin - a.estouroMin);
  return {
    totalDispMin,
    totalUsadoMin,
    totalPct: totalDispMin > 0 ? Math.round((totalUsadoMin / totalDispMin) * 100) : 0,
    porTransportadora: porTransportadoraArr,
    // Estouro de jornada
    totalVeiculoDia,
    diasComEstouro,
    pctDiasComEstouro: totalVeiculoDia > 0 ? Math.round((diasComEstouro / totalVeiculoDia) * 100) : 0,
    minutosEstouroTotal,
    estourosPorTransportadora: estourosPorTransportadoraArr,
    estourosDetalhe: estourosDetalhe.slice(0, 30), // top 30 piores dias-veículo, pra não pesar a tela
  };
}
function _dashFmtHoras(min) {
  return (min / 60).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + 'h';
}
let _dashUltimaJornada = { porTransportadora: [] };
let _dashJornadaDirecao = 'desc'; // 'desc' = maior primeiro, 'asc' = menor primeiro
function dashJornadaSetDirecao(direcao) {
  _dashJornadaDirecao = direcao;
  document.querySelectorAll('.dash-jornadadir-tab').forEach(b => {
    const ativo = b.dataset.dir === direcao;
    b.classList.toggle('active-rank', ativo);
    b.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    b.style.color = ativo ? '#000' : 'var(--text-2)';
  });
  dashRenderJornadaTransportadoras(_dashUltimaJornada);
}
function dashRenderJornadaTransportadoras(dados) {
  _dashUltimaJornada = dados;
  const box = document.getElementById('dash-jornada-transp');
  if (!box) return;
  const lista = [...(dados.porTransportadora || [])]
    .sort((a, b) => _dashJornadaDirecao === 'asc' ? a.pct - b.pct : b.pct - a.pct);
  if (!lista.length) {
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Nenhum dado de jornada para este período/filtro.</div>`;
    return;
  }
  const max = Math.max(...lista.map(t => t.pct), 1);
  box.innerHTML = lista.map((t) => {
    const pct = Math.max(2, (t.pct / max) * 100);
    // Acima de 100% (veículo trabalhando além da jornada nominal) pinta em
    // âmbar como aviso; do contrário, segue a cor padrão do indicador.
    const cor = t.pct > 100 ? '#f0be40' : '#00d9c0';
    return `
      <div style="display:flex;align-items:center;gap:12px;padding:10px 4px;border-bottom:1px solid var(--border-dk);">
        <div style="width:170px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px;font-weight:600;color:var(--text);" title="${t.transportadora}">${t.transportadora}</div>
        <div style="flex:1;background:rgba(255,255,255,.06);border-radius:6px;height:20px;position:relative;overflow:hidden;">
          <div style="height:100%;width:${Math.min(pct,100)}%;background:${cor};border-radius:6px;transition:width .3s;"></div>
        </div>
        <div style="width:60px;flex-shrink:0;text-align:right;font-size:12.5px;font-weight:700;color:var(--text);">${t.pct}%</div>
        <div style="width:170px;flex-shrink:0;text-align:right;font-size:10.5px;color:var(--text-3);">${_dashFmtHoras(t.usadoMin)} / ${_dashFmtHoras(t.dispMin)} · ${t.veiculosDia} veíc./dia</div>
      </div>`;
  }).join('');
}
// Diferente da Jornada (que mede aproveitamento de TEMPO de quem trabalhou,
// vindo das roteirizações), Ociosidade mede quantos VEÍCULOS foram
// DISPONIBILIZADOS pelo transportador no dia — e essa é uma pergunta do
// Painel de Disponibilidade (coleção "availability" do Firestore, um
// registro por placa/dia), não da roteirização. Usar a roteirização pra
// "disponibilizados" inflava o número: quando há mais de uma roteirização
// salva no mesmo dia (turnos/terminais diferentes, não uma correção da
// outra), o mesmo veículo aparecia contado 2x, 3x... — o Painel tem UM
// registro por veículo/dia, então não tem esse risco de duplicar.
// Token pra descartar resultado de uma busca antiga se o filtro mudar de
// novo antes dela terminar (mesmo padrão usado no refino de pedágio do Frete).
let _dashOciosidadeToken = 0;
// Cache simples por intervalo de datas — evita reconsultar o Firestore toda
// vez que o Dashboard re-renderiza (troca de filtro de cidade/cliente, por
// exemplo) pedindo o MESMO período de novo. TTL curto o suficiente pra não
// mostrar dado velho por muito tempo, longo o suficiente pra absorver uma
// sequência de cliques em filtros dentro de poucos minutos.
const _dashOciosidadeCacheDocs = new Map(); // "dataIni_dataFim" -> { docs, expiraEm }
const _DASH_OCIOSIDADE_CACHE_TTL_MS = 120_000; // 2 minutos
async function dashCarregarOciosidade(snapshotsAtivos, cidadesFiltro, diasComViagemPorPlaca, placaCidade, placaCidadePorDia) {
  const vazio = { totalDisponibilizados: 0, totalUsados: 0, pctOciosidade: 0, porTransportadora: [] };
  if (!window.fbDb || !window.fbCollection || !window.fbQuery || !window.fbWhere || !window.fbGetDocs || typeof window.dbGetPlates !== 'function') {
    console.warn('[Ociosidade] Firestore/dbGetPlates não disponível ainda.');
    return vazio;
  }
  const datas = (snapshotsAtivos || []).map(s => (s.savedAt || '').slice(0,10)).filter(Boolean).sort();
  if (!datas.length) return vazio;
  const dataIni = datas[0], dataFim = datas[datas.length - 1];
  const normPlaca = p => (p || '').toString().trim().toUpperCase();
  const normCidade = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
  const cidadesFiltroNorm = cidadesFiltro ? new Set([...cidadesFiltro].map(normCidade)) : null;
  // Placas ativas cadastradas por transportadora — usa config/plates (mesma
  // consulta) pra saber quem está ativo (placa desativada não conta) E qual
  // a OPERAÇÃO cadastrada em cada placa (campo "operacao" do cadastro no
  // sistema pai) — essa é a fonte PRIMÁRIA de cidade/operação pra Ociosidade,
  // por ser o dado cadastrado manualmente e não uma inferência. Como fallback
  // (placa sem operação cadastrada), cai pra cidade inferida do terminal
  // usado na roteirização (placaCidadePorDia/placaCidade, vindos do resto do
  // Dashboard) — mesma lógica usada em Jornada/Ranking/Viagens/etc.
  let todasPlacas = {};
  try { todasPlacas = await window.dbGetPlates(); } catch(e) { console.warn('[Ociosidade] falha ao ler placas:', e); }
  const placaAtiva = new Map(); // placaNorm -> ativo
  const placaOperacao = new Map(); // placaNorm -> operação cadastrada na placa
  Object.values(todasPlacas || {}).forEach(placas => {
    (placas || []).forEach(p => {
      const pNormCad = normPlaca(p.placa);
      placaAtiva.set(pNormCad, p.ativo !== false);
      if (p.operacao) placaOperacao.set(pNormCad, p.operacao);
    });
  });
  // Consulta a coleção "availability" no intervalo de datas do período
  // carregado (mês selecionado, ou min/max de "Todos os períodos") — só
  // depende do intervalo, não do filtro de cidade (esse é aplicado depois,
  // em cima do resultado), então dá pra cachear por dataIni+dataFim.
  const chaveCache = `${dataIni}_${dataFim}`;
  let docs;
  const cacheado = _dashOciosidadeCacheDocs.get(chaveCache);
  if (cacheado && cacheado.expiraEm > Date.now()) {
    docs = cacheado.docs;
  } else {
    try {
      const q = window.fbQuery(
        window.fbCollection(window.fbDb, 'availability'),
        window.fbWhere('dateStr', '>=', dataIni),
        window.fbWhere('dateStr', '<=', dataFim)
      );
      const snap = await window.fbGetDocs(q);
      docs = snap.docs.map(d => d.data());
      _dashOciosidadeCacheDocs.set(chaveCache, { docs, expiraEm: Date.now() + _DASH_OCIOSIDADE_CACHE_TTL_MS });
    } catch(e) {
      console.warn('[Ociosidade] falha ao consultar disponibilidade:', e);
      return vazio;
    }
  }
  const porTransportadora = {};
  let totalDisponibilizados = 0, totalUsados = 0;
  docs.forEach(rec => {
    if (rec.status !== 'disponivel') return; // só conta quem foi marcado Disponível naquele dia no Painel
    const pNorm = normPlaca(rec.plate);
    if (!placaAtiva.has(pNorm) || placaAtiva.get(pNorm) === false) return; // placa desconhecida/desativada no cadastro atual — fora da conta
    if (cidadesFiltroNorm) {
      // Reúne todos os "candidatos" de cidade/operação pra essa placa nesse
      // registro: (1) operação cadastrada na placa — fonte primária; (2)
      // cidades das viagens de fato realizadas por ela NAQUELE DIA; (3) cidades
      // já vistas pra essa placa em qualquer dia (fallback final, cobre
      // disponibilidade marcada num dia sem roteirização salva). Comparação
      // ignora acento/maiúscula pra não quebrar por diferença de digitação
      // entre o cadastro da placa e o cadastro do terminal (ex.: "Paulinia" x
      // "Paulínia").
      const candidatos = [
        placaOperacao.get(pNorm),
        ...(placaCidadePorDia.get(pNorm + '__' + rec.dateStr) || []),
        ...(placaCidade.get(pNorm) || []),
      ].filter(Boolean);
      const bate = candidatos.some(c => cidadesFiltroNorm.has(normCidade(c)));
      if (!bate) return; // nenhuma cidade/operação conhecida bate com o filtro selecionado
    }
    const usado = diasComViagemPorPlaca.has(pNorm + '__' + rec.dateStr);
    totalDisponibilizados++;
    if (usado) totalUsados++;
    const key = rec.carrier || '(sem transportadora)';
    if (!porTransportadora[key]) porTransportadora[key] = { transportadora: key, disponibilizados: 0, usados: 0 };
    porTransportadora[key].disponibilizados++;
    if (usado) porTransportadora[key].usados++;
  });
  // Diagnóstico: se um filtro de cidade está ativo e ZERO placas bateram,
  // loga uma amostra pra investigar direto do console do navegador (nomes
  // reais de operação/cidade cadastrados x nomes selecionados no filtro) —
  // sem isso, um "0/0" não diz se é falta de dado ou divergência de nome.
  if (cidadesFiltroNorm && totalDisponibilizados === 0 && docs.length) {
    const amostra = docs.filter(r => r.status === 'disponivel').slice(0, 5).map(r => {
      const pN = normPlaca(r.plate);
      return {
        placa: r.plate,
        data: r.dateStr,
        operacaoCadastrada: placaOperacao.get(pN) || null,
        cidadesDoDia: [...(placaCidadePorDia.get(pN + '__' + r.dateStr) || [])],
        cidadesConhecidas: [...(placaCidade.get(pN) || [])],
      };
    });
    console.warn('[Ociosidade] 0 placas bateram com o filtro de cidade selecionado. Filtro:', [...cidadesFiltro], 'Amostra de disponibilidade não-batida:', amostra);
  }
  const porTransportadoraArr = Object.values(porTransportadora)
    .map(t => ({ ...t, pctOciosidade: t.disponibilizados > 0 ? Math.round((1 - t.usados / t.disponibilizados) * 100) : 0 }))
    .sort((a, b) => b.pctOciosidade - a.pctOciosidade);
  return {
    totalDisponibilizados,
    totalUsados,
    pctOciosidade: totalDisponibilizados > 0 ? Math.round((1 - totalUsados / totalDisponibilizados) * 100) : 0,
    porTransportadora: porTransportadoraArr,
  };
}
let _dashUltimaOciosidade = { porTransportadora: [] };
let _dashOciosidadeDirecao = 'desc'; // 'desc' = maior primeiro, 'asc' = menor primeiro
function dashOciosidadeSetDirecao(direcao) {
  _dashOciosidadeDirecao = direcao;
  document.querySelectorAll('.dash-ociosidadedir-tab').forEach(b => {
    const ativo = b.dataset.dir === direcao;
    b.classList.toggle('active-rank', ativo);
    b.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    b.style.color = ativo ? '#000' : 'var(--text-2)';
  });
  dashRenderOciosidadeTransportadora(_dashUltimaOciosidade);
}
function dashRenderOciosidadeTransportadora(dados) {
  _dashUltimaOciosidade = dados;
  const box = document.getElementById('dash-ociosidade-transp');
  if (!box) return;
  const lista = [...(dados.porTransportadora || [])]
    .sort((a, b) => _dashOciosidadeDirecao === 'asc' ? a.pctOciosidade - b.pctOciosidade : b.pctOciosidade - a.pctOciosidade);
  if (!lista.length) {
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Nenhum dado de ociosidade para este período/filtro.</div>`;
    return;
  }
  const max = Math.max(...lista.map(t => t.pctOciosidade), 1);
  box.innerHTML = lista.map((t) => {
    const pct = Math.max(2, (t.pctOciosidade / max) * 100);
    return `
      <div style="display:flex;align-items:center;gap:12px;padding:10px 4px;border-bottom:1px solid var(--border-dk);">
        <div style="width:170px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px;font-weight:600;color:var(--text);" title="${t.transportadora}">${t.transportadora}</div>
        <div style="flex:1;background:rgba(255,255,255,.06);border-radius:6px;height:20px;position:relative;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:var(--pet-green,#b5e51d);border-radius:6px;transition:width .3s;"></div>
        </div>
        <div style="width:60px;flex-shrink:0;text-align:right;font-size:12.5px;font-weight:700;color:var(--text);">${t.pctOciosidade}%</div>
        <div style="width:170px;flex-shrink:0;text-align:right;font-size:10.5px;color:var(--text-3);">${t.usados} usados / ${t.disponibilizados} disponibilizados</div>
      </div>`;
  }).join('');
}
// ── Histórico por Veículo ────────────────────────────────────────────────────
// Tabela dia a dia de UM veículo: dia da semana, data, status marcado no
// Painel de Disponibilidade (Firestore) e km planejado nas roteirizações
// (reaproveita _dashUltimoAgregado, o mesmo cálculo já usado pela Jornada).
// Mesmo mapeamento de STATUS_OPTS do main.js — mantido aqui em separado
// porque não é exposto em window; se os status do cadastro mudarem lá,
// replicar aqui também.
const _DASH_HISTVEIC_STATUS = {
  disponivel:   { label: 'Disponível',           cor: '#4caf1f', bg: 'rgba(110,224,74,.16)',  borda: 'rgba(76,175,31,.35)' },
  indisponivel: { label: 'Indisponível',         cor: '#e23c3c', bg: 'rgba(240,96,96,.14)',   borda: 'rgba(226,60,60,.35)' },
  manutencao:   { label: 'Manutenção',           cor: '#c98a00', bg: 'rgba(240,190,64,.18)',  borda: 'rgba(201,138,0,.35)' },
  folga:        { label: 'Folga',                cor: '#3b7fd6', bg: 'rgba(112,168,240,.16)', borda: 'rgba(59,127,214,.35)' },
  programado:   { label: 'Programado/Em viagem', cor: '#8a4fd6', bg: 'rgba(176,126,240,.16)', borda: 'rgba(138,79,214,.35)' },
};
const _DASH_HISTVEIC_DIAS_SEMANA = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
let _dashHistVeiculoOpcoesPlacas = []; // lista bruta {placa, transportadora} — cacheada pra alimentar o filtro de transportadora sem reconsultar
let _dashHistVeiculoAguardandoPlates = false;
async function dashPopularSeletorHistVeiculo() {
  const sel = document.getElementById('dash-histveic-placa');
  if (!sel) return;
  // main.js carrega como módulo ES, em paralelo ao dashboard.js — pode não
  // ter terminado de expor window.dbGetPlates ainda quando o Dashboard
  // renderiza pela primeira vez. Sem essa espera, o seletor ficava preso
  // pra sempre em "Carregando placas...", porque a função só era chamada
  // de novo se algum filtro do Dashboard mudasse depois.
  if (typeof window.dbGetPlates !== 'function') {
    if (_dashHistVeiculoAguardandoPlates) return;
    _dashHistVeiculoAguardandoPlates = true;
    let tentativas = 0;
    const timer = setInterval(() => {
      tentativas++;
      if (typeof window.dbGetPlates === 'function') {
        clearInterval(timer);
        _dashHistVeiculoAguardandoPlates = false;
        dashPopularSeletorHistVeiculo();
      } else if (tentativas >= 50) { // ~10s
        clearInterval(timer);
        _dashHistVeiculoAguardandoPlates = false;
        sel.innerHTML = '<option value="">Não foi possível carregar as placas</option>';
      }
    }, 200);
    return;
  }
  try {
    const todasPlacas = await window.dbGetPlates();
    const opcoes = [];
    Object.entries(todasPlacas || {}).forEach(([transportadora, placas]) => {
      (placas || []).forEach(p => {
        const placa = (p && p.placa || '').trim().toUpperCase();
        if (placa) opcoes.push({ placa, transportadora });
      });
    });
    opcoes.sort((a, b) => a.placa.localeCompare(b.placa));
    _dashHistVeiculoOpcoesPlacas = opcoes;
    dashPopularSeletorTranspHistVeiculo();
    dashFiltrarPlacasHistVeiculo();
  } catch(e) {
    console.warn('[HistVeiculo] falha ao popular placas', e);
    sel.innerHTML = '<option value="">Erro ao carregar placas</option>';
  }
}
// Seletor de transportadora — só filtra a lista de placas já carregada, não
// reconsulta o Firestore (é o mesmo documento único, com todas as
// transportadoras dentro; separar por transportadora aqui é só organização
// visual pra achar a placa mais rápido, não economiza leitura).
function dashPopularSeletorTranspHistVeiculo() {
  const selT = document.getElementById('dash-histveic-transp');
  if (!selT) return;
  const transportadoras = Array.from(new Set(_dashHistVeiculoOpcoesPlacas.map(o => o.transportadora))).sort((a,b) => a.localeCompare(b, 'pt-BR'));
  const atual = selT.value;
  selT.innerHTML = '<option value="">Todas as transportadoras</option>' +
    transportadoras.map(t => `<option value="${t}">${t}</option>`).join('');
  if (atual && transportadoras.includes(atual)) selT.value = atual;
}
function dashFiltrarPlacasHistVeiculo() {
  const sel = document.getElementById('dash-histveic-placa');
  const selT = document.getElementById('dash-histveic-transp');
  if (!sel) return;
  const transpFiltro = selT ? selT.value : '';
  const filtradas = transpFiltro
    ? _dashHistVeiculoOpcoesPlacas.filter(o => o.transportadora === transpFiltro)
    : _dashHistVeiculoOpcoesPlacas;
  const atual = sel.value;
  sel.innerHTML = '<option value="">Selecione um veículo...</option>' +
    filtradas.map(o => `<option value="${o.placa}">${o.placa} — ${o.transportadora}</option>`).join('');
  if (atual && filtradas.some(o => o.placa === atual)) sel.value = atual;
}
// Lista de todos os dias (YYYY-MM-DD) — por padrão, entre o menor e o maior
// savedAt dos snapshots carregados (mesmo período do Dashboard), mas pode
// ser sobreposta pelo calendário "De/Até" da própria seção, pra ver um
// intervalo diferente sem precisar mudar o filtro geral do Dashboard.
function _dashHistVeiculoListaDias() {
  const elDe  = document.getElementById('dash-histveic-de');
  const elAte = document.getElementById('dash-histveic-ate');
  let min = elDe  && elDe.value  ? elDe.value  : null;
  let max = elAte && elAte.value ? elAte.value : null;
  if (!min || !max) return []; // sem os dois extremos definidos, não há o que listar
  if (min > max) { const tmp = min; min = max; max = tmp; } // datas trocadas, inverte
  const dias = [];
  const cursor = new Date(min + 'T00:00:00');
  const fim = new Date(max + 'T00:00:00');
  let seguranca = 0; // trava contra intervalo absurdamente longo por engano
  while (cursor <= fim && seguranca < 366) {
    dias.push(cursor.toISOString().slice(0, 10));
    cursor.setDate(cursor.getDate() + 1);
    seguranca++;
  }
  return dias;
}
// Chamada ao trocar de veículo no seletor — NÃO busca nada sozinha. A
// consulta ao Firestore só roda quando o usuário define o período (De/Até)
// e aperta "Buscar", de propósito: evita gerar leitura toda vez que algum
// filtro do Dashboard mudar (cidade, cliente, mês) com um veículo já
// selecionado, que era o padrão antigo — cada troca de filtro relia a
// consulta sozinha, mesmo sem o usuário pedir.
function dashSelecionarVeiculoHistVeiculo(placaEscolhida) {
  _dashHistVeiculoPlacaAtual = placaEscolhida || '';
  _dashHistVeiculoDados = [];
  const box = document.getElementById('dash-histveic-tabela');
  if (!box) return;
  if (!placaEscolhida) {
    box.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Selecione um veículo acima.</div>';
    return;
  }
  box.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Escolha o período (De/Até) e clique em "Buscar".</div>';
}
function dashLimparFiltroDataHistVeiculo() {
  const elDe  = document.getElementById('dash-histveic-de');
  const elAte = document.getElementById('dash-histveic-ate');
  if (elDe)  elDe.value  = '';
  if (elAte) elAte.value = '';
  const box = document.getElementById('dash-histveic-tabela');
  if (box) box.innerHTML = _dashHistVeiculoPlacaAtual
    ? '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Escolha o período (De/Até) e clique em "Buscar".</div>'
    : '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Selecione um veículo acima.</div>';
  _dashHistVeiculoDados = [];
}
// Disparo explícito, único ponto que efetivamente consulta o Firestore.
function dashBuscarHistoricoVeiculo() {
  const selPlaca = document.getElementById('dash-histveic-placa');
  const placa = selPlaca ? selPlaca.value : _dashHistVeiculoPlacaAtual;
  if (!placa) { showToast('Selecione um veículo antes de buscar.', false); return; }
  const elDe  = document.getElementById('dash-histveic-de');
  const elAte = document.getElementById('dash-histveic-ate');
  if (!elDe?.value || !elAte?.value) { showToast('Defina o período (De e Até) antes de buscar.', false); return; }
  dashCarregarHistoricoVeiculo(placa);
}
// Km planejado por dia pra uma placa — soma o km de ida e volta de todas as
// viagens daquela placa naquele dia (kmIda × 2), a partir do que a Jornada
// já calcula.
function _dashHistVeiculoKmPorDia(placaNorm) {
  const mapa = {};
  (_dashUltimoAgregado?.entradasTransportadora || []).forEach(e => {
    if (e._semViagem) return; // não é uma viagem de verdade
    if ((e.placa || '').trim().toUpperCase() !== placaNorm) return;
    mapa[e.data] = (mapa[e.data] || 0) + (e.kmIda || 0) * 2; // ida + volta
  });
  return mapa;
}
// Cache por placa — evita reconsultar TODO o histórico daquela placa toda vez
// que um filtro do Dashboard muda (cidade, cliente, mês...), já que
// dashRender() re-chama esta função sempre que há um veículo selecionado.
// TTL curto o bastante pra não esconder uma atualização de status por muito
// tempo, longo o bastante pra absorver uma sequência de cliques em filtros.
const _dashHistVeiculoCacheDocs = new Map(); // placa -> { docs, expiraEm }
const _DASH_HISTVEIC_CACHE_TTL_MS = 120_000; // 2 minutos
async function dashCarregarHistoricoVeiculo(placaEscolhida) {
  const box = document.getElementById('dash-histveic-tabela');
  if (!box) return;
  _dashHistVeiculoPlacaAtual = placaEscolhida || '';
  if (!placaEscolhida) {
    box.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Selecione um veículo acima.</div>';
    _dashHistVeiculoDados = [];
    return;
  }
  box.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Carregando...</div>';
  const pNorm = placaEscolhida.trim().toUpperCase();
  const dias = _dashHistVeiculoListaDias();
  if (!dias.length) {
    box.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Escolha o período (De/Até) e clique em "Buscar".</div>';
    _dashHistVeiculoDados = [];
    return;
  }
  const dataIni = dias[0], dataFim = dias[dias.length - 1];
  let statusPorDia = {};
  let erroBusca = null;
  if (window.fbDb && window.fbCollection && window.fbQuery && window.fbWhere && window.fbGetDocs) {
    try {
      let docs;
      const cacheado = _dashHistVeiculoCacheDocs.get(pNorm);
      if (cacheado && cacheado.expiraEm > Date.now()) {
        docs = cacheado.docs;
      } else {
        // Só filtra por placa no Firestore (filtro de igualdade simples, não
        // precisa de índice composto) — o intervalo de datas é aplicado aqui
        // em JS depois. Combinar "placa == X" com "data entre A e B" na
        // mesma consulta do Firestore exige um índice composto específico
        // que normalmente não existe por padrão; sem ele, a consulta falhava
        // silenciosamente (só no console) e a tela mostrava "sem registro"
        // pra tudo, mesmo tendo dado de verdade salvo.
        const q = window.fbQuery(
          window.fbCollection(window.fbDb, 'availability'),
          window.fbWhere('plate', '==', pNorm)
        );
        const snap = await window.fbGetDocs(q);
        docs = snap.docs.map(d => d.data());
        _dashHistVeiculoCacheDocs.set(pNorm, { docs, expiraEm: Date.now() + _DASH_HISTVEIC_CACHE_TTL_MS });
      }
      docs.forEach(dd => {
        // Normaliza a placa do próprio registro na comparação, por segurança
        // contra qualquer inconsistência de maiúscula/minúscula salva antes.
        if ((dd.plate || '').trim().toUpperCase() !== pNorm) return;
        if (dd.dateStr && dd.dateStr >= dataIni && dd.dateStr <= dataFim) {
          statusPorDia[dd.dateStr] = dd.status;
        }
      });
    } catch(e) {
      console.warn('[HistVeiculo] falha ao buscar status de disponibilidade', e);
      erroBusca = e.message || String(e);
    }
  } else {
    erroBusca = 'Firestore ainda não está pronto.';
  }
  // Se o usuário trocou de placa/filtro enquanto essa busca rodava, descarta
  // — evita "piscar" dado de uma placa antiga por cima da nova selecionada.
  if (_dashHistVeiculoPlacaAtual !== placaEscolhida) return;
  const kmPorDia = _dashHistVeiculoKmPorDia(pNorm);
  const linhas = dias.map(dstr => {
    const dt = new Date(dstr + 'T00:00:00');
    return {
      data: dstr,
      diaSemana: _DASH_HISTVEIC_DIAS_SEMANA[dt.getDay()],
      status: statusPorDia[dstr] || null,
      km: kmPorDia[dstr] || 0,
    };
  });
  _dashHistVeiculoDados = linhas;
  dashRenderHistoricoVeiculo(linhas, pNorm, erroBusca);
}
function dashRenderHistoricoVeiculo(linhas, placa, erro) {
  const box = document.getElementById('dash-histveic-tabela');
  if (!box) return;
  if (!linhas.length) {
    box.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Sem dados para este veículo no período.</div>';
    return;
  }
  const avisoErro = erro
    ? `<div style="background:rgba(240,96,96,.12);border:1px solid rgba(226,60,60,.35);border-radius:8px;padding:8px 12px;font-size:11px;color:#e23c3c;margin-bottom:10px;">⚠️ Não foi possível buscar o status de disponibilidade (${erro}). O km planejado abaixo continua correto, só o status ficou de fora.</div>`
    : '';
  const fmtDataCurta = dstr => { const [y,m,d] = dstr.split('-'); return `${d}/${m}`; };
  const chipStatus = (l) => {
    const st = l.status ? (_DASH_HISTVEIC_STATUS[l.status] || { label: l.status, cor: 'var(--text-2)', bg: 'rgba(255,255,255,.06)', borda: 'var(--border-dk)' }) : null;
    if (!st) return `<span style="display:inline-block;padding:3px 9px;border-radius:20px;font-size:10px;color:var(--text-3);border:1px dashed var(--border-dk);white-space:nowrap;">sem registro</span>`;
    return `<span style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:10.5px;font-weight:700;color:${st.cor};background:${st.bg};border:1px solid ${st.borda};white-space:nowrap;">${st.label}</span>`;
  };
  // Fim de semana com fundo levemente destacado, pra achar sábado/domingo
  // de relance numa tabela longa — mesmo truque visual da planilha original.
  const fundoDia = dstr => {
    const dow = new Date(dstr + 'T00:00:00').getDay();
    return (dow === 0 || dow === 6) ? 'background:rgba(255,255,255,.035);' : '';
  };
  const bordaCol = i => (i > 0 ? 'border-left:1px solid var(--border-dk);' : '');
  // Primeira coluna (rótulo da linha) fixa ao rolar na horizontal — sticky
  // exige fundo opaco (senão o conteúdo por baixo aparece por trás dela) e
  // um z-index maior que as demais células pra ficar sempre por cima.
  const colRotulo = 'position:sticky;left:0;z-index:2;padding:6px 8px;font-size:10px;font-weight:700;color:var(--text-3);text-transform:uppercase;white-space:nowrap;background:var(--surface);box-shadow:2px 0 4px -2px rgba(0,0,0,.15);';
  const colunas = linhas.length;
  box.innerHTML = `
    ${avisoErro}
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:8px;">
      <div style="font-size:11px;color:var(--text-3);">Veículo <strong style="color:var(--text);">${placa}</strong> · ${linhas.length} dia(s) no período</div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;">
        ${Object.values(_DASH_HISTVEIC_STATUS).map(st => `<span style="display:flex;align-items:center;gap:4px;font-size:10px;color:var(--text-3);"><span style="width:8px;height:8px;border-radius:50%;background:${st.cor};display:inline-block;"></span>${st.label}</span>`).join('')}
      </div>
    </div>
    <div style="overflow-x:auto;border:0.5px solid var(--border-dk);border-radius:8px;">
      <table style="border-collapse:collapse;width:100%;min-width:${colunas * 62}px;">
        <tbody>
          <tr style="border-bottom:1px solid var(--border-dk);">
            <td style="${colRotulo}">Dia semana</td>
            ${linhas.map((l,i)=>`<td style="padding:6px 4px;text-align:center;font-size:11px;color:var(--text-3);${bordaCol(i)}${fundoDia(l.data)}">${l.diaSemana}</td>`).join('')}
          </tr>
          <tr style="border-bottom:1px solid var(--border-dk);">
            <td style="${colRotulo}">Data</td>
            ${linhas.map((l,i)=>`<td style="padding:6px 4px;text-align:center;font-size:11px;font-weight:700;color:var(--text);${bordaCol(i)}${fundoDia(l.data)}">${fmtDataCurta(l.data)}</td>`).join('')}
          </tr>
          <tr style="border-bottom:1px solid var(--border-dk);">
            <td style="${colRotulo}">Status</td>
            ${linhas.map((l,i)=>`<td style="padding:6px 4px;text-align:center;${bordaCol(i)}${fundoDia(l.data)}">${chipStatus(l)}</td>`).join('')}
          </tr>
          <tr>
            <td style="${colRotulo}">Km do dia</td>
            ${linhas.map((l,i)=>`<td style="padding:6px 4px;text-align:center;font-size:11.5px;font-weight:700;color:${l.km ? 'var(--pet-green,#b5e51d)' : 'var(--text-3)'};${bordaCol(i)}${fundoDia(l.data)}">${l.km?Math.round(l.km).toLocaleString('pt-BR'):'—'}</td>`).join('')}
          </tr>
        </tbody>
      </table>
    </div>`;
}
function dashExportarHistoricoVeiculoCSV() {
  if (!_dashHistVeiculoDados.length || !_dashHistVeiculoPlacaAtual) {
    showToast('Selecione um veículo com dados carregados antes de exportar.', false);
    return;
  }
  const linhasCsv = [
    ['Placa', 'Data', 'Dia da Semana', 'Status', 'Km do Dia (ida+volta)'].join(';'),
    ..._dashHistVeiculoDados.map(l => {
      const st = l.status ? (_DASH_HISTVEIC_STATUS[l.status] || { label: l.status }) : { label: '' };
      return [_dashHistVeiculoPlacaAtual, l.data, l.diaSemana, st.label, l.km ? Math.round(l.km) : ''].join(';');
    }),
  ];
  const csv = '\uFEFF' + linhasCsv.join('\r\n'); // BOM pro Excel abrir acentos certinho
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `historico_${_dashHistVeiculoPlacaAtual}_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function dashExportarHistoricoVeiculoXLSX() {
  if (!_dashHistVeiculoDados.length || !_dashHistVeiculoPlacaAtual) {
    showToast('Selecione um veículo com dados carregados antes de exportar.', false);
    return;
  }
  if (typeof XLSX === 'undefined') {
    showToast('Biblioteca de Excel não carregada — tente recarregar a página.', false);
    return;
  }
  const linhas = _dashHistVeiculoDados.map(l => {
    const st = l.status ? (_DASH_HISTVEIC_STATUS[l.status] || { label: l.status }) : { label: '' };
    return {
      'Placa': _dashHistVeiculoPlacaAtual,
      'Data': l.data,
      'Dia da Semana': l.diaSemana,
      'Status': st.label,
      'Km do Dia (ida+volta)': l.km ? Math.round(l.km) : '',
    };
  });
  const ws = XLSX.utils.json_to_sheet(linhas);
  ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 22 }, { wch: 20 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Histórico');
  XLSX.writeFile(wb, `historico_${_dashHistVeiculoPlacaAtual}_${new Date().toISOString().slice(0,10)}.xlsx`);
}
let _dashRankingOrdem = 'volume'; // 'volume' | 'km' | 'custo'
let _dashRankingDirecao = 'desc'; // 'desc' = maior primeiro, 'asc' = menor primeiro
// ── Zoom dos gráficos por transportadora (botão maior/menor de TAMANHO) ─────
// Usa a propriedade CSS "zoom" direto no container — escala bar, texto e
// tudo dentro de uma vez, sem precisar tocar nas funções de render. Fica
// guardado por box (cada gráfico lembra o zoom escolhido) e sobrevive a
// re-renderizações porque o zoom fica no próprio elemento container (que
// nunca é substituído, só o innerHTML dele).
const _dashChartZoomNiveis = [0.75, 0.85, 1, 1.15, 1.3, 1.5];
const _dashChartZoomAtual = {}; // boxId -> índice em _dashChartZoomNiveis
function dashChartZoom(boxId, direcao) {
  const box = document.getElementById(boxId);
  if (!box) return;
  const atual = _dashChartZoomAtual[boxId] ?? _dashChartZoomNiveis.indexOf(1);
  const novo = Math.min(_dashChartZoomNiveis.length - 1, Math.max(0, atual + direcao));
  _dashChartZoomAtual[boxId] = novo;
  box.style.zoom = _dashChartZoomNiveis[novo];
}
function dashRankingSetOrdem(campo) {
  _dashRankingOrdem = campo;
  document.querySelectorAll('.dash-rank-tab').forEach(b => {
    const ativo = b.dataset.campo === campo;
    b.classList.toggle('active-rank', ativo);
    b.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    b.style.color = ativo ? '#000' : 'var(--text-2)';
  });
  dashRenderRankingTransportadoras(_dashUltimoRanking || []);
}
function dashRankingSetDirecao(direcao) {
  _dashRankingDirecao = direcao;
  document.querySelectorAll('.dash-rankdir-tab').forEach(b => {
    const ativo = b.dataset.dir === direcao;
    b.classList.toggle('active-rank', ativo);
    b.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    b.style.color = ativo ? '#000' : 'var(--text-2)';
  });
  dashRenderRankingTransportadoras(_dashUltimoRanking || []);
}
let _dashUltimoRanking = [];
function dashRenderRankingTransportadoras(lista) {
  _dashUltimoRanking = lista;
  const box = document.getElementById('dash-ranking-transp');
  if (!box) return;
  if (!lista.length) {
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Nenhum dado de transportadora para este período/filtro.</div>`;
    return;
  }
  const campo = _dashRankingOrdem;
  const ordenado = [...lista].sort((a, b) => _dashRankingDirecao === 'asc' ? a[campo] - b[campo] : b[campo] - a[campo]);
  const max = Math.max(...ordenado.map(t => t[campo]), 0.001);
  // Medalha só faz sentido no topo quando está ordenado do maior pro menor
  const medalhas = _dashRankingDirecao === 'desc' ? ['🥇', '🥈', '🥉'] : [];
  const fmtValor = (t) => {
    if (campo === 'volume') return t.volume.toFixed(1) + ' m³';
    if (campo === 'km')     return Math.round(t.km).toLocaleString('pt-BR') + ' km';
    return t.temContrato ? 'R$ ' + t.custo.toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2}) : '—';
  };
  const cores = { volume: '#f0be40', km: '#6ee04a', custo: '#70a8f0' };
  box.innerHTML = ordenado.map((t, i) => {
    const val = t[campo];
    const pct = max > 0 ? Math.max(2, (val / max) * 100) : 2;
    const rank = i < medalhas.length ? medalhas[i] : `#${i+1}`;
    const semContratoTag = (campo === 'custo' && !t.temContrato) ? `<span title="Sem contrato cadastrado na aba Frete para nenhuma placa desta transportadora" style="font-size:9px;color:#F59E0B;background:rgba(245,158,11,.12);border:1px solid rgba(245,158,11,.3);border-radius:4px;padding:1px 6px;margin-left:6px;">sem contrato</span>` : '';
    return `
      <div style="display:flex;align-items:center;gap:12px;padding:10px 4px;border-bottom:1px solid var(--border-dk);">
        <div style="width:28px;flex-shrink:0;text-align:center;font-size:${i<medalhas.length?'16px':'12px'};font-weight:700;color:${i<medalhas.length?'inherit':'var(--text-3)'};">${rank}</div>
        <div style="width:170px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px;font-weight:600;color:var(--text);" title="${t.transportadora}">${t.transportadora}${semContratoTag}</div>
        <div style="flex:1;background:rgba(255,255,255,.06);border-radius:6px;height:20px;position:relative;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${cores[campo]};border-radius:6px;transition:width .3s;"></div>
        </div>
        <div style="width:110px;flex-shrink:0;text-align:right;font-size:12.5px;font-weight:700;color:var(--text);">${fmtValor(t)}</div>
        <div style="width:110px;flex-shrink:0;text-align:right;font-size:10.5px;color:var(--text-3);">${t.viagens} viag. · ${t.nPlacas} placa${t.nPlacas>1?'s':''}</div>
      </div>`;
  }).join('');
}
window.dashRankingSetOrdem = dashRankingSetOrdem;
window.dashRankingSetDirecao = dashRankingSetDirecao;
window.dashJornadaSetDirecao = dashJornadaSetDirecao;
window.dashOciosidadeSetDirecao = dashOciosidadeSetDirecao;
window.dashChartZoom = dashChartZoom;
window.dashSelecionarVeiculoHistVeiculo = dashSelecionarVeiculoHistVeiculo;
window.dashFiltrarPlacasHistVeiculo = dashFiltrarPlacasHistVeiculo;
window.dashBuscarHistoricoVeiculo = dashBuscarHistoricoVeiculo;
window.dashLimparFiltroDataHistVeiculo = dashLimparFiltroDataHistVeiculo;
window.dashExportarHistoricoVeiculoCSV = dashExportarHistoricoVeiculoCSV;
window.dashExportarHistoricoVeiculoXLSX = dashExportarHistoricoVeiculoXLSX;
// ── Renderizar Dashboard ───────────────────────────────────────────────────
// ── Filtro de clientes ────────────────────────────────────────────────────
let _dashClientesSelecionados = null; // null = todos; Set = filtro ativo
let _dashSegmentosSelecionados = null; // GLOBAL (botão do topo) — null = todos; Set = filtro ativo (B2B/B2C/TRR/'—'). Afeta TODA a tela (KPIs, gráficos, tabela, mapa, Excel, e também o card "Clientes sem comprar" abaixo).
let _dashSegmentosInativosSelecionados = null; // LOCAL (botão dentro do card "Clientes sem comprar") — null = todos; Set = filtro ativo. Afeta SÓ aquele card, não mexe no resto da tela.
const _DASH_SEGMENTOS_TODOS = ['B2B', 'B2C', 'TRR', '—']; // '—' = cliente sem segmento cadastrado
// ── Mapa nome → Código SAP ──────────────────────────────────────────────────
// dashClientesEfetivos()/dashRenderClientesInativosUI() só têm o NOME do
// cliente pra trabalhar (a lista de nomes do filtro já existia assim antes
// do Segmento existir, e mudar isso mexeria no filtro de Clientes que já
// funciona). Esse mapa é o jeito de "lembrar" qual SAP corresponde a qual
// nome, preenchido toda vez que dashAgregar()/dashCalcularClientesInativos()
// rodam (que SIM têm o pedido completo, com codigoSAP) — assim
// dashClienteSegmento(nome) consegue montar a MESMA chave que
// dashChaveCliente() usaria (SAP primeiro, nome normalizado como
// fallback), garantindo que "mesmo cliente" aqui é exatamente "mesmo
// cliente" no resto do Dashboard.
let _dashMapaNomeParaSAP = {};
function _dashAtualizarMapaNomeSAP(itens) {
  (itens || []).forEach(it => {
    if (it && it.nome && it.codigoSAP) _dashMapaNomeParaSAP[it.nome] = it.codigoSAP;
  });
}
// ── Segmento do cliente (B2B/B2C/TRR) ───────────────────────────────────────
// Campo vive no CADASTRO de cliente (aba Cadastros → Clientes), não no
// pedido salvo no histórico — então sempre consulta o cadastro ATUAL
// (window.clientes, populado pelo script do Roteirizador, ver comentário
// "var no nível de topo vira window.X" perto da declaração de `clientes`).
// Isso significa que o segmento aplicado é sempre o de HOJE, mesmo pra
// pedidos antigos — não existe "segmento histórico" por pedido, e não faz
// muito sentido ter (é um atributo do cliente, não da entrega).
//
// Usa a MESMA chave de identidade que dashChaveCliente() já usa pra
// agrupar entregas por cliente em todo o resto do Dashboard: Código SAP
// primeiro (exato, sem ambiguidade), nome normalizado (sem acento, sem
// sufixo jurídico como "LTDA", maiúsculo, sem espaço extra) como fallback
// só quando não há SAP. Antes comparava só o nome exato — bastava
// qualquer diferença de maiúscula/sufixo/acento entre o nome salvo no
// pedido histórico e o nome no cadastro pra nunca bater, e TODO cliente
// caía em "sem segmento" mesmo já tendo sido cadastrado.
function _dashSegmentoLookup() {
  const arr = (typeof clientes !== 'undefined' && clientes) || window.clientes || [];
  const map = {};
  arr.forEach(c => {
    if (!c || !c.segmento) return;
    const key = dashChaveCliente({ codigoSAP: c.codigoSAP, cliente: c.nome });
    map[key] = c.segmento;
  });
  return map;
}
function dashClienteSegmento(nome) {
  const sap = _dashMapaNomeParaSAP[nome] || '';
  const key = dashChaveCliente({ codigoSAP: sap, cliente: nome });
  return _dashSegmentoLookup()[key] || '';
}
// ── Conjunto efetivo de clientes visíveis (filtro GLOBAL) ──────────────────
// Combina (E lógico, não OU) o filtro do picker de Clientes com o filtro de
// Segmento GLOBAL (botão do topo) — os dois se cruzam. Ex.: picker com
// "Posto X" selecionado, mas segmento filtrado só em "TRR": se Posto X não
// for TRR, ele fica de fora mesmo estando no picker. Retorna null quando
// NENHUM dos dois filtros está ativo (comportamento idêntico ao "sem
// filtro" de antes, sem custo extra). NÃO usa o filtro local do card
// "Clientes sem comprar" — esse é aplicado só ali, ver
// dashRenderClientesInativosUI().
function dashClientesEfetivos(nomesBase) {
  if (!_dashClientesSelecionados && !_dashSegmentosSelecionados) return null;
  const permitido = (nome) => {
    if (_dashClientesSelecionados && !_dashClientesSelecionados.has(nome)) return false;
    if (_dashSegmentosSelecionados) {
      const seg = dashClienteSegmento(nome) || '—';
      if (!_dashSegmentosSelecionados.has(seg)) return false;
    }
    return true;
  };
  return new Set(nomesBase.filter(permitido));
}
const _DASH_SEGMENTO_OPCOES = [
  { valor: 'B2B', label: 'B2B' },
  { valor: 'B2C', label: 'B2C' },
  { valor: 'TRR', label: 'TRR' },
  { valor: '—',   label: 'Sem segmento cadastrado' },
];
// ── Dropdown de Segmento — mesmo padrão visual/comportamento do filtro de
// Clientes (botão + painel com checkbox + Todos/Limpar/Aplicar), só que com
// 4 opções fixas em vez de lista dinâmica. Existem DUAS instâncias na tela,
// cada uma com seu PRÓPRIO estado — de propósito, não são a mesma coisa:
//   • "dash-seg-panel"          (topo)  → _dashSegmentosSelecionados         → filtra a tela inteira
//   • "dash-seg-inativos-panel" (card)  → _dashSegmentosInativosSelecionados → filtra só o card "Clientes sem comprar"
function _dashSegEstadoDoPainel(painelId) {
  return painelId === 'dash-seg-inativos-panel' ? _dashSegmentosInativosSelecionados : _dashSegmentosSelecionados;
}
function _dashSegmentoListHtml(painelId) {
  const estado = _dashSegEstadoDoPainel(painelId);
  return _DASH_SEGMENTO_OPCOES.map(o => {
    const checked = !estado || estado.has(o.valor);
    return `<div data-seg="${o.valor}" data-checked="${checked ? 1 : 0}" onclick="_dashSegToggleRow(this)"
      style="display:flex;align-items:center;gap:9px;padding:8px 14px;cursor:pointer;font-size:12.5px;color:#111827;"
      onmouseover="this.style.background='rgba(0,0,0,.03)'" onmouseout="this.style.background='none'">
      <span class="dash-seg-checkbox" style="width:16px;height:16px;border-radius:4px;border:1.5px solid ${checked ? 'var(--pet-green,#b5e51d)' : '#bbb'};background:${checked ? 'var(--pet-green,#b5e51d)' : 'transparent'};display:flex;align-items:center;justify-content:center;flex-shrink:0;">${checked ? _dashCheckSVG() : ''}</span>
      ${o.label}
    </div>`;
  }).join('');
}
// IMPORTANTE: onclick="..." no HTML sempre resolve a função no escopo
// GLOBAL (window), nunca dentro da IIFE — mesmo declarada com `function`
// aqui dentro, sem window.X = X ela fica invisível pro atributo onclick, e
// o clique não faz NADA (erro silencioso, sem aviso na tela). Foi o bug do
// checkbox não marcar: a função existia, só não estava exposta.
function _dashSegToggleRow(el) {
  const novo = el.dataset.checked !== '1';
  el.dataset.checked = novo ? '1' : '0';
  const box = el.querySelector('.dash-seg-checkbox');
  box.style.borderColor = novo ? 'var(--pet-green,#b5e51d)' : '#bbb';
  box.style.background  = novo ? 'var(--pet-green,#b5e51d)' : 'transparent';
  box.innerHTML = novo ? _dashCheckSVG() : '';
}
window._dashSegToggleRow = _dashSegToggleRow;
function dashTogglePainelSegmento(painelId) {
  const panel = document.getElementById(painelId);
  if (!panel) return;
  const visible = panel.style.display !== 'none';
  panel.style.display = visible ? 'none' : 'flex';
  if (!visible) {
    // Repopula toda vez que abre — reflete o estado ATUAL DESTE painel
    // especificamente (cada painel tem o seu, ver _dashSegEstadoDoPainel).
    const list = panel.querySelector('.dash-seg-list-inner');
    if (list) list.innerHTML = _dashSegmentoListHtml(painelId);
  }
}
function dashSelecionarTodosSegmento(painelId, sel) {
  const panel = document.getElementById(painelId);
  if (!panel) return;
  panel.querySelectorAll('[data-seg]').forEach(el => {
    el.dataset.checked = sel ? '1' : '0';
    const box = el.querySelector('.dash-seg-checkbox');
    box.style.borderColor = sel ? 'var(--pet-green,#b5e51d)' : '#bbb';
    box.style.background  = sel ? 'var(--pet-green,#b5e51d)' : 'transparent';
    box.innerHTML = sel ? _dashCheckSVG() : '';
  });
}
function _dashAtualizarBadgeSegmento(badgeId, estado) {
  const badge = document.getElementById(badgeId);
  if (!badge) return;
  if (estado) {
    badge.textContent = estado.size;
    badge.style.display = '';
  } else {
    badge.style.display = 'none';
  }
}
function dashAplicarFiltroSegmento(painelId) {
  const panel = document.getElementById(painelId);
  if (!panel) return;
  const selecionados = new Set();
  panel.querySelectorAll('[data-seg][data-checked="1"]').forEach(el => selecionados.add(el.dataset.seg));
  const novoEstado = selecionados.size === _DASH_SEGMENTOS_TODOS.length ? null : selecionados;
  panel.style.display = 'none';
  if (painelId === 'dash-seg-inativos-panel') {
    // Filtro LOCAL — só o card "Clientes sem comprar". Não toca no resto
    // da tela (não chama dashRenderComFiltro).
    _dashSegmentosInativosSelecionados = novoEstado;
    _dashAtualizarBadgeSegmento('dash-seg-inativos-badge', novoEstado);
    dashRenderClientesInativosUI(_dashUltimaListaInativos);
  } else {
    // Filtro GLOBAL — a tela inteira, incluindo o card "Clientes sem
    // comprar" (ver comentário no topo do arquivo: "filtra tudo").
    _dashSegmentosSelecionados = novoEstado;
    _dashAtualizarBadgeSegmento('dash-seg-badge', novoEstado);
    dashRenderComFiltro(); // reaplica sobre os snapshots já carregados, sem reler o disco
    dashRenderClientesInativosUI(_dashUltimaListaInativos);
  }
}
window.dashTogglePainelSegmento  = dashTogglePainelSegmento;
window.dashSelecionarTodosSegmento = dashSelecionarTodosSegmento;
window.dashAplicarFiltroSegmento = dashAplicarFiltroSegmento;
// Fecha os painéis de Segmento ao clicar fora (mesmo padrão dos outros
// dropdowns da tela — Clientes, Operação, Ferramentas).
document.addEventListener('click', function(e) {
  ['dash-seg-panel', 'dash-seg-inativos-panel'].forEach(painelId => {
    const btnId = painelId.replace('-panel', '-btn');
    const panel = document.getElementById(painelId);
    const btn   = document.getElementById(btnId);
    if (panel && panel.style.display !== 'none' && !panel.contains(e.target) && !btn?.contains(e.target)) {
      panel.style.display = 'none';
    }
  });
});
let _dashCidadesSelecionadas  = null; // null = todas as cidades de operação; Set = filtro ativo
let _dashTodasCidades         = [];   // lista completa de cidades de operação do período
let _dashSnapshotsAtivos = [];        // snapshots atualmente carregados
let _dashUltimoAgregado = null;       // último retorno de dashAgregar() — reaproveitado pelo Histórico por Veículo
let _dashHistVeiculoDados = [];       // última tabela renderizada do Histórico por Veículo — usada pelo botão de exportar CSV
let _dashHistVeiculoPlacaAtual = '';  // placa selecionada no momento (pra re-renderizar quando o filtro do Dashboard mudar)
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

// ── Filtro de Cidade da Operação ─────────────────────────────────────────────
// Mesmo padrão visual/funcional do filtro de clientes acima, mas filtra pela
// cidade cadastrada no TERMINAL (base), não pelo nome da base em si — ex.:
// "Betim POTENCIAL Nexta" e "Betim RBZ" são bases diferentes, mas a mesma
// cidade de operação "Betim". Esse filtro atua na FONTE dos dados (dentro de
// dashAgregar), então todo KPI/gráfico/tabela reflete só a(s) cidade(s)
// escolhida(s) — não é um filtro só de exibição.
function dashToggleFiltroCidades() {
  const panel = document.getElementById('dash-cid-panel');
  if (!panel) return;
  const visible = panel.style.display !== 'none';
  panel.style.display = visible ? 'none' : 'flex';
  if (!visible) {
    document.getElementById('dash-cid-search').value = '';
    dashFiltrarListaCidades('');
  }
}
function dashPopularListaCidades() {
  const list = document.getElementById('dash-cid-list');
  if (!list) return;
  list.innerHTML = _dashTodasCidades.map(nome => {
    const sel      = !_dashCidadesSelecionadas || _dashCidadesSelecionadas.has(nome);
    const nomeSafe = nome.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
    const chk      = sel ? _dashCheckSVG() : '';
    const bc       = sel ? 'var(--pet-green,#b5e51d)' : '#bbb';
    const bg       = sel ? 'var(--pet-green,#b5e51d)' : 'transparent';
    return `<div data-cid="${nomeSafe}" data-checked="${sel ? '1' : '0'}"
      style="display:flex;align-items:center;gap:10px;padding:8px 14px;cursor:pointer;border-radius:6px;margin:0 4px;user-select:none;">
      <span class="dash-cb-box" style="flex-shrink:0;width:20px;height:20px;border-radius:5px;border:2px solid ${bc};background:${bg};display:flex;align-items:center;justify-content:center;transition:all .12s;">${chk}</span>
      <span style="font-size:12px;color:var(--text,#111);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;">${nome}</span>
    </div>`;
  }).join('');
  list.querySelectorAll('div[data-cid]').forEach(row => {
    row.addEventListener('mouseenter', () => row.style.background = 'rgba(0,0,0,0.04)');
    row.addEventListener('mouseleave', () => row.style.background = '');
    row.addEventListener('click', () => {
      const checked = row.dataset.checked !== '1';
      row.dataset.checked = checked ? '1' : '0';
      _dashAtualizarBoxVisual(row.querySelector('.dash-cb-box'), checked);
    });
  });
}
function dashSelecionarTodosCidadesVisual(sel) {
  const list = document.getElementById('dash-cid-list');
  if (!list) return;
  list.querySelectorAll('div[data-cid]').forEach(row => {
    if (row.style.display === 'none') return;
    row.dataset.checked = sel ? '1' : '0';
    _dashAtualizarBoxVisual(row.querySelector('.dash-cb-box'), sel);
  });
}
function dashFiltrarListaCidades(busca) {
  const list = document.getElementById('dash-cid-list');
  if (!list) return;
  const b = (busca || '').toLowerCase();
  list.querySelectorAll('div[data-cid]').forEach(row => {
    row.style.display = row.dataset.cid.toLowerCase().includes(b) ? '' : 'none';
  });
}
function dashSelecionarTodosCidades(sel) { dashSelecionarTodosCidadesVisual(sel); }
function dashAplicarFiltroCidades() {
  const list = document.getElementById('dash-cid-list');
  if (!list) return;
  const selecionadas = new Set();
  const _dec = s => s.replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>');
  list.querySelectorAll('div[data-checked="1"]').forEach(row => selecionadas.add(_dec(row.dataset.cid)));
  _dashCidadesSelecionadas = selecionadas.size === _dashTodasCidades.length ? null : selecionadas;
  const badge = document.getElementById('dash-cid-badge');
  if (badge) {
    if (_dashCidadesSelecionadas) { badge.textContent = selecionadas.size; badge.style.display = ''; }
    else badge.style.display = 'none';
  }
  document.getElementById('dash-cid-panel').style.display = 'none';
  dashRenderComFiltro();
}
document.addEventListener('click', function(e) {
  const panel = document.getElementById('dash-cid-panel');
  const btn   = document.getElementById('dash-cid-btn');
  if (panel && panel.style.display !== 'none' && !panel.contains(e.target) && !btn?.contains(e.target)) {
    panel.style.display = 'none';
  }
});
window.dashToggleFiltroCidades   = dashToggleFiltroCidades;
window.dashFiltrarListaCidades    = dashFiltrarListaCidades;
window.dashSelecionarTodosCidades = dashSelecionarTodosCidades;
window.dashAplicarFiltroCidades   = dashAplicarFiltroCidades;

function dashRender(snapshots) {
  _dashSnapshotsAtivos = snapshots || [];
  window._dashSnapshotsAtivos = _dashSnapshotsAtivos; // acessível pra funções fora deste IIFE (ex.: dashDiagnosticarPedagioHoje)
  if (!snapshots || !snapshots.length) {
    document.querySelectorAll('#dk-viagens,#dk-entregas,#dk-volume,#dk-ocup,#dk-km,#dk-clientes,#dk-jornada,#dk-ociosidade,#dk-drop-entregas,#dk-drop-volume,#dk-perc-pedagiadas,#dk-tempo-pedagio')
      .forEach(el => { if(el) el.textContent = '-'; });
    const _elJH = document.getElementById('dk-jornada-horas'); if (_elJH) _elJH.textContent = '';
    const _elOQ = document.getElementById('dk-ociosidade-qtd'); if (_elOQ) _elOQ.textContent = '';
    document.getElementById('dash-tabela-cli-body').innerHTML =
      '<tr><td colspan="6" style="color:var(--text-3);text-align:center;padding:32px;">Nenhum dado para este período</td></tr>';
    return;
  }
  // Atualiza lista global de cidades de operação (pro filtro) — olha TODOS os
  // terminais dos snapshots carregados, mesmo antes de aplicar o filtro,
  // senão o próprio filtro nunca teria opções pra mostrar.
  const _novaListaCidades = Array.from(new Set(
    snapshots.flatMap(s => (s.terminais || []).map(t => t.cidade).filter(Boolean))
  )).sort((a,b) => a.localeCompare(b, 'pt-BR'));
  const _listaCidIgual = _novaListaCidades.length === _dashTodasCidades.length &&
    _novaListaCidades.every((n,i) => n === _dashTodasCidades[i]);
  if (!_listaCidIgual) {
    _dashTodasCidades = _novaListaCidades;
    if (_dashCidadesSelecionadas) {
      const novasCid = new Set(_dashTodasCidades);
      const filtroAtualizadoCid = new Set([..._dashCidadesSelecionadas].filter(n => novasCid.has(n)));
      _dashCidadesSelecionadas = filtroAtualizadoCid.size === _dashTodasCidades.length ? null : filtroAtualizadoCid;
    }
  }
  dashPopularListaCidades();
  // Filtro de cidade da operação aplicado NA FONTE (dentro de dashAgregar) —
  // por isso todo o resto do dashboard (KPIs, gráficos, mapa, ranking) já sai
  // filtrado corretamente, sem precisar re-filtrar depois.
  const d = dashAgregar(snapshots, _dashCidadesSelecionadas);
  _dashAtualizarMapaNomeSAP(d.clientes); // alimenta a busca de Segmento por SAP (ver dashClienteSegmento)
  _dashUltimoAgregado = d; // reaproveitado pelo Histórico por Veículo (km por dia), sem recalcular
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
  // Conjunto efetivo (picker de Clientes ∩ filtro de Segmento) — ver
  // dashClientesEfetivos(). null = nenhum dos dois filtros ativo.
  const _efetivos = dashClientesEfetivos(_novaListaClientes.length ? _novaListaClientes : _dashTodosClientes);
  // Aplica filtro se ativo
  const clientesFiltrados = _efetivos
    ? d.clientes.filter(c => _efetivos.has(c.nome))
    : d.clientes;
  const ocupFiltrados = _efetivos
    ? d.clientes_ocup.filter(c => _efetivos.has(c.nome))
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
  // As duas re-varreduras abaixo (ocupação e viagens) respeitam TAMBÉM o
  // filtro de cidade ativo (_dashCidadesSelecionadas), já que elas iteram os
  // snapshots crus (não passam por dashAgregar).
  let _kpiOcup = d.totalOcup;
  if (_efetivos) {
    let _filtVol = 0, _filtCap = 0;
    const _nomesF = _efetivos;
    _dashSnapshotsAtivos.forEach((snap, sIdx) => {
      const res = snap.resultado || {}, vecs = snap.veiculos || [], terms = snap.terminais || [];
      vecs.forEach(v => {
        const capV = v.capacidade || v.capacidadeTotal || 0;
        const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
        viagens.forEach(vi => {
          if (_dashCidadesSelecionadas && !_dashCidadesSelecionadas.has(dashCidadeOperacaoViagem(vi, v, terms))) return;
          // Verifica se esta viagem atende ao menos um cliente filtrado
          const atendeCliente = vi.paradas.some(par => {
            const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
            return _nomesF.has(n);
          });
          if (!atendeCliente) return;
          // Capacidade: soma capV desta viagem (veículo de 35m³ × 2 viagens = 70m³)
          if (capV > 0) _filtCap += capV;
          // Volume: apenas paradas dos clientes filtrados nesta viagem
          vi.paradas.forEach(par => {
            const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
            if (_nomesF.has(n)) _filtVol += par.volumeTotal || 0;
          });
        });
      });
    });
    _kpiOcup = _filtCap > 0 ? Math.round((_filtVol / _filtCap) * 100) : d.totalOcup;
  }
  // Viagens: conta apenas viagens que atendem ao menos um cliente filtrado
  const _nomesFilter = _efetivos;
  let _kpiViagens = d.totalViagens;
  let _kpiViagensComPedagio = d.totalViagensComPedagio;
  if (_nomesFilter) {
    _kpiViagens = 0;
    _kpiViagensComPedagio = 0;
    _dashSnapshotsAtivos.forEach(snap => {
      const res  = snap.resultado || {};
      const vecs = snap.veiculos  || [];
      const terms = snap.terminais || [];
      vecs.forEach(v => {
        (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length).forEach(vi => {
          const termNomeViagem = vi.terminalOrigem || vi.paradas?.find(p => p.pedido?.terminal)?.pedido?.terminal || v.terminal;
          const term = terms.find(t => t.nome === termNomeViagem);
          const tLat = term?.lat, tLon = term?.lon;
          if (_dashCidadesSelecionadas && !_dashCidadesSelecionadas.has(dashCidadeOperacaoViagem(vi, v, terms))) return;
          const temCliente = vi.paradas.some(par => {
            const nome = (par.pedido||{}).cliente || (par.pedido||{}).nomeCliente || par.nome || '';
            return _nomesFilter.has(nome);
          });
          if (temCliente) {
            _kpiViagens++;
            // Mesma detecção (linha reta, raioKm=3) usada em dashAgregar — ver
            // comentário lá pra detalhes/limitações. Só recalcula aqui porque
            // o filtro de cliente muda quais viagens contam, e esse filtro é
            // aplicado depois do dashAgregar já ter rodado.
            const pontosPedagio = [];
            if (NextaKm.coordenadaValida(tLat, tLon)) {
              pontosPedagio.push({ lat: parseFloat(tLat), lon: parseFloat(tLon) });
            }
            vi.paradas.forEach(par => {
              const coords = latLonEfetivo ? latLonEfetivo(par.pedido) : { lat: par.lat, lon: par.lon };
              const lat = coords?.lat ?? par.lat, lon = coords?.lon ?? par.lon;
              if (NextaKm.coordenadaValida(lat, lon)) {
                pontosPedagio.push({ lat: parseFloat(lat), lon: parseFloat(lon) });
              }
            });
            if (pontosPedagio.length >= 2) {
              const _eixosV = v.eixos || 2;
              const _pd = (typeof detectarPedagiosNaRota === 'function') ? detectarPedagiosNaRota(pontosPedagio, _eixosV, 3) : [];
              if (_pd.length > 0) _kpiViagensComPedagio++;
            }
          }
        });
      });
    });
  }
  set('dk-viagens',  _kpiViagens);
  set('dk-entregas', _kpiEntregas);
  set('dk-volume',   parseFloat(_kpiVol.toFixed(1)).toLocaleString('pt-BR', {minimumFractionDigits:0, maximumFractionDigits:1}) + ' m³');
  set('dk-ocup',     _kpiOcup + '%');
  set('dk-km',       Math.round(_kpiKm).toLocaleString('pt-BR') + ' km');
  set('dk-clientes', clientesFiltrados.length);
  // Drop médio = entregas/volume totais (já filtrados por cliente/cidade/período
  // acima) divididos pelo total de viagens (idem, respeita os mesmos filtros).
  const _kpiDropEntregas = _kpiViagens  > 0 ? _kpiEntregas / _kpiViagens  : 0;
  const _kpiDropVolume   = _kpiEntregas > 0 ? _kpiVol      / _kpiEntregas : 0;
  set('dk-drop-entregas', _kpiViagens  > 0 ? _kpiDropEntregas.toLocaleString('pt-BR', {minimumFractionDigits:1, maximumFractionDigits:1}) : '-');
  set('dk-drop-volume',   _kpiEntregas > 0 ? (_kpiDropVolume.toLocaleString('pt-BR', {minimumFractionDigits:1, maximumFractionDigits:1}) + ' m³') : '-');
  // % de Rotas Pedagiadas: viagens onde a detecção (linha reta terminal→
  // paradas, ver dashAgregar) achou pelo menos 1 pedágio, ÷ total de viagens
  // — mesmos filtros de cliente/cidade/período dos demais KPIs. É uma
  // ESTIMATIVA (linha reta, não o trajeto real via OSRM) — serve pra
  // enxergar tendência/volume de trabalho, não como conferência financeira.
  const _kpiPercPedagiadas = _kpiViagens > 0 ? (_kpiViagensComPedagio / _kpiViagens) * 100 : 0;
  set('dk-perc-pedagiadas', _kpiViagens > 0 ? Math.round(_kpiPercPedagiadas) + '%' : '-');
  // Tempo Gasto com Lançamento de Pedágios: 3 min de lançamento no Sem Parar
  // por viagem pedagiada (média fixa, conforme definido pela operação),
  // acumulado. Mostra em horas (com o total de minutos como referência).
  const _kpiTempoPedagioMin = _kpiViagensComPedagio * 3;
  set('dk-tempo-pedagio', _kpiTempoPedagioMin > 0 ? _dashFmtHoras(_kpiTempoPedagioMin) : '0h');
  const _elTempoPedagioMin = document.getElementById('dk-tempo-pedagio-min');
  if (_elTempoPedagioMin) _elTempoPedagioMin.textContent = _kpiViagensComPedagio > 0 ? `${_kpiTempoPedagioMin.toLocaleString('pt-BR')} min · ${_kpiViagensComPedagio} viagem(ns)` : '';
  // Consumo de Jornada — usa d.entradasTransportadora (já filtrado por
  // cidade/operação na fonte, dentro de dashAgregar, igual ao Ranking de
  // Transportadoras logo abaixo). Não é afetado pelo filtro de CLIENTE de
  // propósito: jornada é sobre o veículo/motorista, não sobre quem ele
  // atendeu naquele dia.
  const _dashJornada = dashAgregarJornada(d.entradasTransportadora);
  set('dk-jornada', _dashJornada.totalPct + '%');
  const _elJornadaHoras = document.getElementById('dk-jornada-horas');
  if (_elJornadaHoras) _elJornadaHoras.textContent = `${_dashFmtHoras(_dashJornada.totalUsadoMin)} / ${_dashFmtHoras(_dashJornada.totalDispMin)}`;
  dashRenderJornadaTransportadoras(_dashJornada);
  // Ociosidade — busca no Painel de Disponibilidade (Firestore), cruzando
  // com d.diasComViagemPorPlaca (já filtrado por cidade/operação na fonte,
  // igual à Jornada acima) pra saber quem foi disponibilizado e ficou parado.
  // É assíncrono (consulta ao Firestore), então atualiza o card e o gráfico
  // quando a resposta chegar — mostra "Carregando..." enquanto isso.
  set('dk-ociosidade', '…');
  const _elOciosidadeQtdIni = document.getElementById('dk-ociosidade-qtd');
  if (_elOciosidadeQtdIni) _elOciosidadeQtdIni.textContent = '';
  const _dashOciosidadeElBox = document.getElementById('dash-ociosidade-transp');
  if (_dashOciosidadeElBox) _dashOciosidadeElBox.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Carregando do Painel de Disponibilidade...</div>';
  const _meuTokenOciosidade = ++_dashOciosidadeToken;
  dashCarregarOciosidade(_dashSnapshotsAtivos, _dashCidadesSelecionadas, d.diasComViagemPorPlaca, d.placaCidade, d.placaCidadePorDia)
    .then(_dashOciosidade => {
      if (_meuTokenOciosidade !== _dashOciosidadeToken) return; // filtro mudou de novo antes de terminar — descarta
      set('dk-ociosidade', _dashOciosidade.pctOciosidade + '%');
      const _elOciosidadeQtd = document.getElementById('dk-ociosidade-qtd');
      if (_elOciosidadeQtd) _elOciosidadeQtd.textContent = `${_dashOciosidade.totalUsados} usados / ${_dashOciosidade.totalDisponibilizados} disponibilizados`;
      dashRenderOciosidadeTransportadora(_dashOciosidade);
    })
    .catch(e => {
      if (_meuTokenOciosidade !== _dashOciosidadeToken) return;
      console.warn('[Ociosidade] erro:', e);
      set('dk-ociosidade', '-');
      if (_dashOciosidadeElBox) _dashOciosidadeElBox.innerHTML = '<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Não foi possível carregar dados do Painel de Disponibilidade.</div>';
    });
  // Histórico por Veículo — só popula o seletor de placas na primeira vez.
  // NÃO busca nada automaticamente aqui: a consulta só roda quando o
  // usuário aperta "Buscar" (dashBuscarHistoricoVeiculo), de propósito —
  // trocar cidade/cliente/mês no resto do Dashboard não deve mais gerar
  // leitura no Firestore sozinho.
  const _selHistVeic = document.getElementById('dash-histveic-placa');
  if (_selHistVeic && _selHistVeic.options.length <= 1) dashPopularSeletorHistVeiculo();
  // Gráfico de barras: volume por cliente
  _dashUltimosClientesFiltrados = clientesFiltrados;
  const _itensVol = [...clientesFiltrados].sort((a, b) => _dashOrdemVol === 'asc' ? a.volume - b.volume : b.volume - a.volume);
  dashBarChart('dash-chart-vol', _itensVol, c=>c.volume.toFixed(1),
    '#f0be40', 'm³', c=>c.nome);
  // Gráfico de barras: entregas por cliente
  const _itensEnt = [...clientesFiltrados].sort((a, b) => _dashOrdemEnt === 'asc' ? a.entregas - b.entregas : b.entregas - a.entregas);
  dashBarChart('dash-chart-ent', _itensEnt,
    c=>c.entregas, '#70a8f0', 'ent.', c=>c.nome);
  // Gráfico Km vs Volume
  dashKmVolChart('dash-chart-km', clientesFiltrados);
  // Gráfico de ocupação por cliente
  dashOcupClienteChart('dash-chart-ocup', ocupFiltrados);
  // Gráfico de Ocupação vs Volume por Operação (cidade) — usa d.operacoes_ocup
  // direto (já veio filtrado por cidade dentro de dashAgregar); não é afetado
  // pelo filtro de CLIENTE de propósito, é uma visão por operação.
  dashOcupVolPorOperacaoChart('dash-chart-op-ocup-vol', d.operacoes_ocup);
  // Mapa — antes não respeitava NENHUM filtro de cliente (nem o do picker
  // de Clientes, que já existia). d.rotasMap é montado direto em
  // dashAgregar() só com o filtro de cidade aplicado; filtra aqui em cima,
  // por parada, mantendo a rota (terminal) mesmo que fique sem nenhuma
  // parada visível após o filtro.
  const _rotasMapFiltradas = _efetivos
    ? d.rotasMap.map(r => ({ ...r, paradas: (r.paradas || []).filter(p => _efetivos.has(p.nome)) }))
    : d.rotasMap;
  dashRenderMapa(_rotasMapFiltradas);
  // Ranking de Transportadoras — já vem filtrado por cidade (d.entradasTransportadora
  // é derivado de dashAgregar, que já aplicou o filtro de cidade na fonte).
  // NÃO é filtrado por cliente de propósito: o ranking é sobre quem prestou o
  // serviço, faz sentido continuar mostrando a transportadora inteira mesmo
  // filtrando por um cliente específico na tabela abaixo.
  dashRenderRankingTransportadoras(dashAgregarTransportadoras(d.entradasTransportadora));
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
  // "Clientes sem comprar" não é recalculado aqui (ele olha o histórico
  // completo, não os snapshots do período/mês selecionado) — só reaplica o
  // filtro de Cliente/Segmento em cima do que já foi calculado, pra ficar
  // consistente com o resto da tela sem reler o disco a cada troca de filtro.
  dashRenderClientesInativosUI(_dashUltimaListaInativos);
}
let _dashOrdemVol = 'desc'; // 'desc' = maior primeiro, 'asc' = menor primeiro
let _dashOrdemEnt = 'desc';
let _dashUltimosClientesFiltrados = []; // guarda o último dado renderizado, pra reordenar sem precisar re-agregar tudo
function _dashAtualizarTabsOrdenacao(classe, ordem) {
  document.querySelectorAll('.' + classe).forEach(b => {
    const ativo = b.dataset.ordem === ordem;
    b.classList.toggle('active-rank', ativo);
    b.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    b.style.color = ativo ? '#000' : 'var(--text-2)';
  });
}
function dashSetOrdemVol(ordem) {
  _dashOrdemVol = ordem;
  _dashAtualizarTabsOrdenacao('dash-ordvol-tab', ordem);
  const itens = [..._dashUltimosClientesFiltrados].sort((a, b) => ordem === 'asc' ? a.volume - b.volume : b.volume - a.volume);
  dashBarChart('dash-chart-vol', itens, c => c.volume.toFixed(1), '#f0be40', 'm³', c => c.nome);
}
function dashSetOrdemEnt(ordem) {
  _dashOrdemEnt = ordem;
  _dashAtualizarTabsOrdenacao('dash-ordent-tab', ordem);
  const itens = [..._dashUltimosClientesFiltrados].sort((a, b) => ordem === 'asc' ? a.entregas - b.entregas : b.entregas - a.entregas);
  dashBarChart('dash-chart-ent', itens, c => c.entregas, '#70a8f0', 'ent.', c => c.nome);
}
window.dashSetOrdemVol = dashSetOrdemVol;
window.dashSetOrdemEnt = dashSetOrdemEnt;
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
  const itens = clientes.filter(c => c.km > 0).sort((a, b) => b.km - a.km);
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
// ── Gráfico de Ocupação vs Volume por Operação (cidade) ────────────────────
// Mesmo estilo de barra dupla do dashKmVolChart, mas agrupado por cidade da
// operação (terminal) em vez de cliente. Sempre reflete o filtro de cidade
// ativo, porque os dados já vêm filtrados de dashAgregar.
let _dashOpOcupVolOrdem = 'volume'; // 'volume' | 'ocup'
let _dashUltimoOpOcupVol = [];
function dashOpOcupVolSetOrdem(campo) {
  _dashOpOcupVolOrdem = campo;
  document.querySelectorAll('.dash-opocup-tab').forEach(b => {
    const ativo = b.dataset.campo === campo;
    b.classList.toggle('active-rank', ativo);
    b.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    b.style.color = ativo ? '#000' : 'var(--text-2)';
  });
  dashOcupVolPorOperacaoChart('dash-chart-op-ocup-vol', _dashUltimoOpOcupVol);
}
window.dashOpOcupVolSetOrdem = dashOpOcupVolSetOrdem;
function dashOcupVolPorOperacaoChart(containerId, operacoes) {
  const el = document.getElementById(containerId);
  if (!el) return;
  _dashUltimoOpOcupVol = operacoes;
  if (!operacoes.length) {
    el.innerHTML = '<div style="color:#888;font-size:12px;padding:12px">Sem dados de operação para este período/filtro.</div>';
    return;
  }
  const itens = [...operacoes].sort((a, b) =>
    _dashOpOcupVolOrdem === 'ocup' ? b.ocup - a.ocup : b.volume - a.volume
  );
  const maxVol = Math.max(...itens.map(o => o.volume || 0), 1);
  el.innerHTML = itens.map((o, i) => {
    const pctVol = Math.round((o.volume / maxVol) * 100);
    const pctOcup = Math.min(Math.round(o.ocup), 100);
    const corOcup = '#4caf50'; // sempre verde, a pedido — antes variava por faixa (verde/amarelo/vermelho)
    const viagemLabel = o.viagens === 1 ? '1 viagem' : `${o.viagens} viagens`;
    return `<div style="padding:14px 16px;margin-bottom:10px;background:rgba(0,0,0,0.025);border:1px solid var(--border-dk);border-radius:10px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:9px;">
        <span style="font-size:13px;font-weight:700;color:var(--text,#111);">${i+1}. ${o.nome}</span>
        <span style="font-size:10.5px;font-weight:600;color:var(--text-3,#777);background:rgba(0,0,0,0.06);padding:2px 9px;border-radius:99px;">${viagemLabel}</span>
      </div>
      <div style="display:flex;gap:4px;align-items:center;">
        <span style="font-size:10px;font-weight:700;color:${corOcup};width:64px;text-align:right;">Ocupação</span>
        <div style="flex:1;height:8px;background:rgba(0,0,0,0.08);border-radius:99px;overflow:hidden;">
          <div style="width:${pctOcup}%;height:100%;background:${corOcup};border-radius:99px;"></div>
        </div>
        <span style="font-size:10px;font-weight:700;color:${corOcup};width:38px;">${pctOcup}%</span>
      </div>
      <div style="display:flex;gap:4px;align-items:center;margin-top:5px;">
        <span style="font-size:10px;font-weight:700;color:#000000;width:64px;text-align:right;">Volume</span>
        <div style="flex:1;height:8px;background:rgba(0,0,0,0.08);border-radius:99px;overflow:hidden;">
          <div style="width:${pctVol}%;height:100%;background:#f0be40;border-radius:99px;"></div>
        </div>
        <span style="font-size:10px;font-weight:700;color:#000000;width:60px;">${o.volume.toFixed(1)} m³</span>
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
    // renderer: L.canvas() faz TODAS as linhas (polylines) do mapa serem
    // desenhadas numa única camada <canvas> compartilhada, em vez de 1
    // elemento SVG por linha no DOM. Com muito histórico (uma linha por
    // trecho de cada viagem já roteirizada), isso evita centenas/milhares
    // de elementos DOM e é bem mais rápido pra desenhar e pra rolar a
    // página. Não muda a aparência nem os marcadores (que continuam
    // divIcon, sempre DOM — só as linhas usam canvas).
    _dashMap = L.map('dash-mapa', { zoomControl: true, attributionControl: false, renderer: L.canvas() });
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
    dashAtualizarClientesInativos();
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
// ── Resumo gerencial do dia — "Hoje" ────────────────────────────────────────
// Mesmo Dashboard de sempre (todos os KPIs já existentes: viagens, entregas,
// volume, ocupação, km, ranking, jornada, estouro, pedágio...), só que
// filtrado pras roteirizações SALVAS hoje — pra um gestor bater o olho no
// fim do dia e ver como a operação de hoje foi, sem precisar navegar por
// mês/período. Não cria nenhum KPI novo — reaproveita 100% do que já existe,
// só troca o conjunto de snapshots que entra no dashRender().
window.dashCarregarHoje = async function() {
  const store = await dashGetStoreMerged();
  const todos = Object.values(store).flat();
  // Data local (não UTC) — evita virar o dia errado perto da meia-noite.
  // Formato DD/MM/AAAA — é assim que datasEntrega guarda as datas (vem de
  // pedido.dataEntregaLogistica), diferente de savedAt (que é AAAA-MM-DD).
  const agora = new Date();
  const hojeBr = `${String(agora.getDate()).padStart(2,'0')}/${String(agora.getMonth()+1).padStart(2,'0')}/${agora.getFullYear()}`;
  // Filtra pela data de ENTREGA real das viagens (datasEntrega), não pela
  // data em que a roteirização foi salva (savedAt) — uma roteirização feita
  // hoje pode ser toda pra entregas de amanhã ou depois, e uma salva ontem
  // pode ter entregas de hoje. "Hoje" precisa refletir a operação de hoje,
  // não quando alguém mexeu no sistema.
  const doDia = todos.filter(s => Array.isArray(s.datasEntrega) && s.datasEntrega.includes(hojeBr));
  dashRender(doDia);
  const sel = document.getElementById('dash-mes-sel');
  if (sel) sel.value = ''; // "Hoje" não é nenhum dos meses do dropdown — desmarca visualmente
};
// ── Ferramentas de Km Real são só pra admin (recalcular/desfazer mexe direto
// nos arquivos do histórico compartilhado — não é algo pra qualquer usuário
// operacional/transportador acionar sem querer). Checa via window.S/
// window.USERS_DB, os mesmos que o resto do app usa pra saber o papel de
// quem está logado (ver nexta-frota-main.js).
function dashAtualizarVisibilidadeFerramentasKm() {
  const el = document.getElementById('dash-km-admin-tools');
  if (!el) return;
  const role = window.USERS_DB && window.S && window.USERS_DB[window.S.user]?.role;
  el.style.display = (role === 'admin') ? 'inline-flex' : 'none';
}
// ── Menu "Ferramentas" (admin) — Km Real / Diagnóstico / Diag. Pedágio ─────
// Antes eram 3 botões coloridos soltos direto na barra, mesmo sendo ações
// avançadas de uso raro. Agora ficam atrás de um menu, mesmo padrão visual
// dos painéis de filtro (Clientes/Operação) já existentes nesta tela.
function dashToggleFerramentas() {
  const menu = document.getElementById('dash-tools-menu');
  if (!menu) return;
  menu.style.display = menu.style.display !== 'none' ? 'none' : 'block';
}
document.addEventListener('click', function(e) {
  const menu = document.getElementById('dash-tools-menu');
  const btn  = document.getElementById('dash-tools-btn');
  if (menu && menu.style.display !== 'none' && !menu.contains(e.target) && !btn?.contains(e.target)) {
    menu.style.display = 'none';
  }
});
window.dashToggleFerramentas = dashToggleFerramentas;
// ═══════════════════════════════════════════════════════════════════════════
// CLIENTES SEM COMPRAR (alerta de inatividade)
// ═══════════════════════════════════════════════════════════════════════════
// Diferente do resto do Dashboard (que respeita o filtro de mês/período do
// topo), este painel sempre olha o HISTÓRICO COMPLETO — o objetivo é avisar
// "esse cliente sumiu", e um cliente que não compra há 40 dias já teria
// desaparecido da tela se o painel só olhasse o mês selecionado (ele
// simplesmente não apareceria em lugar nenhum, o que é o oposto de um alerta).
// Usa dashChaveCliente()/dashNomeCanônico() — as MESMAS funções que o resto
// do Dashboard já usa pra agrupar cliente por código SAP (com fallback pro
// nome normalizado), evitando contar "Auto Posto Moraes LTDA" e "Auto Posto
// Moraes" como dois clientes diferentes.
let _dashInativosLimiarDias = 15; // padrão: alerta a partir de 15 dias sem comprar
let _dashUltimaListaInativos = [];
function _dashParseDataBr(s) {
  const m = String(s || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
}
async function dashCalcularClientesInativos() {
  const store = await dashGetStoreMerged();
  const todos = Object.values(store).flat();
  // key (dashChaveCliente) -> { nome, ultimaData }
  const porCliente = {};
  todos.forEach(snap => {
    (snap.pedidos || []).forEach(p => {
      if (!p || !p.cliente) return;
      const d = _dashParseDataBr(p.dataEntregaLogistica);
      if (!d) return;
      const key = dashChaveCliente(p);
      if (!porCliente[key]) porCliente[key] = { nome: p.cliente, ultimaData: d, codigoSAP: p.codigoSAP ? String(p.codigoSAP).trim() : '' };
      else {
        porCliente[key].nome = dashNomeCanônico(porCliente[key].nome, p.cliente);
        if (!porCliente[key].codigoSAP && p.codigoSAP) porCliente[key].codigoSAP = String(p.codigoSAP).trim();
        if (d > porCliente[key].ultimaData) porCliente[key].ultimaData = d;
      }
    });
  });
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const resultado = Object.values(porCliente).map(c => ({
    nome: c.nome,
    codigoSAP: c.codigoSAP,
    ultimaData: c.ultimaData,
    dias: Math.round((hoje - c.ultimaData) / 86400000),
  })).sort((a, b) => b.dias - a.dias);
  _dashAtualizarMapaNomeSAP(resultado); // alimenta a busca de Segmento por SAP (ver dashClienteSegmento)
  return resultado;
}
function dashRenderClientesInativosUI(lista) {
  _dashUltimaListaInativos = lista;
  const box = document.getElementById('dash-clientes-inativos');
  if (!box) return;
  // Dois filtros se cruzam aqui: o GLOBAL (picker de Clientes + Segmento do
  // topo — "filtra tudo") e o LOCAL (Segmento dentro deste card, só afeta
  // este card). Um cliente só aparece se passar nos dois ao mesmo tempo.
  const _efetivosGlobais = dashClientesEfetivos(lista.map(c => c.nome));
  let listaFiltrada = _efetivosGlobais ? lista.filter(c => _efetivosGlobais.has(c.nome)) : lista;
  if (_dashSegmentosInativosSelecionados) {
    listaFiltrada = listaFiltrada.filter(c => {
      const seg = dashClienteSegmento(c.nome) || '—';
      return _dashSegmentosInativosSelecionados.has(seg);
    });
  }
  if (!listaFiltrada.length) {
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">${lista.length ? 'Nenhum cliente com o filtro atual (Clientes/Segmento).' : 'Nenhum pedido com data de entrega encontrado no histórico.'}</div>`;
    return;
  }
  const limiar = _dashInativosLimiarDias;
  const alertas = listaFiltrada.filter(c => c.dias >= limiar);
  if (!alertas.length) {
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">✅ Nenhum cliente passou de ${limiar} dias sem comprar (${listaFiltrada.length} cliente${listaFiltrada.length > 1 ? 's' : ''} no filtro atual).</div>`;
    return;
  }
  // Crítico = já dobrou o limite escolhido (ex.: limite 15 dias → crítico com 30+)
  box.innerHTML = alertas.map(c => {
    const critico = c.dias >= limiar * 2;
    const cor = critico
      ? { bg: 'rgba(220,38,38,.08)', border: '#DC2626', badge: '#DC2626' }
      : { bg: 'rgba(245,158,11,.08)', border: '#F59E0B', badge: '#F59E0B' };
    const dataFmt = c.ultimaData.toLocaleDateString('pt-BR');
    // Tag de diagnóstico — mostra o que o sistema ENCONTROU pra esse
    // cliente: o segmento (se achou) ou o motivo de não ter achado (sem
    // SAP no pedido + nome não bateu com nenhum cadastro). Fica visível
    // direto na tela, sem precisar abrir o console, pra confirmar se a
    // busca está funcionando ou não pra cada cliente específico.
    const segDetectado = dashClienteSegmento(c.nome);
    const sapUsado = _dashMapaNomeParaSAP[c.nome] || '';
    const tagSeg = segDetectado
      ? `<span style="font-size:9px;font-weight:700;color:#3730A3;background:#EEF2FF;border:1px solid #C7D2FE;border-radius:4px;padding:1px 6px;white-space:nowrap;" title="${sapUsado ? 'Encontrado via SAP ' + sapUsado : 'Encontrado por nome'}">${segDetectado}</span>`
      : `<span style="font-size:9px;font-weight:600;color:#9CA3AF;background:#F3F4F6;border:1px solid #E5E7EB;border-radius:4px;padding:1px 6px;white-space:nowrap;" title="${sapUsado ? 'Tem SAP ' + sapUsado + ' mas nenhum cadastro com esse SAP tem Segmento preenchido' : 'Sem Código SAP neste pedido — tentou por nome (\"' + c.nome + '\") e não achou cadastro com esse nome e Segmento preenchido'}">sem segmento — ⓘ</span>`;
    return `
      <div style="display:flex;align-items:center;gap:12px;padding:10px 12px;margin-bottom:5px;background:${cor.bg};border-left:3px solid ${cor.border};border-radius:0 6px 6px 0;">
        <div style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px;font-weight:600;color:var(--text);" title="${c.nome}">${critico ? '🔴' : '🟡'} ${c.nome}</div>
        ${tagSeg}
        <div style="font-size:11px;color:var(--text-3);white-space:nowrap;">última entrega: ${dataFmt}</div>
        <div style="font-size:11px;font-weight:800;color:#fff;background:${cor.badge};border-radius:99px;padding:3px 10px;white-space:nowrap;">${c.dias} dias sem comprar</div>
      </div>`;
  }).join('');
}
async function dashAtualizarClientesInativos() {
  const box = document.getElementById('dash-clientes-inativos');
  if (!box) return; // aba ainda não montada nessa carga da página
  if (!window.dirHandleHistorico) {
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Selecione a pasta do histórico primeiro.</div>`;
    return;
  }
  try {
    const lista = await dashCalcularClientesInativos();
    dashRenderClientesInativosUI(lista);
  } catch (e) {
    console.warn('[dashAtualizarClientesInativos] falha ao calcular:', e);
    box.innerHTML = `<div style="color:var(--text-3);text-align:center;padding:24px;font-size:12px;">Erro ao calcular — veja o console.</div>`;
  }
}
function dashSetLimiarInativos(dias) {
  _dashInativosLimiarDias = dias;
  document.querySelectorAll('.dash-inativo-limiar-tab').forEach(el => {
    const ativo = parseInt(el.dataset.dias, 10) === dias;
    el.classList.toggle('active-rank', ativo);
    el.style.background = ativo ? 'var(--pet-green,#b5e51d)' : 'transparent';
    el.style.color = ativo ? '#000' : 'var(--text-2)';
  });
  dashRenderClientesInativosUI(_dashUltimaListaInativos);
}
window.dashAtualizarClientesInativos = dashAtualizarClientesInativos;
window.dashSetLimiarInativos = dashSetLimiarInativos;
// ── Hook: popular meses quando abre a aba ─────────────────────────────────
const _origShowTab = window.showTab;
window.showTab = function(tab) {
  if (_origShowTab) _origShowTab(tab);
  if (tab === 'dashboard_rot') {
    dashPopularMeses();
    dashAtualizarVisibilidadeFerramentasKm();
    dashAutoRodarKmRealSeNecessario(); // idempotente (só roda 1x/dia) — reforça caso a checagem do carregamento inicial não tenha rodado ainda
    dashAtualizarClientesInativos(); // independe do filtro de mês/período — olha o histórico completo
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
  dashAtualizarVisibilidadeFerramentasKm();
  dashAutoRodarKmRealSeNecessario(); // dispara sozinho em segundo plano, no máx 1x/dia — ver comentário na função
  dashAtualizarClientesInativos(); // idempotente — silencioso se a pasta do histórico ainda não estiver selecionada

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
  // Respeita filtro de clientes E segmento ativos (mesmo conjunto efetivo
  // usado em toda a tela — ver dashClientesEfetivos())
  const filtroAtivo = dashClientesEfetivos(_dashTodosClientes);

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
      snapshots.forEach(snap => {
        const res = snap.resultado || {}, vecs = snap.veiculos || [];
        vecs.forEach(v => {
          const capV = v.capacidade || v.capacidadeTotal || 0;
          const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas||[]).length);
          viagens.forEach(vi => {
            const atende = vi.paradas.some(par => {
              const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
              return filtroAtivo.has(n);
            });
            if (!atende) return;
            if (capV > 0) _exFiltCap += capV;
            vi.paradas.forEach(par => {
              const n = (par.pedido||{}).cliente||(par.pedido||{}).nomeCliente||par.nome||'';
              if (filtroAtivo.has(n)) _exFiltVol += par.volumeTotal || 0;
            });
          });
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
          const termNomeViagem = vi.terminalOrigem || vi.paradas?.find(p => p.pedido?.terminal)?.pedido?.terminal || v.terminal;
          const term = terms.find(t => t.nome === termNomeViagem);

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
// A chave da OpenRouteService NÃO fica mais aqui — antes ficava exposta
// literalmente no código-fonte (qualquer um com DevTools conseguia copiar
// e gastar a cota diária de 2.000 requisições da empresa). Agora a
// requisição passa por /api/ors-proxy (Vercel Serverless Function — ver
// api/ors-proxy.js), que segura a chave só no servidor, via variável de
// ambiente ORS_API_KEY configurada no painel da Vercel.
const ORS_PROXY_URL = '/api/ors-proxy';
// ── Limitador de ritmo (rate limiter) pro proxy da ORS ──────────────────────
// A ORS free tier tem DOIS limites separados: 2.000 requisições/DIA (já
// controlado pelo teto em dashPreencherKmRealResultado) e um limite de
// RAJADA por minuto (bem mais baixo, ~40/min) — esse segundo limite nunca
// tinha sido tratado, e disparar vários lotes em sequência rápida (mesmo
// com só 3 de concorrência por vez) estourava ele quase imediatamente,
// gerando uma sequência de erros 429 (visível no painel de uso da ORS: só
// ~30 de ~730 chamadas tiveram sucesso numa rodada). A cada chamada, esse
// limitador espera até haver uma "vaga" dentro da janela de 60s.
const ORS_MAX_REQ_POR_MINUTO = 30; // margem de segurança sobre o limite real da ORS (~40/min)
const _orsHistoricoChamadas = [];
async function _orsAguardarVaga() {
  const agora = Date.now();
  // Descarta do histórico qualquer chamada com mais de 60s (já saiu da janela)
  while (_orsHistoricoChamadas.length && agora - _orsHistoricoChamadas[0] > 60000) {
    _orsHistoricoChamadas.shift();
  }
  if (_orsHistoricoChamadas.length >= ORS_MAX_REQ_POR_MINUTO) {
    const esperaMs = 60000 - (agora - _orsHistoricoChamadas[0]) + 200; // +200ms de folga
    await new Promise(r => setTimeout(r, Math.max(esperaMs, 0)));
    return _orsAguardarVaga(); // reavalia depois de esperar (pode ter mais de uma vaga liberada)
  }
  _orsHistoricoChamadas.push(Date.now());
}
async function osrmFetchSegmento(a, b) {
  // ── Cache em memória por par de coordenadas ─────────────────────────────
  // Trechos entre terminal↔cliente e cliente↔cliente se repetem MUITO (o
  // mesmo terminal atende dezenas de rotas, o mesmo cliente aparece em
  // vários dias) — sem cache, cada abertura de mapa refazia a chamada pra
  // ORS do zero, mesmo pra um trecho já calculado minutos antes na mesma
  // sessão. Isso somava tempo de espera (rate limiter de 30/min é
  // compartilhado com TODO o resto do sistema, incluindo o "Km Real" em
  // lote) sem necessidade nenhuma, já que o trajeto de um mesmo par de
  // pontos não muda. Arredonda a 4 casas decimais (~11m de precisão) pra
  // reaproveitar mesmo com pequena variação de ponto flutuante.
  const _round = (n) => Math.round(n * 10000) / 10000;
  const _cacheKey = `${_round(a.lat)},${_round(a.lon)}|${_round(b.lat)},${_round(b.lon)}`;
  if (!window._osrmSegmentoCache) window._osrmSegmentoCache = new Map();
  const _cached = window._osrmSegmentoCache.get(_cacheKey);
  if (_cached) return _cached;
  const _resultado = await _osrmFetchSegmentoSemCache(a, b);
  // Só guarda em cache resultado de rota real (não o fallback de linha
  // reta) — a linha reta é barata de recalcular e não vale a pena travar
  // no cache caso a ORS volte a funcionar logo em seguida.
  if (_resultado && _resultado.coords && _resultado.coords.length > 2) {
    window._osrmSegmentoCache.set(_cacheKey, _resultado);
  }
  return _resultado;
}
async function _osrmFetchSegmentoSemCache(a, b) {
  // ── OpenRouteService (via proxy), perfil driving-hgv (caminhão de verdade) ──
  // Substituiu o OSRM público (router.project-osrm.org), que NUNCA teve
  // perfil de caminhão de verdade — confirmado em issue oficial do projeto
  // OSRM: esse servidor ignora silenciosamente qualquer nome de perfil na
  // URL e sempre devolve rota de carro comum, disfarçada de "sucesso". Isso
  // fazia o sistema calcular caminho de carro (ex.: atravessar a Serra) onde
  // um caminhão de carga (principalmente combustível/inflamável) na prática
  // é proibido e vai pelo litoral — daí pedágio "sumido" mesmo com a praça
  // certa cadastrada, porque a rota calculada nunca passava perto dela.
  // hazmat:true pede pra ORS evitar vias com restrição a carga perigosa
  // quando essa informação existe no mapa (OpenStreetMap) da região — não é
  // garantia 100% (depende de quão bem aquele trecho foi mapeado), mas é
  // reconhecimento de verdade, ausente no OSRM público.
  for (let tentativa = 0; tentativa < 2; tentativa++) {
    try {
      await _orsAguardarVaga();
      const res = await fetch(ORS_PROXY_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          coordinates: [[a.lon, a.lat], [b.lon, b.lat]],
          options: {
            vehicle_type: 'hgv',
            profile_params: { restrictions: { hazmat: true } },
          },
        }),
      });
      if (res.ok) {
        const data = await res.json();
        const feat = data.features?.[0];
        if (feat?.geometry?.coordinates?.length) {
          return {
            coords: feat.geometry.coordinates.map(c => [c[1], c[0]]), // ORS devolve [lon,lat] — inverte pra [lat,lon]
            distKm: (feat.properties?.summary?.distance || 0) / 1000,
          };
        }
        break; // resposta ok mas sem rota utilizável — não adianta repetir, cai pro fallback
      }
      if (res.status === 429 && tentativa === 0) {
        // Rajada estourada mesmo com o limitador (ex.: outra aba/usuário
        // consumindo ao mesmo tempo) — espera um pouco mais e tenta 1 vez
        // antes de desistir e cair pro fallback de carro/linha reta.
        console.warn('[osrmFetchSegmento] 429 (rajada) mesmo com o limitador — aguardando 3s e tentando 1 vez mais.');
        await new Promise(r => setTimeout(r, 3000));
        continue;
      }
      console.warn('[osrmFetchSegmento] proxy/ORS respondeu erro, caindo pro OSRM público (linha reta/carro):', res.status, await res.text().catch(()=>''));
      break;
    } catch (e) {
      console.warn('[osrmFetchSegmento] Falha ao chamar o proxy/ORS, caindo pro OSRM público (linha reta/carro):', e);
      break;
    }
  }
  // ── Fallback 1: OSRM público (car — sem perfil de caminhão de verdade) ──
  // Só usado se a ORS falhar (rede fora, chave inválida/estourada, etc.) —
  // melhor ter uma rota aproximada de carro do que nenhuma.
  try {
    const url = `https://router.project-osrm.org/route/v1/driving/` +
      `${a.lon},${a.lat};${b.lon},${b.lat}` +
      `?overview=full&geometries=geojson&steps=false`;
    const res = await fetch(url);
    if (res.ok) {
      const data = await res.json();
      if (data.code === 'Ok' && data.routes?.[0]) {
        return {
          coords:   data.routes[0].geometry.coordinates.map(c => [c[1], c[0]]),
          distKm:   data.routes[0].distance / 1000,
        };
      }
    }
  } catch (e) { /* cai pro fallback final */ }
  // ── Fallback 2: linha reta com distância haversine ───────────────────────
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

// ═══════════════════════════════════════════════════════════════════════════
// KM REAL DA VIAGEM (trajeto de verdade, via OSRM/ORS) — não mais linha reta
// ═══════════════════════════════════════════════════════════════════════════
// Monta os pontos (terminal → paradas → retorno, se houver) e soma o km REAL
// de cada trecho via osrmFetchSegmento — mesma fonte usada no "Mapa da
// Viagem" e na detecção de pedágio com trajeto real. Resultado vai direto em
// vi._kmAjustado, o MESMO campo que já existia pra ajuste manual no mapa —
// então o Dashboard e o cálculo de frete já usam esse número automaticamente,
// sem precisar mudar mais nada em quem LÊ o km.
// NÃO usa latLonEfetivo (que depende de cadastro de cliente/cidade ao vivo) —
// lê direto p.pedido.lat/lon, pra funcionar igual tanto numa roteirização
// recém-otimizada (fluxo normal) quanto processando um arquivo antigo do
// histórico (fluxo de recálculo em lote), sem depender de estado vivo.
async function dashCalcularKmRealViagem(v, vi, terms) {
  if (!v || !vi || !vi.paradas || !vi.paradas.length) return null;
  // Monta os pontos (terminal → paradas → retorno) e valida coordenada via
  // NextaKm (assets/km-utils.js) — fonte ÚNICA dessa lógica agora, usada
  // também pelo relatório de Frete. Ver comentário no próprio módulo sobre
  // o bug de coordenada (0,0) que essa validação evita.
  const pontos = NextaKm.montarPontosViagem(v, vi, terms).map(p => ({ ...p, _tag: p.tag }));
  const terminalNome = vi.terminalOrigem || v.terminal || vi.paradas[0]?.pedido?.terminal || '';
  if (!(terms || []).find(t => t.nome === terminalNome)) {
    console.warn(`[dashCalcularKmRealViagem] terminal "${terminalNome}" não encontrado na lista de terminais deste arquivo (placa ${v.placa}) — viagem calculada SEM o ponto de origem/retorno, só entre paradas.`);
  }
  vi.paradas.forEach((p, i) => {
    const lat = parseFloat(p.pedido?.lat ?? p.lat);
    const lon = parseFloat(p.pedido?.lon ?? p.lon);
    if (!NextaKm.coordenadaValida(lat, lon)) {
      console.warn(`[dashCalcularKmRealViagem] parada ${i+1} (${p.pedido?.cliente || '?'}) sem coordenada válida (lat=${p.pedido?.lat ?? p.lat}, lon=${p.pedido?.lon ?? p.lon}) — excluída do cálculo, placa ${v.placa}.`);
    }
  });
  if (pontos.length < 2) return null;
  // Trava de sanidade: compara cada trecho REAL contra a distância em linha
  // reta (Haversine) do MESMO par de pontos. Estrada real pode ser mais longa
  // que a reta (curvas, contorno de serra/rio), mas não numa proporção
  // absurda — se algum trecho vier MUITO maior que a reta (coordenada errada,
  // geocodificação ruim, ou a rota calculada foi parar em outro estado/país),
  // rejeita a viagem inteira em vez de gravar um km calculado errado: melhor
  // manter a estimativa por linha reta (que pelo menos é consistente) do que
  // "corrigir" pra um número pior. +15 de folga absoluta cobre trechos curtos
  // onde a razão de multiplicação sozinha seria injusta (ex.: 1km reto virando
  // 3km real por causa de uma via sem saída direta).
  const FATOR_MAX = 3;
  const FOLGA_KM = 15;
  let totalKm = 0;
  let totalRetaKm = 0;
  for (let i = 0; i < pontos.length - 1; i++) {
    try {
      const seg = await osrmFetchSegmento(pontos[i], pontos[i + 1]);
      const segKm = seg?.distKm || 0;
      const retaKm = NextaKm.haversineKm(pontos[i].lat, pontos[i].lon, pontos[i+1].lat, pontos[i+1].lon);
      if (retaKm != null) totalRetaKm += retaKm;
      if (retaKm != null && segKm > retaKm * FATOR_MAX + FOLGA_KM) {
        console.warn(`[dashCalcularKmRealViagem] REJEITADO trecho implausível: ${pontos[i]._tag} → ${pontos[i+1]._tag} | real=${segKm.toFixed(1)}km vs reta=${retaKm.toFixed(1)}km | placa ${v.placa}, terminal "${terminalNome}". Mantendo estimativa por linha reta pra esta viagem inteira.`);
        return null;
      }
      totalKm += segKm;
    } catch (e) {
      return null; // falha de rede num trecho — não salva km parcial/errado
    }
  }
  // Log de auditoria: sempre grava o resultado (mesmo aceito), pra dar pra
  // conferir depois no console se algum valor grande é legítimo (viagem
  // interestadual de verdade) ou ainda suspeito mesmo tendo passado na trava.
  console.log(`[dashCalcularKmRealViagem] placa ${v.placa} | terminal "${terminalNome}" | ${pontos.length} pontos | real=${totalKm.toFixed(1)}km | reta=${totalRetaKm.toFixed(1)}km | razão=${totalRetaKm > 0 ? (totalKm/totalRetaKm).toFixed(2) : '?'}x`);
  return totalKm > 0 ? totalKm : null;
}
// Preenche vi._kmAjustado com o km real de TODAS as viagens de `resultado`
// que ainda não têm (pula as que já foram ajustadas manualmente no mapa, ou
// já processadas numa rodada anterior — torna a função segura de rodar de
// novo em cima do mesmo resultado sem refazer trabalho). Roda com
// concorrência limitada (não sobrecarrega o servidor de rota) e respeita um
// teto de chamadas (a ORS tem cota gratuita de 2.000 requisições/dia — ver
// comentário em api/ors-proxy.js) — se bater o teto, para e devolve
// quantas viagens ainda ficaram pendentes, pra tentar de novo depois (no dia
// seguinte a cota renova).
async function dashPreencherKmRealResultado(veiculosArr, resultado, terms, opts = {}) {
  const CONCORRENCIA = opts.concorrencia || 3;
  const tetoRequisicoes = opts.tetoRequisicoes ?? 1800; // margem de segurança sobre os 2000/dia da ORS
  const onProgresso = opts.onProgresso || (() => {});
  let requisicoesFeitas = 0;
  const pendentes = [];
  (veiculosArr || []).forEach(v => {
    const viagensDoVeic = resultado[v.id];
    if (!Array.isArray(viagensDoVeic)) return; // arquivo de formato antigo/diferente — pula esse veículo, não quebra o resto
    viagensDoVeic.forEach(vi => {
      if (!vi || !vi.paradas || !vi.paradas.length) return;
      if (NextaKm.obterKmViagem(vi).real) return; // já tem km real
      pendentes.push({ v, vi });
    });
  });
  let processadas = 0, falhas = 0;
  let pararPorCota = false;
  for (let start = 0; start < pendentes.length; start += CONCORRENCIA) {
    if (pararPorCota) break;
    const lote = pendentes.slice(start, start + CONCORRENCIA);
    await Promise.all(lote.map(async ({ v, vi }) => {
      if (pararPorCota) return;
      const nSegmentos = 1 + vi.paradas.length; // estimativa de chamadas que este item vai consumir
      if (requisicoesFeitas + nSegmentos > tetoRequisicoes) { pararPorCota = true; return; }
      requisicoesFeitas += nSegmentos;
      const km = await dashCalcularKmRealViagem(v, vi, terms);
      if (km != null) { vi._kmAjustado = km; vi._kmAjustadoAuto = true; processadas++; } else { falhas++; }
    }));
    onProgresso({ processadas, falhas, total: pendentes.length, requisicoesFeitas });
  }
  return { processadas, falhas, pendentesRestantes: pendentes.length - processadas - falhas, pararPorCota, requisicoesFeitas };
}
// ─── Recalcular km real de TODO o histórico salvo (retroativo) ─────────────
// Percorre os arquivos do histórico (mesma pasta usada pelo Dashboard/aba
// Histórico), pula os já substituídos por revisão (não contam pro Dashboard
// mesmo), preenche vi._kmAjustado das viagens que ainda não têm, e regrava
// o arquivo. Seguro de interromper e rodar de novo depois (retoma de onde
// parou, viagem que já tem _kmAjustado nunca é reprocessada) — necessário
// porque a cota gratuita da ORS (2000 req/dia) normalmente não cobre o
// histórico inteiro numa passada só se ele for grande.
// `silencioso`: true = modo automático (chamado sozinho na abertura do app,
// ver hook mais abaixo) — sem confirm()/alert(), sem forçar pedido de
// permissão (só usa se a pasta já estiver com permissão de escrita concedida
// de antes; navegador bloqueia pedido de permissão sem clique do usuário
// mesmo assim), e avisa o progresso por toast em vez de travar a tela.
// Botão manual "🛣️ Recalcular Km Real (Histórico)" continua disponível pra
// forçar uma rodada na hora (silencioso=false), inclusive pra CONCEDER a
// permissão de escrita da primeira vez (isso sim precisa de clique).
// Segunda camada de proteção (além de esconder o botão): mesmo chamando a
// função direto pelo console, sem ser admin ela recusa. Não é segurança de
// verdade (é tudo client-side, dá pra burlar com dev tools), mas evita
// acionar essas ferramentas por engano/curiosidade.
function _dashEhAdmin() {
  const role = window.USERS_DB && window.S && window.USERS_DB[window.S.user]?.role;
  return role === 'admin';
}
window.dashRecalcularKmHistorico = async function(silencioso = false) {
  if (!silencioso && !_dashEhAdmin()) { alert('Restrito ao administrador.'); return; }
  if (!window.dirHandleHistorico) {
    if (!silencioso) alert('Selecione a pasta do histórico primeiro (aba Histórico).');
    return;
  }
  if (!silencioso && !confirm(
    'Isso vai buscar o trajeto REAL (via rota, não linha reta) de cada viagem salva no histórico que ainda não tem, e regravar os arquivos.\n\n' +
    'Pode levar vários minutos e usa a cota diária de requisições de rota (limitada) — se o histórico for grande, pode não terminar tudo hoje; ' +
    'rodar de novo depois (nos dias seguintes) continua de onde parou, sem refazer o que já foi calculado.\n\n' +
    'O modo automático está temporariamente DESLIGADO (voltando a validar após o ajuste da trava de sanidade) — por enquanto só roda quando você clica aqui.\n\n' +
    'Deseja continuar?'
  )) return;
  let permOk = false;
  try { permOk = (await window.dirHandleHistorico.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  if (!permOk && !silencioso) {
    try { permOk = (await window.dirHandleHistorico.requestPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  }
  if (!permOk) {
    // No modo automático isso é esperado até o usuário clicar o botão manual
    // pelo menos 1x (navegador exige gesto do usuário pra conceder escrita) —
    // não interrompe nada, só não roda essa rodada.
    if (!silencioso) alert('Permissão de escrita negada.');
    else console.log('[dashRecalcularKmHistorico] rodada automática pulada — sem permissão de escrita concedida ainda (clique o botão manual 1x pra conceder).');
    return;
  }

  const arquivos = [];
  for await (const [name, handle] of window.dirHandleHistorico.entries()) {
    if (handle.kind === 'file' && name.endsWith('.json')) arquivos.push({ name, handle });
  }

  const btn = document.getElementById('btn-dash-recalcular-km');
  const setStatus = (txt) => { if (btn && !silencioso) btn.textContent = txt; };
  setStatus('⏳ Preparando...');

  let totalProcessadas = 0, totalFalhas = 0, arquivosAtualizados = 0, pararPorCota = false;
  let requisicoesUsadasRodada = 0; // teto é em REQUISIÇÕES (1 por trecho), não em viagens — cada viagem consome vários trechos
  const TETO_REQUISICOES_RODADA = 1800; // margem de segurança sobre os 2000/dia da ORS
  for (let i = 0; i < arquivos.length && !pararPorCota; i++) {
    const { name, handle } = arquivos[i];
    setStatus(`⏳ Km real: arquivo ${i + 1}/${arquivos.length}...`);
    let data;
    try {
      const file = await handle.getFile();
      data = JSON.parse(await file.text());
    } catch (e) { continue; }
    if (data.substituidoPor) continue; // não conta pro Dashboard — pula, economiza cota
    if (!data.resultado || !data.veiculos) continue;

    const r = await dashPreencherKmRealResultado(data.veiculos, data.resultado, data.terminais || [], {
      tetoRequisicoes: TETO_REQUISICOES_RODADA - requisicoesUsadasRodada, // teto GLOBAL da rodada, não por arquivo
      onProgresso: ({ processadas, total }) => setStatus(`⏳ Km real: arquivo ${i + 1}/${arquivos.length} (${processadas}/${total} viagens)...`),
    });
    totalProcessadas += r.processadas;
    totalFalhas += r.falhas;
    requisicoesUsadasRodada += r.requisicoesFeitas;
    if (r.processadas > 0) {
      try {
        const ws = await handle.createWritable();
        await ws.write(JSON.stringify(data, null, 2));
        await ws.close();
        arquivosAtualizados++;
      } catch (e) { console.warn(`[dashRecalcularKmHistorico] falha ao regravar ${name}:`, e); }
    }
    if (r.pararPorCota) pararPorCota = true;
  }

  setStatus('🛣️ Recalcular Km Real (Histórico)');
  if (totalProcessadas === 0 && totalFalhas === 0 && !pararPorCota) {
    // De verdade não tinha NADA pendente (nem sucesso, nem falha) — histórico
    // já 100% em dia. Único caso em que faz sentido dizer "nada a fazer".
    if (!silencioso) showToast('Histórico já está com o km real em dia — nada pra recalcular.', true);
    return;
  }
  if (totalProcessadas === 0 && totalFalhas > 0 && !pararPorCota) {
    // Tinha viagem pendente, mas TODAS falharam — geralmente porque o(s)
    // cliente(s) da viagem não têm coordenada cadastrada (ver aviso
    // "excluída do cálculo" no console) ou o cadastro do terminal está
    // incompleto. Isso é diferente de "nada a fazer": tem trabalho pendente
    // que só vai ser resolvido corrigindo o cadastro, não rodando de novo.
    const msgFalha = `Km real: ${totalFalhas} viagem(ns) pendente(s) tentada(s), mas nenhuma pôde ser calculada (coordenada de cliente/terminal ausente ou inválida — ver Console, F12, pra detalhes). Corrija o cadastro dessas paradas e rode de novo.`;
    if (silencioso) { if (typeof showToast === 'function') showToast(msgFalha, false); }
    else alert(msgFalha);
    return;
  }
  const msg = pararPorCota
    ? `Km real: cota diária atingida. ${totalProcessadas} viagem(ns) atualizada(s) em ${arquivosAtualizados} arquivo(s) até agora — continua sozinho amanhã.`
    : `Km real: ${totalProcessadas} viagem(ns) atualizada(s) em ${arquivosAtualizados} arquivo(s)${totalFalhas ? ` (${totalFalhas} falha(s) — coordenada ausente ou rede)` : ''}.`;
  if (silencioso) { if (typeof showToast === 'function') showToast(msg, !pararPorCota); }
  else alert(msg);
  if (typeof window.dashSincronizar === 'function') window.dashSincronizar();
};
// ── Roda automaticamente 1x por dia, sozinho, sem precisar clicar no botão ──
// Hook fica junto da inicialização do Dashboard (dashPopularMeses, mais
// abaixo) — dispara em segundo plano, sem travar nada, e só se a pasta do
// histórico já tiver permissão de escrita concedida de uma vez manual
// anterior (ver dashRecalcularKmHistorico, modo silencioso).
// DESLIGADO TEMPORARIAMENTE (ver KM_REAL_AUTO_HABILITADO): depois do episódio
// de km inflado, é mais seguro só rodar via clique manual (com o resultado
// visível na hora) até confirmar que a trava de sanidade está funcionando de
// verdade em produção — reative trocando pra `true` quando estiver validado.
const KM_REAL_AUTO_HABILITADO = false;
async function dashAutoRodarKmRealSeNecessario() {
  if (!KM_REAL_AUTO_HABILITADO) return;
  try {
    const hojeStr = new Date().toISOString().slice(0, 10);
    const chaveLS = 'dashKmRealAutoUltimoRun';
    if (localStorage.getItem(chaveLS) === hojeStr) return; // já rodou hoje
    localStorage.setItem(chaveLS, hojeStr); // marca ANTES de rodar — evita disparo duplo se a página recarregar no meio
    if (!window.dirHandleHistorico) return;
    await window.dashRecalcularKmHistorico(true);
  } catch (e) {
    console.warn('[dashAutoRodarKmRealSeNecessario] falhou:', e);
  }
}
// ─── Desfazer km real calculado automaticamente (rollback) ─────────────────
// Remove _kmAjustado de TODA viagem, com ou sem a marca _kmAjustadoAuto —
// inclui de propósito o que foi gravado ANTES dessa marca existir (a
// primeira rodada de recálculo, que gerou km inflado por causa de um bug já
// corrigido, foi salva sem essa marca, então filtrar só por ela deixava esse
// lote de fora e a ferramenta de desfazer não revertia nada).
// TRADE-OFF: se alguém tiver feito algum ajuste manual de verdade no Mapa da
// Viagem antes dessa funcionalidade existir, ele também seria desfeito aqui
// (não tem como distinguir os dois casos com certeza nos arquivos antigos).
// Isso é avisado no confirm() abaixo — se precisar, o ajuste manual é rápido
// de refazer reabrindo a viagem no mapa.
// ─── Diagnóstico: quais viagens estão puxando o Km Total pra cima ──────────
// Não faz NENHUMA requisição de rede nova — só lê o que já está gravado
// (_kmAjustado) e compara com a estimativa antiga (soma de par.distanciaKm,
// a mesma conta que o Dashboard usava antes de existir "km real"), pra achar
// as viagens com a MAIOR razão real/estimativa antiga. Mostra um resumo em
// alert() (fácil de printar) em vez de depender de copiar log do console.
// ── Diagnóstico: por que % Rotas Pedagiadas está em 0 (ou baixo) ────────────
// Roda em cima do que já está carregado na tela (_dashSnapshotsAtivos —
// respeita qualquer filtro ativo, incluindo "Hoje"), sem gastar nenhuma
// requisição nova. Categoriza cada viagem em: sem pontos suficientes pra
// detectar (coordenada de cliente/terminal ausente), com pontos mas sem
// pedágio no trajeto (dado correto, rota realmente não passa perto de
// nenhuma praça conhecida), ou com pedágio detectado — pra separar "é bug"
// de "é dado incompleto" de "é realmente 0 mesmo".
window.dashDiagnosticarPedagioHoje = async function() {
  if (!_dashEhAdmin()) { alert('Restrito ao administrador.'); return; }
  const snapshots = window._dashSnapshotsAtivos || [];
  if (!snapshots.length) { alert('Nenhum dado carregado no filtro atual — selecione um período primeiro.'); return; }
  // A busca de coordenada do cliente (latLonEfetivo → encontrarClienteDoPedido)
  // usa a lista "clientes" carregada NA SESSÃO ATUAL da Otimização Rotas —
  // não busca direto no Firestore. Se essa lista ainda não foi carregada
  // (ex.: usuário foi direto pra aba Dashboard, sem passar por Clientes),
  // clientes cadastrados de verdade apareciam como "sem coordenada" aqui só
  // por causa disso — não porque o cadastro estivesse incompleto. Garante
  // uma carga fresca antes de analisar, pra não dar diagnóstico falso.
  if ((!clientes || !clientes.length) && typeof window.dbGetCadastro === 'function') {
    try {
      const fsCli = await window.dbGetCadastro('clientes');
      if (fsCli && Object.keys(fsCli).length) {
        let id = 100;
        clientes = Object.values(fsCli).map(c => ({ ...c, id: c.id ?? id++ }));
      }
    } catch (e) { console.warn('[dashDiagnosticarPedagioHoje] falha ao carregar clientes do Firestore:', e); }
  }
  let semPontos = 0, comPontosSemPedagio = 0, comPedagio = 0, semTerminalCadastrado = 0;
  const exemplosSemPontos = [];
  const exemplosTerminal = [];
  snapshots.forEach(snap => {
    const res = snap.resultado || {};
    const vecs = snap.veiculos || [];
    const terms = snap.terminais || [];
    vecs.forEach(v => {
      const viagensDoVeic = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas || []).length);
      if (!viagensDoVeic.length) return; // veículo sem viagem no filtro atual — não entra na contagem
      viagensDoVeic.forEach(vi => {
        // Terminal DESTA viagem — vi.terminalOrigem (gravado na criação da
        // viagem) tem prioridade, depois o terminal do primeiro pedido com
        // essa info, só then o terminal do veículo (v.terminal) — que hoje
        // fica vazio de propósito pra veículos sem base fixa (podem carregar
        // em mais de uma base da mesma cidade).
        const termNomeViagem = vi.terminalOrigem || vi.paradas?.find(p => p.pedido?.terminal)?.pedido?.terminal || v.terminal;
        const term = terms.find(t => t.nome === termNomeViagem);
        const tLat = term?.lat, tLon = term?.lon;
        if (!term) {
          semTerminalCadastrado++;
          if (exemplosTerminal.length < 8) {
            exemplosTerminal.push(`${v.placa} — terminal desta viagem: "${termNomeViagem || '(vazio)'}" | terminais disponíveis neste arquivo: ${terms.map(t => `"${t.nome}"`).join(', ') || '(nenhum)'}`);
          }
        } else if (!NextaKm.coordenadaValida(tLat, tLon)) {
          semTerminalCadastrado++;
          if (exemplosTerminal.length < 8) {
            exemplosTerminal.push(`${v.placa} — terminal "${term.nome}" encontrado, mas lat/lon dele é "${tLat}"/"${tLon}" (inválido)`);
          }
        }
        const pontosPedagio = [];
        if (NextaKm.coordenadaValida(tLat, tLon)) pontosPedagio.push({ lat: parseFloat(tLat), lon: parseFloat(tLon) });
        vi.paradas.forEach(par => {
          const coords = latLonEfetivo ? latLonEfetivo(par.pedido) : { lat: par.lat, lon: par.lon };
          const lat = coords?.lat ?? par.lat, lon = coords?.lon ?? par.lon;
          if (NextaKm.coordenadaValida(lat, lon)) {
            pontosPedagio.push({ lat: parseFloat(lat), lon: parseFloat(lon) });
          } else if (exemplosSemPontos.length < 8) {
            // Detalha o motivo exato — pra distinguir "sem match no cadastro"
            // de "match achado mas coordenada do cliente também inválida" de
            // "pedido já vinha com coordenada própria zerada".
            const cliMatch = encontrarClienteDoPedido ? encontrarClienteDoPedido(par.pedido) : null;
            const motivo = !cliMatch
              ? `SEM MATCH no cadastro (nome/SAP não bateu — codigoSAP do pedido: "${par.pedido?.codigoSAP || '(vazio)'}", nome: "${par.pedido?.cliente || '?'}")`
              : `Match achado ("${cliMatch.nome}"), mas lat/lon do cadastro é "${cliMatch.lat}"/"${cliMatch.lon}"`;
            exemplosSemPontos.push(`${par.pedido?.cliente || '?'} — ${motivo}`);
          }
        });
        if (pontosPedagio.length < 2) { semPontos++; return; }
        const eixos = v.eixos || 2;
        const pd = (typeof detectarPedagiosNaRota === 'function') ? detectarPedagiosNaRota(pontosPedagio, eixos, 3) : [];
        if (pd.length > 0) comPedagio++; else comPontosSemPedagio++;
      });
    });
  });
  const total = semPontos + comPontosSemPedagio + comPedagio;
  console.log('[dashDiagnosticarPedagioHoje] Exemplos de terminal não encontrado/inválido:');
  console.table(exemplosTerminal.map(linha => ({ detalhe: linha })));
  console.log('[dashDiagnosticarPedagioHoje] Exemplos sem coordenada suficiente (detalhe completo):');
  console.table(exemplosSemPontos.map(linha => ({ detalhe: linha })));
  alert(
    `DIAGNÓSTICO — % ROTAS PEDAGIADAS (filtro atual)\n\n` +
    `Clientes carregados na sessão pra fazer o match: ${clientes ? clientes.length : 0}\n\n` +
    `Total de viagens analisadas: ${total}\n` +
    `✓ Com pedágio detectado: ${comPedagio}\n` +
    `– Com coordenadas OK, mas sem pedágio no trajeto (correto, não é bug): ${comPontosSemPedagio}\n` +
    `⚠ SEM coordenada suficiente pra sequer tentar detectar (cliente/terminal sem lat-lon): ${semPontos}\n` +
    `⚠ Veículos com terminal não encontrado no cadastro deste arquivo: ${semTerminalCadastrado}\n\n` +
    (exemplosTerminal.length ? `Exemplos de terminal com problema (ver lista completa no Console, F12):\n${exemplosTerminal.slice(0,2).join('\n')}${exemplosTerminal.length > 2 ? `\n...+${exemplosTerminal.length - 2} no Console` : ''}\n\n` : '') +
    (exemplosSemPontos.length ? `Exemplos sem coordenada (ver lista completa e detalhada no Console, F12):\n${exemplosSemPontos.slice(0,3).join('\n')}${exemplosSemPontos.length > 3 ? `\n...+${exemplosSemPontos.length - 3} no Console` : ''}` : '')
  );
};
window.dashDiagnosticarKmReal = async function() {
  if (!_dashEhAdmin()) { alert('Restrito ao administrador.'); return; }
  if (!window.dirHandleHistorico) { alert('Selecione a pasta do histórico primeiro (aba Histórico).'); return; }
  let permOk = false;
  try { permOk = (await window.dirHandleHistorico.queryPermission({ mode: 'read' })) === 'granted'; } catch(e) {}
  if (!permOk) { try { permOk = (await window.dirHandleHistorico.requestPermission({ mode: 'read' })) === 'granted'; } catch(e) {} }
  if (!permOk) { alert('Permissão de leitura negada.'); return; }

  const arquivos = [];
  for await (const [name, handle] of window.dirHandleHistorico.entries()) {
    if (handle.kind === 'file' && name.endsWith('.json')) arquivos.push({ name, handle });
  }

  const linhas = []; // { placa, terminal, arquivo, kmReal, kmEstAntiga, razao }
  let totalKmReal = 0, totalKmEstAntiga = 0, viagensComKmReal = 0;
  for (const { name, handle } of arquivos) {
    let data;
    try { data = JSON.parse(await (await handle.getFile()).text()); } catch(e) { continue; }
    if (data.substituidoPor) continue;
    if (!data.resultado || !data.veiculos) continue;
    data.veiculos.forEach(v => {
      const viagensDoVeic = data.resultado[v.id];
      if (!Array.isArray(viagensDoVeic)) return;
      viagensDoVeic.forEach(vi => {
        if (!vi || !vi.paradas || !vi.paradas.length) return;
        const kmEstAntiga = NextaKm.estimativaLinhaReta(vi);
        totalKmEstAntiga += kmEstAntiga;
        const kmInfo = NextaKm.obterKmViagem(vi);
        if (kmInfo.real) {
          totalKmReal += kmInfo.km;
          viagensComKmReal++;
          linhas.push({
            placa: v.placa,
            terminal: vi.terminalOrigem || v.terminal || '?',
            arquivo: name,
            kmReal: kmInfo.km,
            kmEstAntiga,
            razao: kmEstAntiga > 0 ? kmInfo.km / kmEstAntiga : Infinity,
          });
        }
      });
    });
  }
  linhas.sort((a, b) => b.kmReal - a.kmReal);
  console.log('[dashDiagnosticarKmReal] TOP 15 viagens por km real (maior pra menor):');
  console.table(linhas.slice(0, 15));
  const top10Texto = linhas.slice(0, 10)
    .map((l, i) => `${i+1}. ${l.placa} (${l.terminal}) — real=${l.kmReal.toFixed(0)}km, estimativa antiga=${l.kmEstAntiga.toFixed(0)}km, razão=${l.razao === Infinity ? '∞' : l.razao.toFixed(1)}x [${l.arquivo}]`)
    .join('\n');
  alert(
    `DIAGNÓSTICO KM REAL\n\n` +
    `Viagens com km real calculado: ${viagensComKmReal}\n` +
    `Soma km real: ${totalKmReal.toFixed(0)} km\n` +
    `Soma estimativa antiga (linha reta): ${totalKmEstAntiga.toFixed(0)} km\n` +
    `Razão geral: ${totalKmEstAntiga > 0 ? (totalKmReal/totalKmEstAntiga).toFixed(1) : '?'}x\n\n` +
    `TOP 10 maiores viagens (também no console, F12, com mais detalhe):\n${top10Texto}`
  );
};
window.dashDesfazerKmRealAuto = async function() {
  if (!_dashEhAdmin()) { alert('Restrito ao administrador.'); return; }
  if (!window.dirHandleHistorico) { alert('Selecione a pasta do histórico primeiro (aba Histórico).'); return; }
  if (!confirm(
    'Isso vai REMOVER o km real calculado (pelo botão/rotina "Recalcular Km Real") de TODAS as viagens do histórico, ' +
    'voltando o Km Total pra estimativa por linha reta de antes.\n\n' +
    'ATENÇÃO: como o primeiro recálculo (com o bug já corrigido) foi salvo sem uma marca distintiva, esta limpeza remove _kmAjustado de ' +
    'TODA viagem, inclusive eventuais ajustes manuais feitos no Mapa da Viagem antes de hoje (raro, mas se existir algum, precisa refazer ' +
    'manualmente depois).\n\n' +
    'Confirma?'
  )) return;
  console.log('[dashDesfazerKmRealAuto] iniciando...');
  let permOk = false;
  try { permOk = (await window.dirHandleHistorico.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) { console.warn('[dashDesfazerKmRealAuto] queryPermission falhou:', e); }
  if (!permOk) { try { permOk = (await window.dirHandleHistorico.requestPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) { console.warn('[dashDesfazerKmRealAuto] requestPermission falhou:', e); } }
  if (!permOk) { alert('Permissão de escrita negada.'); return; }
  console.log('[dashDesfazerKmRealAuto] permissão OK, listando arquivos...');

  // Todo o processamento fica dentro de um try/catch amplo: um erro não
  // capturado em QUALQUER ponto do loop (ex.: falha ao listar diretório,
  // arquivo bloqueado por outro processo) interrompia a função inteira sem
  // nunca chegar no alert() final — parecia "não fez nada" mesmo tendo
  // revertido parte dos arquivos. Agora qualquer falha aparece explicitamente
  // (console + alert), em vez de morrer em silêncio no meio do processamento.
  let arquivosAlterados = 0, viagensRevertidas = 0, arquivosLidos = 0, arquivosComKm = 0;
  try {
    const arquivos = [];
    for await (const [name, handle] of window.dirHandleHistorico.entries()) {
      if (handle.kind === 'file' && name.endsWith('.json')) arquivos.push({ name, handle });
    }
    console.log(`[dashDesfazerKmRealAuto] ${arquivos.length} arquivo(s) .json encontrado(s) na pasta.`);

    for (const { name, handle } of arquivos) {
      arquivosLidos++;
      let data;
      try {
        const file = await handle.getFile();
        data = JSON.parse(await file.text());
      } catch (e) { console.warn(`[dashDesfazerKmRealAuto] falha ao ler/parsear ${name}, pulando:`, e); continue; }
      if (!data.resultado || typeof data.resultado !== 'object') continue;
      let mudou = false;
      Object.values(data.resultado).forEach(viagensDoVeic => {
        if (!Array.isArray(viagensDoVeic)) return; // arquivo de formato antigo/diferente — pula esse veículo, não quebra o resto
        viagensDoVeic.forEach(vi => {
          if (vi && typeof vi._kmAjustado === 'number') {
            delete vi._kmAjustado;
            delete vi._kmAjustadoAuto;
            mudou = true;
            viagensRevertidas++;
          }
        });
      });
      if (mudou) {
        arquivosComKm++;
        try {
          const ws = await handle.createWritable();
          await ws.write(JSON.stringify(data, null, 2));
          await ws.close();
          arquivosAlterados++;
        } catch (e) { console.warn(`[dashDesfazerKmRealAuto] falha ao REGRAVAR ${name} (viagens deste arquivo NÃO foram revertidas de verdade):`, e); }
      }
    }
    console.log(`[dashDesfazerKmRealAuto] concluído: ${arquivosLidos} arquivo(s) lido(s), ${arquivosComKm} tinham km calculado, ${arquivosAlterados} regravado(s) com sucesso, ${viagensRevertidas} viagem(ns) revertida(s).`);
  } catch (e) {
    console.error('[dashDesfazerKmRealAuto] ERRO interrompeu o processamento no meio do caminho:', e);
    alert(`Erro durante o desfazer (parou no meio): ${e.message}\n\nAté aqui: ${viagensRevertidas} viagem(ns) revertida(s) em ${arquivosAlterados} arquivo(s). Abra o Console (F12) e me mande o erro completo — dá pra rodar de novo depois de corrigir, sem perder o que já foi revertido.`);
    if (typeof window.dashSincronizar === 'function') window.dashSincronizar();
    return;
  }
  // Também zera a marca de "já rodou hoje" pra permitir recalcular de novo
  // imediatamente (com a trava de sanidade agora em vigor), sem esperar o
  // próximo dia.
  // IMPORTANTE: marca hoje como "já rodou" (em vez de limpar a marca) — se
  // limpasse, o gatilho automático (dashAutoRodarKmRealSeNecessario) podia
  // disparar de novo sozinho no próximo reload/troca de aba e recolocar
  // _kmAjustado nas viagens antes mesmo de você conferir o resultado do
  // Desfazer — cada clique reduzia só uma fatia porque o automático tava
  // remendando por trás. Pra recalcular de novo de propósito, use o botão
  // "🛣️ Recalcular Km Real" manualmente (ele roda na hora, ignorando a marca).
  try { localStorage.setItem('dashKmRealAutoUltimoRun', new Date().toISOString().slice(0, 10)); } catch(e) {}
  alert(`Desfeito: ${viagensRevertidas} viagem(ns) revertida(s) em ${arquivosAlterados} arquivo(s). Km Total volta pra estimativa por linha reta.`);
  if (typeof window.dashSincronizar === 'function') window.dashSincronizar();
};



