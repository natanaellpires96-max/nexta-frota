// ═══════════════════════════════════════════════════════════════════════════
// NEXTA — ROTEIRIZADOR (motor de roteirização outbound)
// ═══════════════════════════════════════════════════════════════════════════
// Script comum (NÃO é type="module"): suas funções e variáveis de nível
// superior já ficam automaticamente em `window`, por isso não precisa de
// exposição manual como o script do app principal (acima, no <head>).
// Esse comportamento é o que permite que o HTML deste shell (#roteirizador-
// shell, gerado estaticamente no <body>) chame essas funções via onclick="".
//
// PONTO DE ENTRADA REAL: window.initRoteirizadorIntegrado() — chamado pelo
// app principal (renderRoteirizador() / setTab('roteirizador')) sempre que a
// aba "Roteirizador" é aberta. Não há login próprio aqui: o acesso já é
// controlado pelo login do NEXTA (apenas perfis admin/operacional).
//
// Depende de 3 funções do app principal, expostas via window.* por ele:
//   window.showToast, window.dateStr, window.sincronizarDisponibilidadeVeiculos
//
// ÍNDICE
//   Modal data de carregamento ............. abre/confirma data de carregamento antes de otimizar
//   Upload overlay pós-login ................ tela de upload (hoje sem uso ativo, ver Init)
//   Cadastral ................................ leitura de terminais/clientes/veículos (Excel/JSON)
//   Pedidos ................................... leitura/persistência de pedidos
//   Estado global ............................. pedidos, veiculos, terminaisCad, clientes, etc
//   Contador de IDs de viagem ................. gera ID único por viagem (petId)
//   DADOS PRÉ-CADASTRADOS ...................... import de dados.xlsx ou fallback JSON
//   Utilitários ................................ helpers gerais (datas, horas, formatação)
//   Filtro multi-terminal do Resumo Operação .... filtros do dashboard operacional
//   Centroides de cidades ........................ fallback de lat/lon por cidade
//   Best-insertion helpers ........................ heurística de inserção de parada
//   Combobox pesquisável .......................... componente de busca reutilizável
//   TERMINAIS ...................................... CRUD de terminais
//   CLIENTES ........................................ CRUD de clientes
//   PEDIDOS .......................................... CRUD de pedidos
//   Quebrar Pedido ................................... divide pedido em múltiplas entregas
//   Pedidos Liberados (Excel) ........................ importação de pedidos liberados
//   Upload de pedidos pelo usuário .................... upload manual de pedidos
//   CSV ................................................ exportação/leitura CSV
//   VEÍCULOS ............................................ CRUD de veículos/turnos
//   OTIMIZAÇÃO ........................................... motor de roteirização (função otimizar())
//   Exportação Herrlog .................................. exportação no formato Herrlog
//   Ajuste manual (drag-and-drop) ........................ mover viagens entre placas manualmente
//   Drag de parada (entrega individual) .................. mover parada individual
//   Troca de placa (veículo) .............................. substituir veículo de uma viagem
//   Mapa completo do veículo .............................. visualização de rota no mapa (Leaflet)
//   SHAREPOINT ............................................. integração de upload/leitura via SharePoint
//   HISTÓRICO DE ROTEIRIZAÇÕES ............................. File System Access API (salvar/ler local)
//   Init .................................................... window.initRoteirizadorIntegrado (ponto de entrada)
//   RELATÓRIO EXCEL — Roteirização Completa ................ exportarRelatorioRoteirizacao()
// ═══════════════════════════════════════════════════════════════════════════
// ─── Autenticação ─────────────────────────────────────────────────────────────
// O Roteirizador não possui login próprio: o acesso a esta aba já é controlado
// pelo login único do NEXTA (Firebase Auth) — apenas perfis "admin" e
// "operacional" conseguem abrir a aba "Roteirizador" (ver setTab() no app
// principal). Um sistema de login paralelo existia aqui anteriormente
// (hash local sem salt) mas foi removido por ser redundante e inseguro.
// ── Modal data de carregamento ────────────────────────────────────────────────
var _modoOtimizarPendente = null;
function abrirModalDataCarga(modo) {
  _modoOtimizarPendente = modo;
  // Usa data do seletor do header — preenche o campo e passa por confirmarOtimizar
  // para garantir a verificação do histórico antes de roteirizar
  const dataOp = document.getElementById('rot-data-operacao')?.value;
  if (dataOp) {
    document.getElementById('input-data-carga').value = dataOp;
    confirmarOtimizar();
    return;
  }
  // Fallback: abre modal com hoje
  const hoje = new Date();
  const iso = `${hoje.getFullYear()}-${String(hoje.getMonth()+1).padStart(2,'0')}-${String(hoje.getDate()).padStart(2,'0')}`;
  document.getElementById('input-data-carga').value = iso;
  const el = document.getElementById('modal-data-carga');
  el.style.display = 'flex';
}
function fecharModalDataCarga() {
  document.getElementById('modal-data-carga').style.display = 'none';
  _modoOtimizarPendente = null;
}
async function confirmarOtimizar() {
  const val = document.getElementById('input-data-carga').value;
  if (!val) { alert('Informe a data de carregamento.'); return; }

  // ── Pasta de histórico obrigatória — sem ela não há roteirização ──────────
  // Garante que nenhum ID de viagem seja duplicado entre roteirizações.
  if (!dirHandleHistorico) {
    alert('⚠️ Pasta de histórico não selecionada.\n\nSelecione a pasta na aba "Histórico" antes de roteirizar.\nO sistema precisa consultar o histórico para garantir IDs únicos.');
    fecharModalDataCarga();
    showTab('historico');
    return;
  }

  // Verifica permissão — pode expirar entre sessões do navegador
  let permOk = false;
  try { permOk = (await dirHandleHistorico.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  if (!permOk) {
    // Tenta reautorizar com interação do usuário
    try { permOk = (await dirHandleHistorico.requestPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  }
  if (!permOk) {
    alert('⚠️ Permissão da pasta de histórico negada.\n\nClique em "Selecionar pasta" na aba "Histórico" para reautorizar o acesso.\nSem isso não é possível garantir IDs únicos.');
    fecharModalDataCarga();
    showTab('historico');
    return;
  }

  const modoFinal = _modoOtimizarPendente || 'padrao';
  const dataFinal = new Date(val + 'T00:00:00');
  fecharModalDataCarga();
  otimizar(modoFinal, dataFinal);
}
// ─── Upload overlay pós-login ─────────────────────────────────────────────────
var _cadastralCarregado = false;
var _pedidosCarregados  = false;
function abrirUploadOverlay() {
  _cadastralCarregado = false;
  _pedidosCarregados  = false;
  const el = document.getElementById('upload-overlay');
  el.style.display = 'flex';
  document.getElementById('upload-status').style.display     = 'none';
  document.getElementById('upload-ped-status').style.display = 'none';
  document.getElementById('upload-file-input').value = '';
  document.getElementById('upload-ped-input').value  = '';
}
function fecharUploadOverlay() {
  document.getElementById('upload-overlay').style.display = 'none';
}
async function continuar() {
  fecharUploadOverlay();
  if (!_cadastralCarregado) await carregarDadosFixos();
  if (!_pedidosCarregados)  await carregarPedidosLiberados();
}
// ── Cadastral ─────────────────────────────────────────────────────────────────
function _setStatus(id, tipo, msg) {
  const el = document.getElementById(id);
  if (!el) return;
  el.style.display = '';
  const estilos = {
    carregando: ['#FFFBEB','#92400E','1px solid #FCD34D'],
    ok:         ['#ECFDF5','#065F46','1px solid #6EE7B7'],
    erro:       ['#FEE2E2','#B91C1C','1px solid #FECACA'],
  };
  const [bg, cor, brd] = estilos[tipo] || estilos.carregando;
  el.style.background = bg; el.style.color = cor; el.style.border = brd;
  el.textContent = msg;
}
async function uploadArquivoSelecionado(input) {
  const file = input.files[0];
  if (!file) return;
  _setStatus('upload-status', 'carregando', 'Carregando...');
  const ok = await carregarDoExcelArquivo(file);
  if (ok) {
    _cadastralCarregado = true;
    _setStatus('upload-status', 'ok', `✓ ${file.name} carregado.`);
  } else {
    _setStatus('upload-status', 'erro', '✕ Não foi possível ler o arquivo. Verifique se contém as planilhas Terminais, Clientes e Placas.');
    document.getElementById('upload-file-input').value = '';
  }
}
async function carregarDoExcelArquivo(file) {
  if (typeof XLSX === 'undefined') return false;
  try {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellDates: true });
    const sheet = nome => wb.Sheets[nome]
      ? XLSX.utils.sheet_to_json(wb.Sheets[nome], { defval: '' })
      : null;
    const rowsT = sheet('Terminais');
    const rowsC = sheet('Clientes');
    const rowsP = sheet('Placas');
    let carregouAlgo = false;
    if (rowsT?.length) {
      terminaisCad = rowsT.map(xlsxMapTerminal);
      renderTerminais(); atualizarDropdownsTerminais();
      carregouAlgo = true;
    }
    if (rowsC?.length) {
      clientes = rowsC.map(xlsxMapCliente);
      renderClientes(); atualizarDropdownsClientes();
      carregouAlgo = true;
    }
    if (rowsP?.length) {
      veiculos = rowsP.map(xlsxMapPlaca);
      renderVeiculos();
      carregouAlgo = true;
    }
    if (carregouAlgo) spSetStatus('ok', `✓ ${file.name} carregado`);
    return carregouAlgo;
  } catch (e) {
    console.warn('[upload excel]', e.message);
    return false;
  }
}
function uploadDragOver(e) {
  e.preventDefault();
  const el = document.getElementById('upload-drop-area');
  el.style.borderColor = 'var(--pet-green)'; el.style.background = 'var(--pet-green-bg)'; el.style.color = 'var(--pet-green)';
}
function uploadDragLeave() {
  const el = document.getElementById('upload-drop-area');
  el.style.borderColor = ''; el.style.background = ''; el.style.color = '';
}
async function uploadDrop(e) {
  e.preventDefault(); uploadDragLeave();
  const file = e.dataTransfer?.files?.[0];
  if (file) await uploadArquivoSelecionado({ files: [file] });
}
// ── Pedidos ───────────────────────────────────────────────────────────────────
async function uploadPedidosOverlay(input) {
  const file = input.files[0];
  if (!file) return;
  if (typeof XLSX === 'undefined') { _setStatus('upload-ped-status', 'erro', '✕ SheetJS não carregado.'); return; }
  _setStatus('upload-ped-status', 'carregando', 'Carregando...');
  try {
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array', cellDates: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    if (!rows.length) { _setStatus('upload-ped-status', 'erro', '✕ Planilha vazia ou sem dados reconhecidos.'); return; }
    const novos = xlsxMapPedidosLiberadosRows(rows);
    if (!novos.length) { _setStatus('upload-ped-status', 'erro', '✕ Nenhum pedido reconhecido. Verifique o modelo.'); return; }
    pedidos = novos;
    renderPedidos();
    _pedidosCarregados = true;
    _setStatus('upload-ped-status', 'ok', `✓ ${novos.length} pedido(s) carregado(s).`);
  } catch (err) {
    _setStatus('upload-ped-status', 'erro', `✕ Erro ao processar: ${err.message}`);
  } finally {
    input.value = '';
  }
}
function uploadPedDragOver(e) {
  e.preventDefault();
  const el = document.getElementById('upload-ped-drop-area');
  el.style.borderColor = 'var(--pet-green)'; el.style.background = 'var(--pet-green-bg)'; el.style.color = 'var(--pet-green)';
}
function uploadPedDragLeave() {
  const el = document.getElementById('upload-ped-drop-area');
  el.style.borderColor = ''; el.style.background = ''; el.style.color = '';
}
async function uploadPedDrop(e) {
  e.preventDefault(); uploadPedDragLeave();
  const file = e.dataTransfer?.files?.[0];
  if (file) await uploadPedidosOverlay({ files: [file] });
}
// ─── Estado global ────────────────────────────────────────────────────────────
// IMPORTANTE: usar var (não let/const) em todo este bloco. Este <script> e o
// <script> do Dashboard (mais abaixo no documento) são dois escopos de script
// separados — "let"/"const" no nível de topo NÃO viram propriedades de window
// e portanto não atravessam para o outro <script> (ReferenceError silencioso,
// mascarado por try/catch). "var" no nível de topo vira window.X, visível dos
// dois lados. Isto já causou bugs reais: mapaViagem, pedidos e ultimoResultado
// eram usados pelo módulo do mapa de viagem / Dashboard sem estarem acessíveis.
var pedidos = [], veiculos = [], terminaisCad = [], clientes = [];
var numComps = 0, numProdForm = 0;
var editandoTerminalId = null, editandoPedidoId = null, editandoClienteId = null, editandoVeiculoId = null;
var ultimoResultado = null, ultimoControleTempo = null;
var _sugestoesSplitDedicado = []; // sugestões de quebra manual pós-otimização modo dedicado
var _motoristasOverride = {}; // { petId: nome } — override manual por viagem, ignora regra diurno/noturno
var lockedTerminals = new Set(); // terminais com ajustes travados (não re-otimizados)
// ── Contador de IDs de viagem (P + MM + YY + seq 3 dígitos) ─────────────────
// Formato: P{MM}{YY}{NNN} ex: P0626001
// Chave interna: MM+YY ex: "0626" — sequência global por mês/ano, sem repetição
function _getPetSeq(chave) {
  try { return JSON.parse(localStorage.getItem('petSeq') || '{}')[chave] || 0; } catch { return 0; }
}
// Migra chaves legadas (MM+YY+DD = 6 dígitos) para o novo formato (MM+YY = 4 dígitos)
(function _migrarPetSeqLegado() {
  try {
    const stored = JSON.parse(localStorage.getItem('petSeq') || '{}');
    let changed = false;
    const merged = {};
    for (const [chave, seq] of Object.entries(stored)) {
      if (chave.length === 6) {
        // legado: MM+YY+DD → nova chave: MM+YY (primeiros 4 dígitos)
        const novaChave = chave.slice(0, 4);
        merged[novaChave] = Math.max(merged[novaChave] || 0, seq);
        changed = true;
      } else {
        merged[chave] = Math.max(merged[chave] || 0, seq);
      }
    }
    if (changed) localStorage.setItem('petSeq', JSON.stringify(merged));
  } catch(e) {}
})();
// Atribui petId a qualquer viagem real que ainda não tenha um
function atribuirPetIds(resultado, dataRef) {
  if (!resultado) return;
  if (!dataRef || isNaN(dataRef?.getTime?.() ?? NaN)) dataRef = new Date();
  // Tenta pegar data do primeiro pedido do resultado
  for (const viagens of Object.values(resultado)) {
    if (!Array.isArray(viagens)) continue;
    for (const vi of viagens) {
      if (vi._vazio || !(vi.paradas||[]).length) continue;
      const dl = vi.paradas[0]?.pedido?.dataEntregaLogistica;
      if (dl) {
        const pts = dl.split('/');
        if (pts.length >= 3) {
          const d = new Date(parseInt(pts[2]), parseInt(pts[1])-1, parseInt(pts[0]));
          if (!isNaN(d.getTime())) { dataRef = d; break; }
        }
      }
    }
  }
  const mm    = String(dataRef.getMonth()+1).padStart(2,'0');
  const yy    = String(dataRef.getFullYear()).slice(-2);
  const chave = mm + yy;   // ex: "0626" — chave por mês/ano (sem dia)
  let seq = _getPetSeq(chave);
  // Garante que nunca reutilize um ID já existente no resultado atual
  // (cobre viagens travadas de rodadas anteriores que mantêm seus IDs)
  Object.values(resultado).filter(Array.isArray).forEach(viagens => viagens.forEach(vi => {
    const m = (vi.petId||'').match(/^P(\d{4})(\d{3,})$/);
    if (m && m[1] === chave) seq = Math.max(seq, parseInt(m[2]));
  }));
  for (const viagens of Object.values(resultado)) {
    if (!Array.isArray(viagens)) continue;
    for (const vi of viagens) {
      if (vi._vazio || !(vi.paradas||[]).length) continue;
      if (!vi.petId) { seq++; vi.petId = `P${mm}${yy}${String(seq).padStart(3,'0')}`; }
    }
  }
  // NÃO persiste o contador aqui — só persiste ao salvar no histórico,
  // evitando que rodadas sucessivas de otimização avancem a sequência sem necessidade.
}
function _atualizarPetSeq(resultado) {
  if (!resultado) return;
  let stored = {};
  try { stored = JSON.parse(localStorage.getItem('petSeq') || '{}'); } catch {}
  Object.values(resultado).forEach(viagens => {
    if (!Array.isArray(viagens)) return;
    viagens.forEach(vi => {
      // Suporta formato novo P{MM}{YY}{NNN} (4 dígitos de chave)
      // e legado P{MM}{YY}{DD}{NNN} (6 dígitos) — migra chave legada para MM+YY
      let m = (vi.petId || '').match(/^P(\d{4})(\d{3,})$/);
      if (!m) {
        // tentativa de leitura do formato legado: extrai MM+YY dos primeiros 4 dígitos
        const ml = (vi.petId || '').match(/^P(\d{2})(\d{2})\d{2}(\d{3,})$/);
        if (ml) m = [null, ml[1] + ml[2], ml[3]];
      }
      if (!m) return;
      const [, chave, seq] = m;
      if (!stored[chave] || parseInt(seq) > stored[chave]) stored[chave] = parseInt(seq);
    });
  });
  try { localStorage.setItem('petSeq', JSON.stringify(stored)); } catch {}
}
var resultadoOriginal    = null;   // cópia do resultado puro da otimização
var historicoManual      = [];     // pilha de estados para desfazer
var dragInfo             = null;   // { tipo, veiculoId, tripIndex, ... } durante drag
var dirHandleHistorico = null; // precisa ficar em window para ser
// Cache de metadados do histórico: { [filename]: { savedAt, resumo, datasEntrega, salvoPor, versao } }
// Evita reler arquivos completos a cada "Atualizar lista" — só relê arquivos novos/modificados.
var _histMetaCache = {};
                                // acessível pelo módulo do Dashboard, em outro <script>.
var estadoAtual = null;
var ultimoItensOtimizacao = [];
var mapaViagem = null, camadaViagem = null, dadosMapaAtual = null;
var mapaGeral = null, camadaMapaGeral = null;
var filtroMapaPlaca = '';
var filtroMapaTerminais = new Set(); // terminais selecionados no filtro do resumo (vazio = todos)
const DIAS_NOMES     = ['DOM','SEG','TER','QUA','QUI','SEX','SÁB'];
const PRODUTOS = [
  '2000016 - GASOLINA TIPO C S 50',
  '2000017 - GASOLINA TIPO C S 50 ADIT',
  '2000019 - OLEO DIESEL B S10 ADIT',
  '2000020 - OLEO DIESEL B S500 COMUM',
  '2000021 - OLEO DIESEL B S500 ADIT',
  '2000023 - ETANOL HIDRATADO',
  '2000031 - PETRONAS GASOLINA COMUM',
  '2000032 - PETRONAS PRIMAX GASOLINA',
  '2000033 - PETRONAS ETANOL COMUM',
  '2000034 - PETRONAS PRIMAX ETANOL',
  '2000035 - PETRONAS DIESEL S10',
  '2000036 - PETRONAS DIESEL S500',
  '2000043 - PETRONAS PRIMAX GASOLINA ALTA OCTANAGEM',
  '2000045 - OLEO DIESEL B S10 COMUM',
  '2000047 - PETRONAS DYNAMIC DIESEL S10',
  '2000048 - PETRONAS DYNAMIC DIESEL S500',
];
const TIPOS_CAMINHAO = ['Truck','Bitruck','Cavalo Mecânico'];
// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  DADOS PRÉ-CADASTRADOS                                                       ║
// ║  Prioridade 1: dados.xlsx  (planilhas: Terminais | Clientes | Placas)        ║
// ║  Prioridade 2: terminais.json · clientes.json · placas.json (fallback)       ║
// ║  Coloque os arquivos na mesma pasta do HTML.                                 ║
// ╚══════════════════════════════════════════════════════════════════════════════╝
// ── Helpers de leitura de célula Excel ───────────────────────────────────────
function xlsxStr(row, ...keys) {
  for (const k of keys) {
    const v = row[k];
    if (v !== undefined && v !== null && v !== '') return String(v).trim();
  }
  return '';
}
function xlsxNum(row, fallback, ...keys) {
  for (const k of keys) { const v = parseFloat(row[k]); if (!isNaN(v)) return v; }
  return fallback;
}
function xlsxHora(val) {
  if (!val && val !== 0) return '';
  if (typeof val === 'string') return val.trim().slice(0, 5);
  if (typeof val === 'number') {
    const m = Math.round(val * 1440);
    return `${String(Math.floor(m/60)%24).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`;
  }
  if (val instanceof Date) {
    return `${String(val.getHours()).padStart(2,'0')}:${String(val.getMinutes()).padStart(2,'0')}`;
  }
  return String(val).trim().slice(0, 5);
}
// ── Mapeamentos planilha → modelo ────────────────────────────────────────────
function xlsxMapTerminal(r, i) {
  const diasStr = xlsxStr(r, 'Dias Ativos', 'DiasAtivos');
  const diasAtivos = diasStr
    ? diasStr.split(',').map(d => parseInt(d.trim())).filter(d => !isNaN(d))
    : [1, 2, 3, 4, 5];
  return {
    id: i + 1,
    nome:                    xlsxStr(r, 'Nome'),
    cidade:                  xlsxStr(r, 'Cidade'),
    distribuidora:           xlsxStr(r, 'Distribuidora', 'CIA'),
    empresaLocalExpedicao:   xlsxStr(r, 'Empresa Local Expedição', 'Empresa Local Expedicao', 'EmpresaLocalExpedicao'),
    lat:                     xlsxNum(r, 0,  'Latitude',  'Lat'),
    lon:                     xlsxNum(r, 0,  'Longitude', 'Lon'),
    fuso:                    xlsxStr(r, 'Fuso') || '-3',
    tempoCarregamentoMedioMin: xlsxNum(r, 60, 'Tempo Carregamento (min)', 'Tempo Carregamento'),
    aberturaPadrao:          xlsxHora(r['Abertura Padrão'] ?? r['Abertura']) || '06:00',
    fechamentoPadrao:        xlsxHora(r['Fechamento Padrão'] ?? r['Fechamento']) || '18:00',
    diasAtivos,
    horarios: {},
  };
}
function xlsxMapCliente(r, i) {
  const tiposStr = xlsxStr(r, 'Tipos Caminhão', 'TiposCaminhao');
  return {
    id: 100 + i,
    codigoSAP:         xlsxStr(r, 'Código SAP', 'CodigoSAP', 'SAP'),
    nome:              xlsxStr(r, 'Nome'),
    cidade:            xlsxStr(r, 'Cidade'),
    lat:               xlsxNum(r, 0,  'Latitude',  'Lat'),
    lon:               xlsxNum(r, 0,  'Longitude', 'Lon'),
    tempoDescargaMediaMin: xlsxNum(r, 45, 'Tempo Descarga (min)', 'Tempo Descarga'),
    restricaoHorario:  xlsxStr(r, 'Restrição Horário', 'Restricao', 'Restrição'),
    tiposCaminhao:     tiposStr ? tiposStr.split(';').map(t => t.trim()).filter(Boolean) : [],
    observacoes:       xlsxStr(r, 'Observações', 'Observacoes', 'Obs'),
    identidadePetronas: ['sim','s','true','1','yes'].includes((xlsxStr(r, 'Identidade Petronas', 'IdentidadePetronas') || '').toLowerCase()),
  };
}
function xlsxMapPlaca(r, i) {
  const comps = [];
  for (let c = 1; c <= 8; c++) {
    const cap = xlsxNum(r, 0, `C${c}`, `C${c} (m³)`, `Comp${c}`);
    if (cap > 0) comps.push({ cap, produto: xlsxStr(r, `C${c} Produto`, `C${c}Produto`) });
  }
  if (!comps.length) comps.push({ cap: 0, produto: '' });
  const jornadaInicio = xlsxHora(r['Jornada Início'] ?? r['JornadaInicio'] ?? r['Início']) || '06:00';
  const jornadaFim    = xlsxHora(r['Jornada Fim']    ?? r['JornadaFim']    ?? r['Fim'])    || '18:00';
  const terminalRaw   = xlsxStr(r, 'Terminal');
  let cidadeBase      = xlsxStr(r, 'Cidade Base', 'CidadeBase', 'Cidade');
  let terminal        = terminalRaw;
  if (terminalRaw && !terminalPorNome(terminalRaw)) {
    // Quando a planilha vier no formato "base por cidade" usando a coluna Terminal,
    // tratamos o valor como cidade base para não bloquear a roteirização.
    if (!cidadeBase) cidadeBase = terminalRaw;
    terminal = '';
  }
  return {
    id: 200 + i,
    placa:            xlsxStr(r, 'Placa'),
    implemento:       xlsxStr(r, 'Implemento'),
    transportadora:   xlsxStr(r, 'Transportadora'),
    tipo:             xlsxStr(r, 'Tipo') || 'Truck',
    terminal,
    cidadeBase,
    turno:            xlsxStr(r, 'Turno') || 'personalizado',
    jornadaInicio,
    jornadaFim,
    jornadaMin:       duracaoJornadaMin(jornadaInicio, jornadaFim),
    tempoPerdidoMin:  (() => {
      // Aceita valor numérico (minutos) na coluna nova ou HH:MM nas colunas legadas
      const numVal = xlsxNum(r, NaN, 'Tempo Refeicao (min)');
      if (!isNaN(numVal)) return numVal;
      return parseHoraMin(xlsxHora(r['Tempo Perdido'] ?? r['TempoPerdido'])) || 0;
    })(),
    velMediaCarregado: xlsxNum(r, 45, 'Vel Carregado (km/h)', 'VelCarregado'),
    velMediaVazio:     xlsxNum(r, 55, 'Vel Vazio (km/h)',     'VelVazio'),
    compartimentos:   comps,
    capacidade:       comps.reduce((s, c) => s + c.cap, 0),
    identidadePetronas: ['sim','s','true','1','yes'].includes((xlsxStr(r, 'Identidade Petronas', 'IdentidadePetronas') || '').toLowerCase()),
    disponibilidade:  'Disponível', // status sempre vem do Painel de Disponibilidade, não da planilha
    contrato:         xlsxStr(r, 'Contrato') || 'Dedicado',
    motoristaDiurno:  xlsxStr(r, 'Motorista Diurno', 'MotoristaDiurno', 'Motorista Diurno ', 'motorista diurno', 'Motorista'),
    motoristaNt:      xlsxStr(r, 'Motorista Noturno', 'MotoristaNoturno', 'Motorista Noturno ', 'motorista noturno', 'Motorista Nortuno', 'MotoristaNoturno '),
  };
}
// ── Leitura do Excel ─────────────────────────────────────────────────────────
async function carregarDoExcel(arquivo = 'dados.xlsx') {
  if (typeof XLSX === 'undefined') return false;
  try {
    const resp = await fetch(arquivo + '?t=' + Date.now());
    if (!resp.ok) return false;
    const wb = XLSX.read(await resp.arrayBuffer(), { type: 'array', cellDates: true });
    const sheet = nome => wb.Sheets[nome]
      ? XLSX.utils.sheet_to_json(wb.Sheets[nome], { defval: '' })
      : null;
    const rowsT = sheet('Terminais');
    const rowsC = sheet('Clientes');
    const rowsP = sheet('Placas');
    if (rowsT?.length) {
      terminaisCad = rowsT.map(xlsxMapTerminal);
      renderTerminais(); atualizarDropdownsTerminais();
    }
    if (rowsC?.length) {
      clientes = rowsC.map(xlsxMapCliente);
      renderClientes(); atualizarDropdownsClientes();
    }
    if (rowsP?.length) {
      veiculos = rowsP.map(xlsxMapPlaca);
      renderVeiculos();
    }
    spSetStatus('ok', `✓ dados.xlsx carregado`);
    return true;
  } catch (e) {
    console.warn('[dados.xlsx]', e.message);
    return false;
  }
}
// ── Leitura dos JSONs (fallback) ──────────────────────────────────────────────
async function carregarDosJsons() {
  async function lerJson(arq) {
    try {
      const r = await fetch(arq);
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      return await r.json();
    } catch (e) { console.warn(`[${arq}]`, e.message); return null; }
  }
  const [listaT, listaC, listaP] = await Promise.all([
    lerJson('terminais.json'), lerJson('clientes.json'), lerJson('placas.json'),
  ]);
  if (listaT?.length) {
    terminaisCad = listaT.map((t, i) => ({ ...t, id: t.id ?? i + 1 }));
    renderTerminais(); atualizarDropdownsTerminais();
  }
  if (listaC?.length) {
    clientes = listaC.map((c, i) => ({ ...c, id: c.id ?? 100 + i }));
    renderClientes(); atualizarDropdownsClientes();
  }
  if (listaP?.length) {
    veiculos = listaP.map((v, i) => ({
      ...v, id: v.id ?? 200 + i,
      capacidade: v.compartimentos.reduce((s, c) => s + c.cap, 0),
      jornadaMin: duracaoJornadaMin(v.jornadaInicio || '06:00', v.jornadaFim || '18:00'),
    }));
    renderVeiculos();
  }
}
// ── Ponto de entrada ──────────────────────────────────────────────────────────
async function carregarDadosFixos() {
  const ok = await carregarDoExcel('dados.xlsx');
  if (!ok) await carregarDosJsons();
}
// ── Gera e baixa o modelo Excel vazio com cabeçalhos ─────────────────────────
function baixarModeloExcel() {
  if (typeof XLSX === 'undefined') { alert('SheetJS não carregado.'); return; }
  const wb = XLSX.utils.book_new();
  const addSheet = (nome, headers) => {
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    ws['!cols'] = headers.map(() => ({ wch: 22 }));
    XLSX.utils.book_append_sheet(wb, ws, nome);
  };
  addSheet('Terminais', [
    'Nome', 'Cidade', 'Distribuidora', 'Latitude', 'Longitude', 'Fuso',
    'Tempo Carregamento (min)', 'Abertura Padrão', 'Fechamento Padrão', 'Dias Ativos',
  ]);
  addSheet('Clientes', [
    'Código SAP', 'Nome', 'Cidade', 'Latitude', 'Longitude',
    'Tempo Descarga (min)', 'Restrição Horário', 'Tipos Caminhão', 'Observações',
  ]);
  addSheet('Placas', [
    'Placa', 'Implemento', 'Transportadora', 'Tipo', 'Terminal', 'Cidade Base', 'Turno',
    'Jornada Início', 'Jornada Fim', 'Vel Carregado (km/h)', 'Vel Vazio (km/h)',
    'C1', 'C1 Produto', 'C2', 'C2 Produto', 'C3', 'C3 Produto', 'C4', 'C4 Produto',
  ]);
  XLSX.writeFile(wb, 'dados_modelo.xlsx');
}
// ╚══════════════════════════════════════════════════════════════════════════════╝
// ─── Utilitários ──────────────────────────────────────────────────────────────
function nomeTerminais() { return terminaisCad.map(t => t.nome); }
function showTab(name) {
  const tabNames = ['terminais','clientes','pedidos','veiculos','resultado','mapa','operacao','dashboard_rot','frete','historico'];
  const routeRoot = document.getElementById('roteirizador-shell') || document;
  routeRoot.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', tabNames[i] === name));
  routeRoot.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  routeRoot.querySelector('#tab-' + name)?.classList.add('active');
  if (name === 'operacao') renderTemplateOperacao();
  if (name === 'frete') {
    freteRenderContratos();
    freteRenderSpot();
    freteSetVista(_freteVista);
    const hoje = new Date();
    const y = hoje.getFullYear(), m = String(hoje.getMonth()+1).padStart(2,'0');
    const de  = document.getElementById('frete-de');
    const ate = document.getElementById('frete-ate');
    if (de  && !de.value)  de.value  = y + '-' + m + '-01';
    if (ate && !ate.value) ate.value = hoje.toISOString().slice(0,10);
  }
  if (name === 'mapa') {
    renderMapaGeral();
    // Garante invalidateSize após o CSS aplicar display:block
    setTimeout(() => { try { if (mapaGeral) mapaGeral.invalidateSize(); } catch(e){} }, 120);
    setTimeout(() => { try { if (mapaGeral) mapaGeral.invalidateSize(); } catch(e){} }, 400);
  }
}
function exportarDadosPainel() {
  try {
    if (typeof XLSX === 'undefined') {
      alert('SheetJS não carregado para exportação Excel.');
      return;
    }
    const wb = XLSX.utils.book_new();
    const termHeaders = [
      'Nome', 'Cidade', 'Distribuidora', 'Empresa Local Expedição', 'Latitude', 'Longitude', 'Fuso',
      'Tempo Carregamento (min)', 'Abertura Padrão', 'Fechamento Padrão', 'Dias Ativos',
    ];
    const cliHeaders = [
      'Código SAP', 'Nome', 'Cidade', 'Latitude', 'Longitude',
      'Tempo Descarga (min)', 'Restrição Horário', 'Tipos Caminhão', 'Observações', 'Identidade Petronas',
    ];
    const placaHeaders = [
      'Placa', 'Implemento', 'Transportadora', 'Tipo', 'Disponibilidade', 'Contrato',
      'Motorista Diurno', 'Motorista Noturno',
      'Terminal', 'Cidade Base', 'Turno',
      'Jornada Início', 'Jornada Fim', 'Vel Carregado (km/h)', 'Vel Vazio (km/h)',
      'Tempo Refeicao (min)',
      'Identidade Petronas',
      'C1', 'C1 Produto', 'C2', 'C2 Produto', 'C3', 'C3 Produto', 'C4', 'C4 Produto',
      'C5', 'C5 Produto', 'C6', 'C6 Produto', 'C7', 'C7 Produto', 'C8', 'C8 Produto',
    ];
    const rowsTerm = (terminaisCad || []).map(t => ({
      'Nome': t.nome || '',
      'Cidade': t.cidade || '',
      'Distribuidora': t.distribuidora || '',
      'Empresa Local Expedição': t.empresaLocalExpedicao || '',
      'Latitude': t.lat ?? '',
      'Longitude': t.lon ?? '',
      'Fuso': t.fuso || '-3',
      'Tempo Carregamento (min)': t.tempoCarregamentoMedioMin ?? 60,
      'Abertura Padrão': t.aberturaPadrao || '06:00',
      'Fechamento Padrão': t.fechamentoPadrao || '18:00',
      'Dias Ativos': (t.diasAtivos || []).join(','),
    }));
    const rowsCli = (clientes || []).map(c => ({
      'Código SAP': c.codigoSAP || '',
      'Nome': c.nome || '',
      'Cidade': c.cidade || '',
      'Latitude': c.lat ?? '',
      'Longitude': c.lon ?? '',
      'Tempo Descarga (min)': c.tempoDescargaMediaMin ?? 45,
      'Restrição Horário': c.restricaoHorario || '',
      'Tipos Caminhão': (c.tiposCaminhao || []).join(';'),
      'Observações': c.observacoes || '',
      'Identidade Petronas': c.identidadePetronas ? 'Sim' : 'Não',
    }));
    const rowsPlaca = (veiculos || []).map(v => {
      const row = {
        'Placa': v.placa || '',
        'Implemento': v.implemento || '',
        'Transportadora': v.transportadora || '',
        'Tipo': v.tipo || '',
        'Disponibilidade': v.disponibilidade || 'Disponível',
        'Contrato': v.contrato || 'Dedicado',
        'Motorista Diurno': v.motoristaDiurno || '',
        'Motorista Noturno': v.motoristaNt || '',
        'Terminal': v.terminal || '',
        'Cidade Base': v.cidadeBase || '',
        'Turno': v.turno || 'personalizado',
        'Jornada Início': v.jornadaInicio || '06:00',
        'Jornada Fim': v.jornadaFim || '18:00',
        'Vel Carregado (km/h)': v.velMediaCarregado ?? 45,
        'Vel Vazio (km/h)': v.velMediaVazio ?? 55,
        'Tempo Refeicao (min)': v.tempoPerdidoMin ?? 0,
        'Identidade Petronas': v.identidadePetronas ? 'Sim' : 'Não',
      };
      const comps = v.compartimentos || [];
      for (let i = 1; i <= 8; i++) {
        const c = comps[i - 1];
        row[`C${i}`] = c?.cap ?? '';
        row[`C${i} Produto`] = c?.produto || '';
      }
      return row;
    });
    const wsTerm = XLSX.utils.json_to_sheet(rowsTerm, { header: termHeaders });
    const wsCli = XLSX.utils.json_to_sheet(rowsCli, { header: cliHeaders });
    const wsPlacas = XLSX.utils.json_to_sheet(rowsPlaca, { header: placaHeaders });
    wsTerm['!cols'] = termHeaders.map(() => ({ wch: 22 }));
    wsCli['!cols'] = cliHeaders.map(() => ({ wch: 22 }));
    wsPlacas['!cols'] = placaHeaders.map(() => ({ wch: 18 }));
    XLSX.utils.book_append_sheet(wb, wsTerm, 'Terminais');
    XLSX.utils.book_append_sheet(wb, wsCli, 'Clientes');
    XLSX.utils.book_append_sheet(wb, wsPlacas, 'Placas');
    XLSX.writeFile(wb, 'dados.xlsx');
  } catch (err) {
    console.error('Falha ao exportar dados:', err);
    alert('Não foi possível exportar os dados atuais.');
  }
}
function cancelarForm(id) { document.getElementById(id).classList.add('hidden'); }
function termTag(nome) {
  if (!nome) return '';
  return `<span class="terminal-badge">⬡ ${nome}</span>`;
}
function txt(v) { return (v || '').toString().trim().toLowerCase(); }
function containsFiltro(valor, filtro) { return !txt(filtro) || txt(valor).includes(txt(filtro)); }
function valId(id) { return document.getElementById(id)?.value || ''; }
function cidadeDoTerminal(nomeTerminal) {
  const t = terminaisCad.find(x => x.nome === nomeTerminal);
  return t?.cidade || '';
}
function distribuidoraDoTerminal(nomeTerminal) {
  const t = terminaisCad.find(x => x.nome === nomeTerminal);
  return t?.distribuidora || '';
}
function normalizarCidade(txtCidade) {
  return (txtCidade || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s*(\/|-)\s*[a-z]{2}$/i, '')
    .replace(/[^a-z0-9]/g, '');
}
function terminalPorNome(nome) {
  const alvo = (nome || '').toString().trim().toLowerCase();
  if (!alvo) return null;
  return terminaisCad.find(t => (t.nome || '').toString().trim().toLowerCase() === alvo) || null;
}
function cidadeBaseVeiculo(v) {
  return (v?.cidadeBase || v?.cidade || cidadeDoTerminal(v?.terminal) || '').toString().trim();
}
function veiculoAtendeTerminal(v, terminalNome) {
  // Pedido sem terminal → aceita qualquer veículo
  // (a roteirização usará a cidade base do veículo como origem)
  if (!terminalNome) return true;
  // Veículo tem terminal fixo cadastrado → match exato
  if (v?.terminal) {
    if (v.terminal === terminalNome) return true;
    const existeTerminalVeiculo = !!terminalPorNome(v.terminal);
    // Terminal do veículo existe no cadastro mas é diferente → não atende
    if (existeTerminalVeiculo) return false;
    // Terminal do veículo não existe no cadastro → trata como cidade base abaixo
  }
  // Fallback por cidade: compara cidade base do veículo com cidade do terminal do pedido
  const cidadeV = normalizarCidade(cidadeBaseVeiculo(v));
  const cidadeT = normalizarCidade(cidadeDoTerminal(terminalNome));
  // Se o terminal do pedido não está cadastrado, tenta normalizar diretamente
  // o nome do terminal como referência de cidade (ex: "Paulínia TORRÃO Nexta" → "paulinia")
  if (!cidadeT) {
    const cidadeTerminalDireto = normalizarCidade(terminalNome.split(' ')[0]);
    return !!cidadeV && !!cidadeTerminalDireto && cidadeV.startsWith(cidadeTerminalDireto);
  }
  return !!cidadeV && !!cidadeT && cidadeV === cidadeT;
}
function baseVeiculoLabel(v) {
  if (v?.terminal) return v.terminal;
  const cidade = cidadeBaseVeiculo(v);
  return cidade ? `Cidade ${cidade}` : 'Sem base';
}
function criarCompsDisp(v) {
  return (v.compartimentos || []).map((c, i) => ({
    ...c,
    disponivel: c.cap,
    cpt: i + 1,
  }));
}
function volInt(v) { return Math.round((Number(v) || 0) * 1000); }
function compElegivelProduto(c, produto) {
  return (c.disponivel || 0) > 0 && (!c.produto || c.produto === produto);
}
function selecionarCompartimentos(comps, volumeAlvo, modo='exact') {
  // Ordena decrescente: tenta compartimentos maiores primeiro, evitando combinar
  // pequenos quando um único grande já resolve (ex: evita 2+3 antes de encontrar 5).
  const elegiveis = (comps || []).filter(c => (c.disponivel || 0) > 0)
    .sort((a, b) => (b.disponivel || 0) - (a.disponivel || 0));
  const n = elegiveis.length;
  if (!n) return null;
  const alvo = volInt(volumeAlvo);
  let melhor = null;
  let melhorSoma = -1;
  let melhorGeq = null;      // geq: menor soma que ainda >= alvo
  let melhorSomaGeq = Infinity;
  const total = 1 << n;
  for (let mask = 1; mask < total; mask++) {
    let soma = 0;
    const subset = [];
    for (let i = 0; i < n; i++) {
      if ((mask & (1 << i)) !== 0) {
        const c = elegiveis[i];
        const capI = volInt(c.disponivel);
        if (capI <= 0) continue;
        soma += capI;
        subset.push(c);
      }
    }
    if (!subset.length) continue;
    if (modo === 'exact') {
      if (soma === alvo) return subset;
      continue;
    }
    if (modo === 'geq') {
      // Menor combinação de capacidade >= volume (desperdício mínimo)
      if (soma >= alvo && soma < melhorSomaGeq) {
        melhorSomaGeq = soma;
        melhorGeq = subset;
      }
      continue;
    }
    if (soma <= alvo && soma > melhorSoma) {
      melhorSoma = soma;
      melhor = subset;
    }
  }
  if (modo === 'geq') return melhorGeq;
  return modo === 'maxLE' ? melhor : null;
}
function atualizarFiltroMapaTransportadora() {
  const el = document.getElementById('f-map-transp');
  if (!el) return;
  const atual = el.value || '';
  const transportadoras = [...new Set(
    (veiculos || [])
      .map(v => (v.transportadora || '').toString().trim())
      .filter(Boolean)
  )].sort((a,b) => a.localeCompare(b, 'pt-BR'));
  el.innerHTML = '<option value="">Todas transportadoras</option>'
    + transportadoras.map(t => `<option value="${t.replace(/"/g, '&quot;')}">${t}</option>`).join('');
  if (atual && transportadoras.includes(atual)) el.value = atual;
}
function onFiltroMapaTransportadoraChange() {
  renderMapaGeral();
}
// ── Filtro multi-terminal do Resumo Operação ──────────────────────────────────
function renderFiltroTerminaisMapa() {
  const panel = document.getElementById('term-filter-panel');
  const label = document.getElementById('term-filter-label');
  const btn   = document.getElementById('term-filter-btn');
  if (!panel || !label || !btn) return;
  // Coleta terminais disponíveis no resultado atual
  const terminaisDisp = [...new Set(
    (veiculos || []).map(v => baseVeiculoLabel(v)).filter(Boolean)
  )].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  panel.innerHTML = terminaisDisp.length ? terminaisDisp.map(t => `
    <label class="term-filter-item">
      <input type="checkbox" onchange="toggleFiltroTerminalMapa('${t.replace(/'/g,"\\'")}',this.checked)"
        ${filtroMapaTerminais.has(t) ? 'checked' : ''}>
      <span>${t}</span>
    </label>`).join('') : '<div style="padding:8px 12px;font-size:12px;color:var(--text-3);">Sem terminais</div>';
  const n = filtroMapaTerminais.size;
  label.textContent = n === 0 ? 'Todos os terminais' : `${n} terminal${n > 1 ? 'is' : ''} selecionado${n > 1 ? 's' : ''}`;
  btn.classList.toggle('tem-filtro', n > 0);
}
function toggleFiltroTerminalMapa(terminal, checked) {
  if (checked) filtroMapaTerminais.add(terminal);
  else filtroMapaTerminais.delete(terminal);
  renderFiltroTerminaisMapa();
  renderMapaGeral();
}
function toggleTermFiltroPanel(event) {
  event.stopPropagation();
  const panel = document.getElementById('term-filter-panel');
  const btn   = document.getElementById('term-filter-btn');
  if (!panel) return;
  const aberto = panel.classList.toggle('aberto');
  btn?.classList.toggle('aberto', aberto);
  if (aberto) renderFiltroTerminaisMapa();
}
// Fecha ao clicar fora
document.addEventListener('click', (e) => {
  const wrap = document.querySelector('.term-filter-wrap');
  if (wrap && !wrap.contains(e.target)) {
    document.getElementById('term-filter-panel')?.classList.remove('aberto');
    document.getElementById('term-filter-btn')?.classList.remove('aberto');
  }
});
function limparFiltros(aba) {
  const mapa = {
    terminais: ['f-term-terminal','f-term-cidade'],
    clientes: ['f-cli-cliente','f-cli-cidade'],
    pedidos: ['f-ped-cliente','f-ped-cidade','f-ped-terminal'],
    veiculos: ['f-vei-placa','f-vei-transp','f-vei-terminal','f-vei-cidade','f-vei-tipo'],
    resultado: ['f-res-cliente','f-res-cidade','f-res-terminal','f-res-placa','f-res-transp'],
    operacao: ['f-op-cliente','f-op-cidade','f-op-terminal','f-op-placa','f-op-transp'],
    mapa: ['f-map-transp'],
  };
  (mapa[aba] || []).forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  if (aba === 'terminais') renderTerminais();
  if (aba === 'clientes') renderClientes();
  if (aba === 'pedidos') renderPedidos();
  if (aba === 'veiculos') {
    // Usa data de operação do header; fallback para hoje
    const inputData = document.getElementById('rot-data-operacao')?.value;
    const ds = inputData || (() => { const h = new Date(); return `${h.getFullYear()}-${String(h.getMonth()+1).padStart(2,'0')}-${String(h.getDate()).padStart(2,'0')}`; })();
    // Usa data do roteirizador, não do painel de disponibilidade
    const inputDataAba = document.getElementById('rot-data-operacao')?.value;
    const dsAba = inputDataAba || ds;
    sincronizarDisponibilidadeVeiculos(dsAba).then(() => renderVeiculos()).catch(() => renderVeiculos());
  }
  if (aba === 'mapa') renderMapaGeral();
}
function aplicarFiltrosResultado() {
  if (ultimoResultado) renderResultado(ultimoResultado, ultimoControleTempo || {});
}
function aplicarFiltrosOperacao() {
  if (ultimoResultado) renderTemplateOperacao();
}
function fecharModalViagem(ev=null) {
  if (ev && ev.target && ev.target.id !== 'modal-viagem') return;
  document.getElementById('modal-viagem').classList.remove('show');
  // Re-renderiza a lista de viagens para refletir imediatamente o km ajustado
  // (e outras mudanças feitas no mapa) no card de "Otimização Rotas" e,
  // por consequência, manter consistência com o template de "Envio Transportadora".
  if (ultimoResultado) renderResultado(ultimoResultado, ultimoControleTempo || {});
}
// NOTA: garantirMapaViagem() está definida no <script> do Dashboard (mais abaixo
// no documento), não aqui. Como function declarations no nível de topo de
// <script> não-module são propriedades de window, e o script do Dashboard
// carrega depois deste, window.garantirMapaViagem aponta para aquela versão
// (mais completa: zoomControl, limpeza de _mvPolylines/_mvUserWaypoints/
// _mvWaypoints e do badge de vias). Mantida apenas uma definição para evitar
// ambiguidade.
function abrirModalViagem(veiculoId, idxViagem) {
  if (!ultimoResultado || !ultimoResultado[veiculoId]) return;
  const v = veiculos.find(x => x.id === veiculoId);
  if (!v) return;
  const viagem = ultimoResultado[veiculoId][idxViagem];
  if (!viagem || !viagem.paradas || !viagem.paradas.length) return;
  const terminalNomeOrigem = viagem.terminalOrigem || v.terminal || viagem.paradas[0]?.pedido?.terminal || '';
  const terminal = terminaisCad.find(t => t.nome === terminalNomeOrigem);
  if (!terminal) { alert('Terminal de origem do veículo não possui coordenadas cadastradas.'); return; }
  const pontos = [{
    nome: `Origem: ${terminal.nome}`,
    lat: terminal.lat,
    lon: terminal.lon,
    tipo: 'origem'
  }];
  viagem.paradas.forEach((p, i) => {
    const _coord = latLonEfetivo(p.pedido);
    pontos.push({
      nome: `${i+1}. ${p.pedido.cliente}`,
      lat: _coord.lat,
      lon: _coord.lon,
      isCentroide: _coord.isCentroide,
      tipo: 'destino',
      volume: p.volumeTotal || 0,
      cidade: p.pedido.cidade || '-',
      restricao: p.pedido.restricao || '',
    });
  });
  const ultimo = viagem.paradas[viagem.paradas.length - 1];
  if ((ultimo?.deslocVazioMin || 0) > 0) {
    pontos.push({
      nome: `Retorno: ${terminal.nome}`,
      lat: terminal.lat,
      lon: terminal.lon,
      tipo: 'retorno'
    });
  }
  const validos = pontos.filter(p => !isNaN(parseFloat(p.lat)) && !isNaN(parseFloat(p.lon)));
  if (validos.length < 2) { alert('Coordenadas insuficientes para montar o mapa da viagem.'); return; }
  document.getElementById('mv-titulo').textContent = `${v.placa} · Viagem ${idxViagem+1}${v.transportadora ? ' · '+v.transportadora : ''}`;
  const vol = viagem.paradas.reduce((s,p)=>s+(p.volumeTotal||0),0);
  document.getElementById('mv-resumo').innerHTML =
    `Terminal: <strong>${terminalNomeOrigem || '-'}</strong> · Destinos: <strong>${viagem.paradas.length}</strong> · Volume: <strong>${vol.toFixed(1)} m³</strong> · Tempo: <strong>${(viagem.tempoConsumidoMin||0).toFixed(0)} min</strong>`;
  document.getElementById('mv-lista').innerHTML = viagem.paradas.map((p,i) => `
    <div class="card" style="margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <span class="tag tag-blue">${i+1}</span>
        <span style="font-weight:700;">${p.pedido.cliente}</span>
        <span style="color:#4A6535;">${p.pedido.cidade||'-'}</span>
        ${p.pedido.restricao ? `<span class="tag tag-yellow">${p.pedido.restricao}</span>` : ''}
        <span class="tag tag-lime">${(p.volumeTotal||0).toFixed(1)} m³</span>
        ${(() => { const _c = latLonEfetivo(p.pedido); return _c.isCentroide ? `<span style="font-size:11px;color:#92400E;">📍 ~centróide ${p.pedido.cidade}</span>` : `<span style="font-size:11px;color:#8AAA70;">${_c.lat.toFixed(5)}, ${_c.lon.toFixed(5)}</span>`; })()}
      </div>
    </div>
  `).join('');
  document.getElementById('modal-viagem').classList.add('show');
  garantirMapaViagem();
  dadosMapaAtual = validos;
  window._mvVeiculoId  = veiculoId;
  window._mvIdxViagem  = idxViagem;
  renderCustoMapaViagem();
  // Inicializa waypoints arrastáveis
  _mvWaypoints = validos.map((p, i) => ({ ...p, idx: i, marker: null }));
  _mvPolylines = [];
  // Restaurar vias salvas anteriormente para esta viagem.
  // Prioriza o dado persistido no próprio objeto da viagem (viagem._userWaypoints),
  // que é a fonte confiável e sobrevive mesmo se o cache em memória (_mvRotasSalvas)
  // tiver sido perdido (ex.: viagem carregada do histórico, página recarregada).
  const chaveAtual = `${veiculoId}||${idxViagem}`;
  const waypointsPersistidos = (viagem._userWaypoints && viagem._userWaypoints.length)
    ? viagem._userWaypoints
    : _mvRotasSalvas[chaveAtual]?.userWaypoints;
  if (waypointsPersistidos?.length) {
    _mvUserWaypoints = waypointsPersistidos.map(uw => ({ ...uw, marker: null }));
    _mvAtualizarBadgeVias();
  }
  const latlngs = validos.map(p => [p.lat, p.lon]);
  // Desenha rota arrastável (sem bloquear UI)
  mvDesenharRota().catch(() => {});
  // Distâncias para popups (em background) — não desenha polylines aqui, mvDesenharRota faz isso
  const distAcum = [];
  // Calcula horário absoluto de chegada para cada parada
  const _jIniRaw3 = parseHoraMin(v.jornadaInicio || '06:00');
  let jIniMin = isNaN(_jIniRaw3) ? 360 : _jIniRaw3;
  if (v._horarioDisponivelAPartirDe) { const _dm3 = parseHoraMin(v._horarioDisponivelAPartirDe); if (!isNaN(_dm3) && _dm3 > jIniMin) jIniMin = _dm3; }
  const viagens = (ultimoResultado[v.id] || []).filter(vi => !vi._vazio && vi.paradas?.length);
  let clock = inicioViagemAbsMin(viagens, viagens.indexOf(viagem), jIniMin, v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1);
  clock += (viagem.esperaTerminalMin || 0);
  function absMinToHora(m) {
    const d = Math.floor(m / 1440);
    const hm = ((m % 1440) + 1440) % 1440;
    const h = String(Math.floor(hm / 60)).padStart(2, '0');
    const min = String(hm % 60).padStart(2, '0');
    return d > 0 ? `${h}:${min} (+${d}d)` : `${h}:${min}`;
  }
  validos.forEach((p, idx) => {
    if (idx > 0 && viagem.paradas[idx - 1]) {
      const par = viagem.paradas[idx - 1];
      if (idx === 1) clock += (par.tempoCarregamentoMin || 0);
      clock += (par.deslocCarregadoMin || 0);
      clock += (par.tempoEsperaRestricaoMin || 0);
      clock += (par.tempoDescargaMin || 0);
      clock += (par.deslocVazioMin || 0);
    }
    const isOrigem  = p.tipo === 'origem';
    const isRetorno = p.tipo === 'retorno';
    const seq = isOrigem ? '⬤' : isRetorno ? '↩' : String(idx);
    const bgColor = isOrigem ? '#111827' : isRetorno ? '#6B7280' : '#4F46E5';
    const icon = L.divIcon({
      className: '',
      html: `<div style="width:28px;height:28px;border-radius:50%;background:${bgColor};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,0.35);font-family:Inter,sans-serif;">${seq}</div>`,
      iconSize: [28, 28], iconAnchor: [14, 14],
    });
    // Distância acumulada do terminal até este ponto
    const kmAcum = distAcum[idx] != null ? distAcum[idx].toFixed(1) : null;
    const distLabel = kmAcum ? `<div style="font-size:11px;color:#2563EB;margin-bottom:6px;">📏 <strong>${kmAcum} km</strong> do terminal</div>` : '';
    let popupHtml = '';
    if (isOrigem) {
      popupHtml = `
        <div style="font-family:Inter,sans-serif;min-width:200px;">
          <div style="font-weight:700;font-size:13px;margin-bottom:4px;">🏭 ${p.nome}</div>
          <div style="font-size:11px;color:#6B7280;">Partida: <strong>${absMinToHora(clock)}</strong></div>
        </div>`;
    } else if (isRetorno) {
      const kmTotal = distAcum[idx] != null ? distAcum[idx].toFixed(1) : '-';
      popupHtml = `
        <div style="font-family:Inter,sans-serif;min-width:200px;">
          <div style="font-weight:700;font-size:13px;margin-bottom:4px;">↩ ${p.nome}</div>
          <div style="font-size:11px;color:#6B7280;">Retorno estimado: <strong>${absMinToHora(clock)}</strong></div>
          <div style="font-size:11px;color:#2563EB;margin-top:4px;">📏 Percurso total: <strong>${kmTotal} km</strong></div>
        </div>`;
    } else {
      const par = viagem.paradas[idx - 1];
      const vol = (par?.volumeTotal || 0).toFixed(1);
      // Distância deste segmento (entre ponto anterior e este)
      const kmSeg = (distAcum[idx] != null && distAcum[idx-1] != null)
        ? (distAcum[idx] - distAcum[idx-1]).toFixed(1) : null;
      const distSegLabel = kmSeg
        ? `<div style="font-size:11px;color:#6B7280;">↳ ${kmSeg} km do ponto anterior</div>` : '';
      const produtos = (par?.pedido?.produtos || []).map(pr => {
        const nomeProd = (pr.produto || pr.nome || '').replace(/^\d+\s*-\s*/, '');
        return `<div style="display:flex;justify-content:space-between;gap:12px;padding:2px 0;border-bottom:1px solid #F3F4F6;">
          <span style="color:#374151;">${nomeProd}</span>
          <strong style="color:#111827;">${pr.volume} m³</strong>
        </div>`;
      }).join('') || `<div style="color:#374151;">Volume total: <strong>${vol} m³</strong></div>`;
      const janela = par?.pedido?.restricao
        ? `<div style="margin-top:6px;padding:4px 8px;background:#FFFBEB;color:#92400E;border-radius:4px;font-size:10px;border:1px solid #FDE68A;">⏱ Janela: ${par.pedido.restricao}</div>` : '';
      const descarga = par?.tempoDescargaMin
        ? `<div style="font-size:11px;color:#6B7280;margin-top:2px;">⏳ Descarga: ~${par.tempoDescargaMin} min</div>` : '';
      const espera = par?.tempoEsperaRestricaoMin > 0
        ? `<div style="font-size:11px;color:#D97706;">⏸ Espera janela: ${par.tempoEsperaRestricaoMin} min</div>` : '';
      const centroide = p.isCentroide
        ? `<div style="font-size:10px;color:#92400E;margin-top:4px;">⚠ Posição aproximada (centróide)</div>` : '';
      const contrato = par?.pedido?.contrato
        ? `<div style="font-size:11px;color:#6B7280;">Contrato: ${par.pedido.contrato}</div>` : '';
      popupHtml = `
        <div style="font-family:Inter,sans-serif;width:260px;">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E5E7EB;">
            <div style="width:24px;height:24px;border-radius:50%;background:#4F46E5;color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;">${idx}</div>
            <div>
              <div style="font-weight:700;font-size:13px;line-height:1.2;">${p.nome.replace(/^\d+\.\s*/, '')}</div>
              <div style="font-size:11px;color:#6B7280;">${p.cidade || ''}</div>
            </div>
          </div>
          <div style="margin-bottom:6px;">
            <div style="font-size:12px;font-weight:600;color:#111827;">🕐 Chegada estimada: ${absMinToHora(clock)}</div>
            ${distLabel}${distSegLabel}
          </div>
          <div style="background:#F9FAFB;border-radius:6px;padding:8px;margin-bottom:6px;font-size:11px;">
            <div style="font-weight:600;color:#6B7280;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px;">Carga entregue</div>
            ${produtos}
          </div>
          ${contrato}${descarga}${espera}${janela}${centroide}
        </div>`;
    }
    // Marcador arrastável — atualiza waypoint e redesenha rota ao soltar
    const icon_drag = L.divIcon({
      className: '',
      html: `<div style="width:28px;height:28px;border-radius:50%;background:${bgColor};color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;border:2.5px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,0.4);cursor:grab;font-family:Inter,sans-serif;">${seq}</div>`,
      iconSize: [28,28], iconAnchor: [14,14]
    });
    const marker = L.marker([p.lat, p.lon], { icon: icon_drag, draggable: true }).addTo(camadaViagem);
    marker.bindPopup(popupHtml, { maxWidth: 300 });
    marker.on('dragend', async (e) => {
      const ll = e.target.getLatLng();
      _mvWaypoints[idx].lat = ll.lat;
      _mvWaypoints[idx].lon = ll.lng;
      await mvDesenharRota();
    });
    marker.on('click', () => marker.openPopup());
    if (_mvWaypoints[idx]) _mvWaypoints[idx].marker = marker;
  });
  mapaViagem.fitBounds(latlngs, { padding: [32, 32] });
  // OBS: não chamar mvRecalcularDistancia() aqui — mvDesenharRota() (chamada acima,
  // logo após montar os waypoints) já recalcula e exibe a distância ao final.
  // Chamar de novo aqui criava uma segunda execução concorrente que invalidava
  // o token da primeira antes dela terminar de desenhar a polyline, fazendo o
  // traçado nunca aparecer (apenas o km, calculado por esta segunda chamada).
}
function abrirNavegacaoExterna() {
  if (!dadosMapaAtual || dadosMapaAtual.length < 2) return;
  const origem = dadosMapaAtual[0];
  const destino = dadosMapaAtual[dadosMapaAtual.length - 1];
  const waypoints = dadosMapaAtual.slice(1, -1).map(p => `${p.lat},${p.lon}`).join('|');
  const url = `https://www.google.com/maps/dir/?api=1&origin=${origem.lat},${origem.lon}&destination=${destino.lat},${destino.lon}&travelmode=driving${waypoints ? `&waypoints=${encodeURIComponent(waypoints)}` : ''}`;
  window.open(url, '_blank');
}
function corPorPlaca(placa) {
  const paleta = ['#00A499','#20419A','#763F98','#f28e2b','#e15759','#59a14f','#edc948','#4e79a7','#9c755f','#ff9da7','#2f4b7c','#9c27b0'];
  const s = (placa || '').toString().toUpperCase();
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = ((hash << 5) - hash) + s.charCodeAt(i);
  return paleta[Math.abs(hash) % paleta.length];
}
function garantirMapaGeral() {
  if (!mapaGeral) {
    mapaGeral = L.map('mapa-geral-mapa');
    L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap &copy; CARTO'
    }).addTo(mapaGeral);
    mapaGeral.setView([-14.235, -51.925], 4);
  }
  if (camadaMapaGeral) camadaMapaGeral.remove();
  camadaMapaGeral = L.layerGroup().addTo(mapaGeral);
  mapaGeral.invalidateSize(); // síncrono — garante tamanho correto antes do fitBounds
}
function renderMapaGeral() {
  const top = document.getElementById('mapa-geral-top');
  const legenda = document.getElementById('mapa-geral-legenda');
  const kpisPedidos = document.getElementById('mapa-geral-kpis-pedidos');
  const kpisPlacas = document.getElementById('mapa-geral-kpis-placas');
  if (!top || !legenda || !kpisPedidos || !kpisPlacas) return;
  if (!ultimoResultado) {
    top.textContent = 'Execute a otimização para visualizar as rotas consolidadas.';
    kpisPedidos.innerHTML = '';
    kpisPlacas.innerHTML = '';
    legenda.innerHTML = '<div class="empty" style="padding:16px;">Sem rotas para exibir.</div>';
    return;
  }
  if (typeof L === 'undefined') {
    top.textContent = 'Mapa indisponível: biblioteca Leaflet não carregada.';
    kpisPedidos.innerHTML = '';
    kpisPlacas.innerHTML = '';
    legenda.innerHTML = '<div class="empty" style="padding:16px;">Falha ao carregar mapa.</div>';
    return;
  }
  garantirMapaGeral();
  const viagensMapa = [];
  const bounds = [];
  const porPlaca = new Map();
  atualizarFiltroMapaTransportadora();
  const transpHeader = valId('f-map-transp');
  const veiculosEscopo = veiculos.filter(v => {
    if (txt(transpHeader) && txt(v.transportadora) !== txt(transpHeader)) return false;
    if (filtroMapaTerminais.size > 0 && !filtroMapaTerminais.has(baseVeiculoLabel(v))) return false;
    return true;
  });
  renderFiltroTerminaisMapa(); // atualiza chips/label com terminais disponíveis
  const placasAlocadasSet = new Set();
  const placasProblemaSet = new Set();
  const alocadoPorPedidoGlobal = new Map();
  let _idxViagemGlobal = 0; // contador global de viagens para coloração única
  veiculos.forEach(vg => {
    const viagensG = (ultimoResultado[vg.id] || []).filter(vi => (vi.paradas || []).length);
    viagensG.forEach(vi => {
      (vi.paradas || []).forEach(p => {
        alocadoPorPedidoGlobal.set(p.pedido.id, (alocadoPorPedidoGlobal.get(p.pedido.id) || 0) + (p.volumeTotal || 0));
      });
    });
  });
  veiculosEscopo.forEach(v => {
    const viagens = (ultimoResultado[v.id] || []).filter(vi => (vi.paradas || []).length);
    if (viagens.length) placasAlocadasSet.add(v.id);
    viagens.forEach(vi => {
      const vol = (vi.paradas || []).reduce((s, p) => s + (p.volumeTotal || 0), 0);
      if ((v.capacidade || 0) > 0 && vol > 0 && vol < (v.capacidade - 0.01)) placasProblemaSet.add(v.id);
    });
    viagens.forEach((vi, idx) => {
      const paradasRaw = vi.paradas || [];
      const paradas = paradasRaw.filter(p => { const _c = latLonEfetivo(p.pedido); return !isNaN(_c.lat) && !isNaN(_c.lon); });
      if (!paradas.length) return;
      const terminalNome = vi.terminalOrigem || v.terminal || paradas[0]?.pedido?.terminal || '';
      const terminal = terminaisCad.find(t => t.nome === terminalNome && !isNaN(parseFloat(t.lat)) && !isNaN(parseFloat(t.lon)));
      const pontos = [];
      if (terminal) {
        pontos.push({ lat: parseFloat(terminal.lat), lon: parseFloat(terminal.lon), tipo: 'origem', nome: terminal.nome });
      }
      // Monta sequência intercalando waypoints manuais (vi._userWaypoints) salvos
      // pelo usuário ao ajustar a rota no mapa individual da viagem.
      // A estrutura é idêntica à usada em _mvMontarSequencia() no mapa de viagem:
      // cada ponto fixo (origem→parada1→parada2→…) pode ter vias intermediárias
      // inseridas pelo usuário (segmento = índice do ponto fixo de origem do segmento).
      const userWps = Array.isArray(vi._userWaypoints) ? vi._userWaypoints : [];
      // pontos fixos: índice 0 = terminal (origem), 1…n = paradas, último = retorno
      const pontosFixos = pontos; // reutiliza o array já iniciado com a origem acima
      paradas.forEach((p, pi) => {
        const _coord = latLonEfetivo(p.pedido);
        pontosFixos.push({
          lat: _coord.lat,
          lon: _coord.lon,
          isCentroide: _coord.isCentroide,
          tipo: 'destino',
          nome: `${pi + 1}. ${p.pedido.cliente}`,
          cidade: p.pedido.cidade || '',
        });
      });
      if (terminal && ((paradas[paradas.length - 1]?.deslocVazioMin || 0) > 0)) {
        pontosFixos.push({ lat: parseFloat(terminal.lat), lon: parseFloat(terminal.lon), tipo: 'retorno', nome: `Retorno ${terminal.nome}` });
      }
      if (pontosFixos.length < 2) return;
      // Intercalar vias manuais do usuário entre os pontos fixos para desenho da rota
      const pontosRota = [];
      for (let fi = 0; fi < pontosFixos.length; fi++) {
        pontosRota.push(pontosFixos[fi]);
        const viasDoSegmento = userWps
          .filter(uw => uw.segmento === fi)
          .sort((a, b) => a.ordem - b.ordem);
        viasDoSegmento.forEach(uw => pontosRota.push({ lat: uw.lat, lon: uw.lon, tipo: 'via' }));
      }
      // pontosRota → desenho da linha (inclui desvios manuais)
      // pontos / pontosFixos → marcadores dos clientes (sem vias intermediárias)
      const _paletaViagem = ['#00A499','#20419A','#763F98','#f28e2b','#e15759','#59a14f','#edc948','#4e79a7','#9c755f','#ff9da7','#e76f51','#2f4b7c','#a8dadc','#e63946','#06d6a0','#118ab2'];
      const cor = _paletaViagem[_idxViagemGlobal % _paletaViagem.length];
      _idxViagemGlobal++;
      const volViagem = paradas.reduce((s, p) => s + (p.volumeTotal || 0), 0);
      viagensMapa.push({ v, idx, pontos, pontosRota, cor, volViagem, terminalNome, destinos: paradas.length, viagem: vi, paradas });
      const cur = porPlaca.get(v.placa) || { cor, ciclos: 0, destinos: 0, volume: 0, transportadora: v.transportadora || '-' };
      cur.ciclos += 1;
      cur.destinos += paradas.length;
      cur.volume += volViagem;
      porPlaca.set(v.placa, cur);
    });
  });
  const qtdPlacasAlocadas = placasAlocadasSet.size;
  const qtdPlacasSemAloc = Math.max(0, veiculosEscopo.length - qtdPlacasAlocadas);
  const qtdPlacasProblema = placasProblemaSet.size;
  const demandaTotalM3 = (pedidos || []).reduce((s, p) => s + totalVolPedido(p), 0);
  const transportadoM3 = [...alocadoPorPedidoGlobal.values()].reduce((s, v) => s + (v || 0), 0);
  const pctAtendimentoDemanda = demandaTotalM3 > 0 ? Math.round((transportadoM3 / demandaTotalM3) * 100) : 0;
  const totalPedidos = (pedidos || []).length;
  const pedidosComProblemaLista = (pedidos || []).filter(p => (alocadoPorPedidoGlobal.get(p.id) || 0) < totalVolPedido(p) - 0.01);
  const pedidosComProblema = pedidosComProblemaLista.length;
  const pedidosAtendidos = Math.max(0, totalPedidos - pedidosComProblema);
  const esc = (s) => (s || '').toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  const limitarLinhas = (linhas, max=10) => {
    if (!linhas.length) return ['Sem detalhes no filtro atual.'];
    if (linhas.length <= max) return linhas;
    return [...linhas.slice(0, max), `... +${linhas.length - max} item(ns)`];
  };
  const kpiCard = (valor, label, detalheLinhas, warn=false) => `
    <div class="mapa-geral-kpi has-tip ${warn ? 'warn' : ''}">
      <div class="v">${valor}</div>
      <div class="l">${label}</div>
      <div class="mapa-geral-kpi-tip">${limitarLinhas(detalheLinhas).map(esc).join('\n')}</div>
    </div>
  `;
  const semAlocPorTransp = new Map();
  veiculosEscopo
    .filter(v => !placasAlocadasSet.has(v.id))
    .forEach(v => {
      const nomeT = (v.transportadora || 'Sem transportadora').toString().trim() || 'Sem transportadora';
      semAlocPorTransp.set(nomeT, (semAlocPorTransp.get(nomeT) || 0) + 1);
    });
  const placasSemAlocLinhas = [...semAlocPorTransp.entries()]
    .sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
    .map(([t, q]) => `${t}: ${q} placa(s)`);
  const alocPorTransp = new Map();
  veiculosEscopo
    .filter(v => placasAlocadasSet.has(v.id))
    .forEach(v => {
      const nomeT = (v.transportadora || 'Sem transportadora').toString().trim() || 'Sem transportadora';
      alocPorTransp.set(nomeT, (alocPorTransp.get(nomeT) || 0) + 1);
    });
  const placasAlocLinhas = [...alocPorTransp.entries()]
    .sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
    .map(([t, q]) => `${t}: ${q} placa(s)`);
  const placasParcialLinhas = veiculosEscopo.flatMap(v => {
    const viagens = (ultimoResultado[v.id] || []).filter(vi => (vi.paradas || []).length);
    if (!viagens.length || !(v.capacidade > 0)) return [];
    const parciais = viagens
      .map(vi => {
        const vol = (vi.paradas || []).reduce((s,p) => s + (p.volumeTotal || 0), 0);
        const pct = Math.round((vol / v.capacidade) * 100);
        return pct < 100 ? pct : null;
      })
      .filter(p => p !== null);
    if (!parciais.length) return [];
    const nomeT = (v.transportadora || 'Sem transportadora').toString().trim() || 'Sem transportadora';
    return [`${nomeT} · ${v.placa}: ${parciais.map(p => `${p}%`).join(', ')}`];
  });
  const pedidosProblemaLinhas = pedidosComProblemaLista.map(p => {
    const dem = totalVolPedido(p);
    const aloc = alocadoPorPedidoGlobal.get(p.id) || 0;
    const falt = Math.max(0, dem - aloc);
    return `${p.cliente} · faltante ${falt.toFixed(1)} m³`;
  });
  const pedidosAtendidosLinhas = (pedidos || [])
    .filter(p => (alocadoPorPedidoGlobal.get(p.id) || 0) >= totalVolPedido(p) - 0.01)
    .map(p => p.cliente);
  const demandaDetalheLinhas = [
    `Demanda total: ${demandaTotalM3.toFixed(1)} m³`,
    `Transportado: ${transportadoM3.toFixed(1)} m³`,
    `Gap: ${(Math.max(0, demandaTotalM3 - transportadoM3)).toFixed(1)} m³`,
    `Atendimento: ${pctAtendimentoDemanda}%`,
  ];
  kpisPedidos.innerHTML =
    kpiCard(`${pedidosComProblema}`, 'Pedidos com problema', pedidosProblemaLinhas, pedidosComProblema > 0) +
    kpiCard(`${pedidosAtendidos}/${totalPedidos}`, 'Pedidos atendidos', pedidosAtendidosLinhas, false) +
    kpiCard(`${transportadoM3.toFixed(1)} / ${demandaTotalM3.toFixed(1)} m³`, `Demanda x transportado (${pctAtendimentoDemanda}%)`, demandaDetalheLinhas, pctAtendimentoDemanda < 100);
  kpisPlacas.innerHTML =
    kpiCard(`${qtdPlacasSemAloc}`, 'Placas sem alocação', placasSemAlocLinhas, qtdPlacasSemAloc > 0) +
    kpiCard(`${qtdPlacasAlocadas}`, 'Placas alocadas', placasAlocLinhas, false) +
    kpiCard(`${qtdPlacasProblema}`, 'Placas com alocação parcial', placasParcialLinhas, qtdPlacasProblema > 0);
  const placaFiltroAtual = (filtroMapaPlaca || '').toString().trim().toUpperCase();
  const viagensVisiveis = placaFiltroAtual
    ? viagensMapa.filter(v => (v.v?.placa || '').toString().toUpperCase() === placaFiltroAtual)
    : viagensMapa;
  if (!viagensVisiveis.length) {
    if (placaFiltroAtual) {
      filtroMapaPlaca = '';
      return renderMapaGeral();
    }
    top.textContent = 'Nenhuma rota com coordenadas válidas para desenhar no mapa.';
    legenda.innerHTML = '<div class="empty" style="padding:16px;">Sem rotas para exibir.</div>';
    mapaGeral.setView([-14.235, -51.925], 4);
    return;
  }
  viagensVisiveis.forEach(item => {
    // pontosRota inclui as vias manuais inseridas pelo usuário no mapa da viagem
    // individual; item.pontos são os pontos fixos (terminal + clientes), usados
    // abaixo apenas para posicionar os marcadores, sem as vias intermediárias.
    osrmRoute(item.pontosRota || item.pontos, camadaMapaGeral, item.cor, 4, 0.86).catch(()=>{});
    // Calcula clock de chegada para cada parada desta viagem
    const v_item = item.v;
    const vi_item = item.viagem;
    const jIniMin_item = (() => { const h = parseHoraMin(v_item.jornadaInicio||'06:00'); return isNaN(h)?360:h; })();
    const todasViagens_item = (ultimoResultado[v_item.id]||[]).filter(vi=>!vi._vazio&&(vi.paradas||[]).length);
    let clock_item = inicioViagemAbsMin(todasViagens_item, todasViagens_item.indexOf(vi_item), jIniMin_item, v_item.tempoPerdidoMin||0, doisTurnos(v_item)?2:1);
    clock_item += (vi_item.esperaTerminalMin||0);
    // clocksPorParada[i] = minutos absolutos de chegada no ponto i (0 = origem)
    const clocksPorParada = [clock_item];
    (item.paradas||[]).forEach((par, pi) => {
      if (pi === 0) clock_item += (par.tempoCarregamentoMin||0);
      clock_item += (par.deslocCarregadoMin||0) + (par.tempoEsperaRestricaoMin||0);
      clocksPorParada.push(clock_item);
      clock_item += (par.tempoDescargaMin||0) + (par.deslocVazioMin||0);
    });
    function absMinToHM(m) {
      if (m == null || isNaN(m)) return null;
      const mRound = Math.round(m); // garante inteiro — evita decimais no resultado
      const hm = ((mRound%1440)+1440)%1440;
      const d = Math.floor(mRound/1440);
      return String(Math.floor(hm/60)).padStart(2,'0')+':'+String(hm%60).padStart(2,'0')+(d>0?` (+${d}d)`:'');
    }
    item.pontos.forEach((p, i) => {
      const tipoFim  = (i === item.pontos.length - 1 && p.tipo === 'retorno');
      const isOrigem = p.tipo === 'origem';
      const seq = isOrigem ? '⬤' : tipoFim ? '↩' : String(i);
      const bg  = isOrigem ? '#111827' : tipoFim ? '#6B7280' : item.cor;
      const icon = L.divIcon({
        className: '',
        html: `<div style="width:24px;height:24px;border-radius:50%;background:${bg};color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,0.35);font-family:Inter,sans-serif;cursor:pointer;">${seq}</div>`,
        iconSize: [24, 24], iconAnchor: [12, 12],
      });
      // i=0 é origem (terminal), i>0 são paradas da viagem
      const par = (!isOrigem && !tipoFim) ? item.paradas?.[i - 1] : null;
      const produtos = par?.pedido?.produtos?.length
        ? par.pedido.produtos.map(pr => {
            const nomeProd = (pr.produto || pr.nome || '').replace(/^\d+\s*-\s*/, '');
            return `<div style="display:flex;justify-content:space-between;gap:12px;padding:2px 0;border-bottom:1px solid #F3F4F6;">
              <span style="color:#374151;">${nomeProd}</span>
              <strong style="color:#111827;">${pr.volume} m³</strong>
            </div>`;
          }).join('')
        : par?.volumeTotal
          ? `<div style="color:#374151;">Volume total: <strong>${(par.volumeTotal).toFixed(1)} m³</strong></div>`
          : '';
      // Distância haversine do terminal até este ponto (em km)
      const distKm = (() => {
        if (isOrigem || !item.pontos[0]) return null;
        const orig = item.pontos[0];
        const R = 6371;
        const dLat = (p.lat - orig.lat) * Math.PI / 180;
        const dLon = (p.lon - orig.lon) * Math.PI / 180;
        const a = Math.sin(dLat/2)**2 + Math.cos(orig.lat*Math.PI/180)*Math.cos(p.lat*Math.PI/180)*Math.sin(dLon/2)**2;
        return (R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a))).toFixed(1);
      })();
      const distLabel = distKm ? `<div style="font-size:11px;color:#2563EB;margin-bottom:4px;">📏 <strong>${distKm} km</strong> do terminal</div>` : '';
      // Horário estimado de chegada
      const horaCheg = absMinToHM(clocksPorParada[i]);
      const horaLabel = (horaCheg && !isOrigem && !tipoFim) ? `<div style="font-size:12px;font-weight:600;color:#111827;margin-bottom:4px;">🕐 Chegada estimada: <strong>${horaCheg}</strong></div>` : '';
      const descarga  = par?.tempoDescargaMin    ? `<div style="font-size:11px;color:#6B7280;margin-top:2px;">⏳ Descarga: ~${par.tempoDescargaMin} min</div>` : '';
      const janela    = par?.pedido?.restricao   ? `<div style="margin-top:6px;padding:4px 8px;background:#FFFBEB;color:#92400E;border-radius:4px;font-size:10px;border:1px solid #FDE68A;">⏱ Janela: ${par.pedido.restricao}</div>` : '';
      const centroide = p.isCentroide            ? `<div style="font-size:10px;color:#92400E;margin-top:4px;">⚠ Posição aproximada</div>` : '';
      const placa     = `<div style="font-size:11px;color:#6B7280;margin-bottom:6px;">🚛 ${item.v.placa} · ${item.v.transportadora || ''} · Viagem ${item.idx + 1}</div>`;
      let popupHtml = '';
      if (isOrigem) {
        popupHtml = `
          <div style="font-family:Inter,sans-serif;min-width:200px;">
            <div style="font-weight:700;font-size:13px;margin-bottom:4px;">🏭 ${p.nome}</div>
            ${placa}
          </div>`;
      } else if (tipoFim) {
        popupHtml = `
          <div style="font-family:Inter,sans-serif;min-width:200px;">
            <div style="font-weight:700;font-size:13px;margin-bottom:4px;">↩ Retorno — ${p.nome}</div>
            ${placa}
          </div>`;
      } else {
        popupHtml = `
          <div style="font-family:Inter,sans-serif;width:260px;">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E5E7EB;">
              <div style="width:24px;height:24px;border-radius:50%;background:${item.cor};color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;">${i}</div>
              <div>
                <div style="font-weight:700;font-size:13px;line-height:1.2;">${p.nome.replace(/^\d+\.\s*/, '')}</div>
                <div style="font-size:11px;color:#6B7280;">${p.cidade || ''}</div>
              </div>
            </div>
            ${placa}
            ${horaLabel}
            ${distLabel}
            ${produtos ? `<div style="background:#F9FAFB;border-radius:6px;padding:8px;margin-bottom:6px;font-size:11px;">
              <div style="font-weight:600;color:#6B7280;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px;">Carga entregue</div>
              ${produtos}
            </div>` : ''}
            ${descarga}${janela}${centroide}
          </div>`;
      }
      const marcador = L.marker([p.lat, p.lon], { icon }).addTo(camadaMapaGeral);
      marcador.bindPopup(popupHtml, { maxWidth: 300, autoPan: true });
      marcador.on('click', () => marcador.openPopup());
      bounds.push([p.lat, p.lon]);
    });
  });
  const placasTotais = porPlaca.size;
  const ciclosTotais = viagensVisiveis.length;
  const destinosTotais = viagensVisiveis.reduce((s, v) => s + v.destinos, 0);
  if (placaFiltroAtual) {
    top.innerHTML = `Placa ${placaFiltroAtual} · ${ciclosTotais} viagem(ns) · ${destinosTotais} destino(s)`
      + ` <button class="btn btn-sm" onclick="limparFiltroMapaPlaca()" style="margin-left:6px;">Limpar filtro</button>`;
  } else {
    top.textContent = `${placasTotais} placa(s) · ${ciclosTotais} viagem(ns) · ${destinosTotais} destino(s)${transpHeader ? ` · ${transpHeader}` : ''}`;
  }
  // Veículos alocados mas sem rota visível (coordenadas ausentes)
  const placasNoMapa = new Set(porPlaca.keys());
  const semCoordenadas = veiculosEscopo.filter(v => {
    const viagens = (ultimoResultado[v.id] || []).filter(vi => (vi.paradas || []).length);
    return viagens.length > 0 && !placasNoMapa.has(v.placa);
  });
  const itensLegenda = viagensMapa
    .sort((a, b) => a.v.placa.localeCompare(b.v.placa) || a.idx - b.idx)
    .map(item => `
      <div class="mapa-geral-item ${placaFiltroAtual === item.v.placa.toUpperCase() ? 'ativo' : ''}" onclick="toggleFiltroMapaPlaca('${item.v.placa.replace(/'/g, "\\'")}')">
        <span class="mapa-geral-cor" style="background:${item.cor};"></span>
        <div>
          <div style="font-weight:700;line-height:1.1;">${item.v.placa} <span style="font-size:10px;font-weight:400;color:#6B7280;">Viagem ${item.idx + 1}</span></div>
          <div style="font-size:11px;color:#4A6535;">${item.v.transportadora||'-'} · ${item.destinos} destino(s) · ${item.volViagem.toFixed(1)} m³</div>
        </div>
      </div>
    `).join('');
  const itensSemCoord = semCoordenadas.map(v => {
    const destinos = (ultimoResultado[v.id] || []).flatMap(vi => vi.paradas || []);
    const cidades  = [...new Set(destinos.map(p => p.pedido?.cidade).filter(Boolean))].join(', ');
    return `<div class="mapa-geral-item" style="opacity:0.65;cursor:default;" title="Coordenadas ausentes — cadastre lat/lon dos clientes para visualizar no mapa">
      <span class="mapa-geral-cor" style="background:#9CA3AF;"></span>
      <div>
        <div style="font-weight:700;line-height:1.1;">${v.placa} <span style="font-size:9px;color:#B91C1C;font-weight:600;">📍 sem coordenadas</span></div>
        <div style="font-size:11px;color:#6B7280;">${v.transportadora||'-'} · ${destinos.length} entrega(s) · ${cidades}</div>
      </div>
    </div>`;
  }).join('');
  legenda.innerHTML = (itensLegenda + itensSemCoord) || '<div class="empty" style="padding:16px;">Sem legenda.</div>';
  // Adia fitBounds para depois do browser reprocessar o layout do container
  const _b = bounds.slice();
  setTimeout(() => {
    mapaGeral.invalidateSize();
    if (_b.length > 1) mapaGeral.fitBounds(_b, { padding: [20, 20] });
  }, 80);
}
function toggleFiltroMapaPlaca(placa) {
  const alvo = (placa || '').toString().trim().toUpperCase();
  const atual = (filtroMapaPlaca || '').toString().trim().toUpperCase();
  filtroMapaPlaca = (atual === alvo) ? '' : alvo;
  renderMapaGeral();
}
function limparFiltroMapaPlaca() {
  filtroMapaPlaca = '';
  renderMapaGeral();
}
// ── Centroides de cidades (fallback quando pedido não tem lat/lon) ───────────
const CENTROIDES_CIDADE = {
  'sao paulo':[-23.5505,-46.6333],'campinas':[-22.9056,-47.0608],'santos':[-23.9618,-46.3322],
  'guaruja':[-23.9929,-46.2567],'sorocaba':[-23.5015,-47.4526],'jundiai':[-23.1864,-46.8842],
  'mogi das cruzes':[-23.5229,-46.1875],'sao bernardo do campo':[-23.6944,-46.5654],
  'sao caetano do sul':[-23.6228,-46.5531],'sao caetano':[-23.6228,-46.5531],
  'piracicaba':[-22.7253,-47.6492],'sao jose dos campos':[-23.1794,-45.8869],
  'ribeirao preto':[-21.1775,-47.8103],'bauru':[-22.3246,-49.0619],
  'sao jose do rio preto':[-20.8197,-49.3794],'osasco':[-23.5325,-46.7919],
  'santo andre':[-23.6639,-46.5383],'guarulhos':[-23.4553,-46.5333],
  'paulinia':[-22.7589,-47.1539],'cubatao':[-23.8978,-46.4253],
  'indaiatuba':[-23.0903,-47.2183],'limeira':[-22.5647,-47.4019],
  'americana':[-22.7394,-47.3331],'araras':[-22.3567,-47.3839],
  'botucatu':[-22.8860,-48.4447],'marilia':[-22.2139,-49.9458],
  'presidente prudente':[-22.1258,-51.3886],'catanduva':[-21.1378,-48.9775],
  'franca':[-20.5394,-47.4009],'araraquara':[-21.7942,-48.1758],
  'sao carlos':[-22.0174,-47.8908],'jau':[-22.2956,-48.5575],
  'lencois paulista':[-22.5981,-48.8006],'ourinhos':[-22.9781,-49.8703],
  'assis':[-22.6625,-50.4125],'itapetininga':[-23.5917,-48.0531],
  'itapeva':[-23.9819,-48.8756],'registro':[-24.4875,-47.8431],
  'caraguatatuba':[-23.6213,-45.4122],'sao sebastiao':[-23.7964,-45.4097],
  'ubatuba':[-23.4336,-45.0838],'atibaia':[-23.1172,-46.5506],
  'braganca paulista':[-22.9519,-46.5419],'mogi mirim':[-22.4319,-46.9575],
  'itapira':[-22.4361,-46.8228],'amparo':[-22.7044,-46.7661],
  'salto':[-23.2014,-47.2889],'itu':[-23.2642,-47.2997],
  'tatui':[-23.3556,-47.8572],'capao bonito':[-24.0039,-48.3506],
  'sao manuel':[-22.7317,-48.5706],'botucatu':[-22.8860,-48.4447],
  'jacarei':[-23.2981,-45.9658],'taubate':[-23.0256,-45.5553],
  'pindamonhangaba':[-22.9228,-45.4614],'aparecida':[-22.8481,-45.2314],
  'lorena':[-22.7328,-45.1236],'cacapava':[-23.0908,-45.7069],
  'sao jose dos campos':[-23.1794,-45.8869],'maua':[-23.6678,-46.4611],
  'diadema':[-23.6850,-46.6228],'sao caetano':[-23.6228,-46.5531],
  'riberiao pires':[-23.7108,-46.4156],'rio grande da serra':[-23.7439,-46.3961],
  'suzano':[-23.5428,-46.3108],'ferraz de vasconcelos':[-23.5408,-46.3686],
  'po':[-23.5311,-46.3406],'poá':[-23.5311,-46.3406],
  'itaquaquecetuba':[-23.4883,-46.3486],'aruja':[-23.3956,-46.3208],
  'santa isabel':[-23.3156,-46.2239],'biritiba mirim':[-23.5703,-46.0442],
  // Litoral SP
  'bertioga':[-23.8531,-46.1381],'praia grande':[-24.0056,-46.4122],
  'mongagua':[-24.0853,-46.6281],'itanhaem':[-24.1833,-46.7881],
  'peruibe':[-24.3194,-47.0000],'iguape':[-24.7081,-47.5578],
  'ilha comprida':[-24.7306,-47.5281],'cananeia':[-25.0147,-47.9294],
  'miracatu':[-24.2831,-47.4597],'pedro de toledo':[-24.2769,-47.2328],
  'juquia':[-24.3217,-47.6358],'sete barras':[-24.3856,-47.9239],
  'eldorado':[-24.5228,-48.1108],'iporanga':[-24.5939,-48.5958],
  'apiai':[-24.5083,-48.8428],'ribeira':[-24.6556,-49.0011],
  'ilha bela':[-23.7789,-45.3578],'sao sebastiao':[-23.7964,-45.4097],
  'salesopolis':[-23.5331,-45.8447],'santa branca':[-23.3978,-45.8808],
  'natividade da serra':[-23.3728,-45.4428],'cunha':[-23.0736,-44.9567],
  'paraibuna':[-23.3878,-45.6628],'redenção da serra':[-23.2636,-45.5428],
  'lagoinha':[-23.0881,-45.1561],'sao luis do paraitinga':[-23.2217,-45.3078],
  'guaratingueta':[-22.8164,-45.1925],'pindamonhangaba':[-22.9228,-45.4614],
  'potim':[-22.8244,-45.1836],'roseira':[-22.8944,-45.3058],
  'campos do jordao':[-22.7394,-45.5919],'santo antonio do pinhal':[-22.8231,-45.6628],
};
function _normCidade(cidade) {
  return (cidade || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s*\/\s*[a-z]{2,}$/i, '').replace(/\s+/g, ' ').trim();
}
function centroideCidade(cidade) {
  const c = CENTROIDES_CIDADE[_normCidade(cidade)];
  return c ? { lat: c[0], lon: c[1], isCentroide: true } : null;
}
// Retorna lat/lon efetivos do pedido: coordenadas próprias → cliente cadastrado → centróide da cidade
function latLonEfetivo(pedido) {
  const lat = parseFloat(pedido?.lat);
  const lon = parseFloat(pedido?.lon);
  if (!isNaN(lat) && !isNaN(lon) && (Math.abs(lat) > 0.001 || Math.abs(lon) > 0.001)) {
    return { lat, lon, isCentroide: false };
  }
  // Tenta coordenadas do cliente cadastrado
  const cli = encontrarClienteDoPedido(pedido);
  if (cli) {
    const clLat = parseFloat(cli.lat);
    const clLon = parseFloat(cli.lon);
    if (!isNaN(clLat) && !isNaN(clLon) && (Math.abs(clLat) > 0.001 || Math.abs(clLon) > 0.001)) {
      return { lat: clLat, lon: clLon, isCentroide: false };
    }
  }
  // Fallback: centróide da cidade
  const c = centroideCidade(pedido?.cidade);
  if (c) return c;
  return { lat: NaN, lon: NaN, isCentroide: false };
}
function haversine(lat1,lon1,lat2,lon2) {
  const R=6371, dLat=(lat2-lat1)*Math.PI/180, dLon=(lon2-lon1)*Math.PI/180;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function parseHoraMin(hora) {
  if (!hora || !hora.includes(':')) return NaN;
  const [h,m] = hora.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return NaN;
  return h*60 + m;
}
function duracaoJornadaMin(inicio, fim) {
  const ini = parseHoraMin(inicio);
  const end = parseHoraMin(fim);
  if (isNaN(ini) || isNaN(end)) return 0;
  if (end > ini) return end - ini;
  if (end === ini) return 24 * 60;
  return (24 * 60 - ini) + end;
}
function fmtHora(minAbs) {
  if (isNaN(minAbs)) return '--:--';
  const total = Math.round(minAbs);
  const dia = Math.floor(total / (24 * 60));
  const minDia = ((total % (24 * 60)) + (24 * 60)) % (24 * 60);
  const h = Math.floor(minDia / 60);
  const m = minDia % 60;
  const base = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  return base;
}
function parseJanelaRestricao(texto) {
  const s = (texto || '').toString();
  const matches = s.match(/\d{1,2}:\d{2}/g);
  if (!matches || matches.length < 2) return null;
  const ini = parseHoraMin(matches[0]);
  const fim = parseHoraMin(matches[1]);
  if (isNaN(ini) || isNaN(fim)) return null;
  return { inicioMin: ini, fimMin: fim };
}
function parseDataBr(texto) {
  const s = (texto || '').toString().trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return null;
}
function minutoNoDia(minAbs) {
  const d = 24 * 60;
  return ((Math.round(minAbs) % d) + d) % d;
}
function dentroJanela(minAbs, janela) {
  const m = minutoNoDia(minAbs);
  if (janela.fimMin >= janela.inicioMin) return m >= janela.inicioMin && m <= janela.fimMin;
  return m >= janela.inicioMin || m <= janela.fimMin;
}
// Recalcula timing de uma viagem no contexto do novo veículo (após drag-and-drop).
// Atualiza: esperaTerminalMin, tempoEsperaRestricaoMin de cada parada, tempoConsumidoMin.
// Recalcula timing de todas as viagens de um veículo em sequência
function recalcularTodasViagens(...veiculoIds) {
  veiculoIds.forEach(id => {
    const v = veiculos.find(x => x.id === id);
    if (!v) return;
    (ultimoResultado[id] || []).filter(vi => !vi._vazio && vi.paradas?.length).forEach(vi => {
      recalcularTimingViagem(vi, v);
    });
  });
  recalcularControleTempo();
}
// Ressincroniza controleTempo.usadoMin de todos os veículos a partir dos tempos reais
// das viagens. Deve ser chamada após qualquer edição manual (drag-and-drop, troca de placa,
// override de horário), pois recalcularTimingViagem atualiza as viagens mas não o contador.
function recalcularControleTempo() {
  if (!ultimoControleTempo || !ultimoResultado) return;
  veiculos.forEach(v => {
    const ct = ultimoControleTempo[String(v.id)];
    if (!ct) return;
    const viagensReais = (ultimoResultado[v.id] || [])
      .filter(vi => !vi._vazio && vi.paradas?.length);
    const produtivo = viagensReais.reduce((soma, vi) =>
      soma + (vi.paradas || []).reduce((s, p) =>
        s + (p.tempoCarregamentoMin || 0)
          + (p.deslocCarregadoMin   || 0)
          + (p.tempoDescargaMin     || 0)
          + (p.deslocVazioMin       || 0), 0), 0);
    // Pausa de refeição ocorre uma vez (entre 1ª e 2ª viagem) e conta como jornada
    const pausaRefeicao = viagensReais.length >= (doisTurnos(v) ? 3 : 2) ? (v.tempoPerdidoMin || 0) : 0;
    ct.usadoMin = produtivo + pausaRefeicao;
  });
}
function recalcularTimingViagem(viagem, v) {
  if (!viagem || viagem._vazio || !(viagem.paradas || []).length) return;
  const viagens = (ultimoResultado[v.id] || []).filter(vi => !vi._vazio && vi.paradas?.length);
  const idx     = viagens.indexOf(viagem);
  const _jIniRaw4 = parseHoraMin(v.jornadaInicio || '06:00');
  let jIniMin = isNaN(_jIniRaw4) ? 360 : _jIniRaw4;
  if (v._horarioDisponivelAPartirDe) { const _dm4 = parseHoraMin(v._horarioDisponivelAPartirDe); if (!isNaN(_dm4) && _dm4 > jIniMin) jIniMin = _dm4; }

  let clock;
  if (viagem.horarioCargaManualMin !== undefined && !isNaN(viagem.horarioCargaManualMin)) {
    // Override manual: o clock desta viagem começa EXATAMENTE no horário escolhido,
    // independente do término da viagem anterior. O usuário assume a responsabilidade.
    // Descobre em qual dia absoluto colocar o horário:
    // usa o clock da viagem anterior como referência de dia, mas não de sequência.
    const clockAnterior = idx > 0
      ? inicioViagemAbsMin(viagens, idx, jIniMin, v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1)
      : jIniMin;
    const baseDay = Math.floor(clockAnterior / 1440) * 1440;
    let alvo = baseDay + viagem.horarioCargaManualMin;
    // Se o horário manual é antes do início calculado, avança para o próximo dia
    if (alvo < clockAnterior - 0.001) alvo += 1440;
    clock = alvo;
    viagem.esperaTerminalMin = 0; // espera já está embutida no clock manual
  } else {
    // Sem override: comportamento normal — clock baseado na sequência de viagens
    clock = inicioViagemAbsMin(viagens, Math.max(idx, 0), jIniMin, v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1);
    viagem.esperaTerminalMin = calcEsperaTerminal(viagem.terminalOrigem, clock);
    clock += viagem.esperaTerminalMin;
  }
  // Salva o horário de início de carga — é este clock neste momento,
  // sem nenhum atraso adicional. É o mesmo valor que aparece no campo "Carga" da tela.
  viagem._inicioCargaMin = clock;
  // Nota: esperaTerminalMin já foi somado dentro do else acima.
  // Não somar novamente aqui para evitar dupla contagem.
  // Recalcula cada parada
  let produtivo = 0;
  viagem.paradas.forEach((p, idxP) => {
    // Corrige origemDeslocamento de acordo com a posição real na viagem
    p.origemDeslocamento = idxP === 0 ? 'Terminal' : 'Entrega anterior';
    if (idxP === 0) {
      // Primeira parada sem tempoCarregamento (stop vindo de posição não-primeira)
      // → recalcula distâncias a partir do terminal atual
      if (!(p.tempoCarregamentoMin > 0)) {
        const base = dadosCiclo(v, p.pedido, viagem.terminalOrigem);
        p.tempoCarregamentoMin = base.tempoCarregamentoMin;
        p.deslocCarregadoMin   = base.deslocCarregadoMin;
        p.deslocVazioMin       = base.deslocVazioMin;
      }
    } else {
      // Paradas não-primeiras nunca têm carregamento no terminal
      p.tempoCarregamentoMin = 0;
    }
    // Carga: apenas na 1ª parada
    const carga = idxP === 0 ? (p.tempoCarregamentoMin || 0) : 0;
    clock   += carga;
    produtivo += carga;
    // Deslocamento até o cliente
    clock   += (p.deslocCarregadoMin || 0);
    produtivo += (p.deslocCarregadoMin || 0);
    // Recalcula espera por restrição de horário do cliente
    if (p.pedido?.restricao && !p.overnight) {
      const pts = (p.pedido.restricao.match(/\d{1,2}:\d{2}/g) || []).map(parseHoraMin);
      if (pts.length >= 1 && !isNaN(pts[0])) {
        const chegadaDia = ((clock % 1440) + 1440) % 1440;
        p.tempoEsperaRestricaoMin = chegadaDia < pts[0] ? pts[0] - chegadaDia : 0;
      } else {
        p.tempoEsperaRestricaoMin = 0;
      }
    } else {
      p.tempoEsperaRestricaoMin = 0;
    }
    clock     += (p.tempoEsperaRestricaoMin || 0) + (p.tempoDescargaMin || 0) + (p.deslocVazioMin || 0);
    produtivo += (p.tempoDescargaMin || 0) + (p.deslocVazioMin || 0);
  });
  viagem.tempoConsumidoMin = produtivo;
}
// Veículos com jornada > 14h operam 2 turnos: a pausa de refeição só entra
// a partir da 3ª viagem (troca de turno entre 1ª e 2ª não conta como pausa).
// Veículos com jornada ≤ 14h (turno único): pausa entra entre 1ª e 2ª viagem.
function doisTurnos(v) {
  const dur = v?.jornadaMin || duracaoJornadaMin(v?.jornadaInicio || '06:00', v?.jornadaFim || '18:00');
  return dur > 840;
}
// numMaxBreaks: quantas pausas de refeição são somadas ao relógio
//   1 (padrão, turno único): pausa apenas antes da 2ª viagem
//   2 (2 turnos/São Caetano): pausa antes da 2ª E antes da 3ª viagem
// A pausa de refeição NÃO é armazenada em tempoConsumidoMin das viagens;
// é adicionada aqui para que o relógio absoluto fique correto.
function inicioViagemAbsMin(viagensVeiculo, idxViagem, jornadaInicioMin, tempoPerdidoMin = 0, numMaxBreaks = 1) {
  let t = jornadaInicioMin;
  for (let i = 0; i < idxViagem; i++) {
    const vi = viagensVeiculo[i];
    // Se a viagem seguinte tem horário manual fixo, usa o horário manual diretamente
    // como início, descartando a cadeia de tempoConsumidoMin que viria antes.
    const proxVi = viagensVeiculo[i + 1];
    if (proxVi?.horarioCargaManualMin !== undefined && i + 1 === idxViagem) {
      // Retorna o horário manual absoluto (ajustado para o dia correto)
      const baseDay = Math.floor((t + (vi?.tempoConsumidoMin || 0)) / 1440) * 1440;
      let alvo = baseDay + proxVi.horarioCargaManualMin;
      if (alvo < t - 0.001) alvo += 1440;
      return alvo;
    }
    t += vi?.tempoConsumidoMin || 0;
  }
  t += Math.min(idxViagem, numMaxBreaks) * tempoPerdidoMin;
  return t;
}
// Retorna minutos de espera no terminal até abertura. 0 se já aberto ou sem restrição.
function calcEsperaTerminal(terminalNome, clockAbsMin) {
  const tc = terminaisCad.find(t => t.nome === terminalNome);
  if (!tc) return 0;
  const abMin = parseHoraMin(tc.aberturaPadrao || '00:00');
  if (!abMin || abMin <= 0) return 0;
  const clockNoDia = ((clockAbsMin % (24 * 60)) + 24 * 60) % (24 * 60);
  return clockNoDia >= abMin ? 0 : abMin - clockNoDia;
}
// Calcula timings corretos para entrega overnight:
// carrega antes do terminal fechar, aguarda na base, segue viagem para chegar na janela.
function calcOvernightViagem(v, item, clockAbsMin, terminalNome, diaAlvoEntrega=null) {
  const tc = terminaisCad.find(t => t.nome === terminalNome);
  const janela = parseJanelaRestricao(item.pedido?.restricao);
  if (!janela || !tc) return null;
  const feMin = parseHoraMin(tc.fechamentoPadrao || '');
  if (!feMin || feMin <= 0) return null;
  const abMin = parseHoraMin(tc.aberturaPadrao || '00:00') || 0;
  const base = dadosCiclo(v, item.pedido, terminalNome);
  const { tempoCarregamentoMin, deslocCarregadoMin, tempoDescargaMin, deslocVazioMin } = base;
  // Início do dia atual em minutos absolutos
  const clockDia = ((clockAbsMin % 1440) + 1440) % 1440;
  const dayBase  = clockAbsMin - clockDia;
  // Carrega o mais tarde possível antes do terminal fechar (hoje ou amanhã)
  let loadStartAbs = dayBase + feMin - tempoCarregamentoMin;
  if (loadStartAbs < clockAbsMin) loadStartAbs += 1440; // já passou hoje → amanhã
  // Não pode iniciar antes de abrir
  const loadDayBase = loadStartAbs - ((loadStartAbs % 1440 + 1440) % 1440);
  if ((loadStartAbs % 1440) < abMin) loadStartAbs = loadDayBase + abMin;
  const loadEndAbs = loadStartAbs + tempoCarregamentoMin;
  const minArrivalAbs = loadEndAbs + deslocCarregadoMin;
  let targetArrivalAbs = null;
  if (Number.isInteger(diaAlvoEntrega)) {
    const startAbs = diaAlvoEntrega * 1440 + janela.inicioMin;
    const endAbs = janela.fimMin >= janela.inicioMin
      ? diaAlvoEntrega * 1440 + janela.fimMin
      : (diaAlvoEntrega + 1) * 1440 + janela.fimMin;
    targetArrivalAbs = Math.max(startAbs, minArrivalAbs);
    if (targetArrivalAbs > endAbs) return null;
  } else {
    // Próxima ocorrência da janela do cliente após o carregamento
    const loadEndDia  = ((loadEndAbs % 1440) + 1440) % 1440;
    const loadDayBase2 = loadEndAbs - loadEndDia;
    targetArrivalAbs = loadDayBase2 + janela.inicioMin;
    if (targetArrivalAbs < minArrivalAbs) targetArrivalAbs += 1440;
  }
  const targetDepartureAbs  = targetArrivalAbs - deslocCarregadoMin;
  const waitAfterLoadingMin = Math.max(0, targetDepartureAbs - loadEndAbs);
  const waitBeforeLoadingMin = Math.max(0, loadStartAbs - clockAbsMin);
  const productiveMin    = tempoCarregamentoMin + deslocCarregadoMin + tempoDescargaMin + deslocVazioMin;
  const totalElapsedMin  = waitBeforeLoadingMin + tempoCarregamentoMin + waitAfterLoadingMin + deslocCarregadoMin + tempoDescargaMin + deslocVazioMin;
  if (waitAfterLoadingMin < 0 || totalElapsedMin <= 0) return null;
  return { ...base, waitBeforeLoadingMin, waitAfterLoadingMin, productiveMin, totalElapsedMin };
}
function chegadaPrevistaAbsMin(viagem, detalheParada, viagensVeiculo, idxViagem, jornadaInicioMin, tempoPerdidoMin = 0, numMaxBreaks = 1) {
  let t = inicioViagemAbsMin(viagensVeiculo, idxViagem, jornadaInicioMin, tempoPerdidoMin, numMaxBreaks);
  t += viagem.esperaTerminalMin || 0; // espera até abertura do terminal
  (viagem.paradas || []).forEach(p => {
    // waitAfterLoadingMin = espera overnight na base; tempoEsperaRestricaoMin = espera no cliente por janela
    t += (p.tempoCarregamentoMin || 0) + (p.waitAfterLoadingMin || 0)
       + (p.deslocCarregadoMin || 0) + (p.tempoEsperaRestricaoMin || 0)
       + (p.tempoDescargaMin || 0) + (p.deslocVazioMin || 0);
  });
  t += (detalheParada?.tempoCarregamentoMin || 0) + (detalheParada?.deslocCarregadoMin || 0);
  return t;
}
function avaliarRestricaoPedido(pedido, chegadaAbsMin, diaAlvoEntrega=null) {
  const janela = parseJanelaRestricao(pedido?.restricao);
  if (!janela) return { ok: true, esperaMin: 0 };
  if (Number.isInteger(diaAlvoEntrega)) {
    const startAbs = diaAlvoEntrega * 1440 + janela.inicioMin;
    const endAbs = janela.fimMin >= janela.inicioMin
      ? diaAlvoEntrega * 1440 + janela.fimMin
      : (diaAlvoEntrega + 1) * 1440 + janela.fimMin;
    if (chegadaAbsMin > endAbs) return { ok: false, esperaMin: 0 };
    if (chegadaAbsMin < startAbs) return { ok: true, esperaMin: startAbs - chegadaAbsMin };
    return { ok: true, esperaMin: 0 };
  }
  const m = minutoNoDia(chegadaAbsMin);
  if (janela.fimMin >= janela.inicioMin) {
    if (m > janela.fimMin) return { ok: false, esperaMin: 0 };
    if (m < janela.inicioMin) return { ok: true, esperaMin: janela.inicioMin - m };
    return { ok: true, esperaMin: 0 };
  }
  if (m >= janela.inicioMin || m <= janela.fimMin) return { ok: true, esperaMin: 0 };
  return { ok: true, esperaMin: (janela.inicioMin - m + 24 * 60) % (24 * 60) };
}
function encontrarClienteDoPedido(pedido) {
  if (!pedido) return null;
  if (pedido.codigoSAP) {
    const cSap = clientes.find(c => (c.codigoSAP || '').trim() === String(pedido.codigoSAP).trim());
    if (cSap) return cSap;
  }
  return clientes.find(c => (c.nome || '').trim().toLowerCase() === (pedido.cliente || '').trim().toLowerCase()) || null;
}
// Fator de tortuosidade: converte distância ponto-a-ponto (haversine) em
// distância rodoviária estimada. Ajuste aqui se a calibração precisar mudar.
const FATOR_DISTANCIA = 1.3;
function tempoDeslocamentoMin(distKm, velKmh) {
  return velKmh > 0 ? (distKm * FATOR_DISTANCIA / velKmh) * 60 : 0;
}
function dadosCiclo(v, pedido, terminalOrigemNome=null) {
  const terminalNome = terminalOrigemNome || v.terminal || pedido?.terminal || '';
  const terminal = terminaisCad.find(t => t.nome === terminalNome);
  const cliente = encontrarClienteDoPedido(pedido);
  const tempoCarregamentoMin = terminal?.tempoCarregamentoMedioMin || 60;
  const tempoDescargaMin = cliente?.tempoDescargaMediaMin || 45;
  const velCarregado = v.velMediaCarregado || 45;
  const velVazio = v.velMediaVazio || 55;
  const _coordP = latLonEfetivo(pedido);
  const latPedido = _coordP.lat;
  const lonPedido = _coordP.lon;
  const distanciaKm = (terminal && !isNaN(latPedido) && !isNaN(lonPedido))
    ? haversine(terminal.lat, terminal.lon, latPedido, lonPedido)
    : 0;
  const deslocCarregadoMin = tempoDeslocamentoMin(distanciaKm, velCarregado);
  const deslocVazioMin = tempoDeslocamentoMin(distanciaKm, velVazio);
  const cicloMin = tempoCarregamentoMin + deslocCarregadoMin + tempoDescargaMin + deslocVazioMin;
  return {
    terminal,
    cliente,
    distanciaKm,
    tempoCarregamentoMin,
    tempoDescargaMin,
    deslocCarregadoMin,
    deslocVazioMin,
    cicloMin,
  };
}
function dadosIncrementoParada(viagem, v, pedido, terminalOrigemNome=null) {
  const terminalNome = terminalOrigemNome || viagem?.terminalOrigem || v.terminal || pedido?.terminal || '';
  const base = dadosCiclo(v, pedido, terminalNome);
  const velCarregado = v.velMediaCarregado || 45;
  const velVazio = v.velMediaVazio || 55;
  if (!viagem.paradas.length) {
    return {
      incrementoMin: base.cicloMin,
      distanciaKm: base.distanciaKm,
      tempoCarregamentoMin: base.tempoCarregamentoMin,
      deslocCarregadoMin: base.deslocCarregadoMin,
      tempoDescargaMin: base.tempoDescargaMin,
      deslocVazioMin: base.deslocVazioMin,
      origemDeslocamento: 'Terminal',
      removeRetornoAnterior: false,
    };
  }
  const ultima = viagem.paradas[viagem.paradas.length - 1];
  const terminal = terminaisCad.find(t => t.nome === terminalNome);
  const _coordUlt  = latLonEfetivo(ultima?.pedido);
  const _coordNova = latLonEfetivo(pedido);
  const latUlt  = _coordUlt.lat;  const lonUlt  = _coordUlt.lon;
  const latNova = _coordNova.lat; const lonNova = _coordNova.lon;
  const distUltNova = (!isNaN(latUlt) && !isNaN(lonUlt) && !isNaN(latNova) && !isNaN(lonNova))
    ? haversine(latUlt, lonUlt, latNova, lonNova)
    : 0;
  const distNovaTerm = (terminal && !isNaN(latNova) && !isNaN(lonNova))
    ? haversine(latNova, lonNova, terminal.lat, terminal.lon)
    : 0;
  const retornoAnteriorMin = ultima.deslocVazioMin || 0;
  const deslocEntreMin = tempoDeslocamentoMin(distUltNova, velCarregado);
  const retornoNovoMin = tempoDeslocamentoMin(distNovaTerm, velVazio);
  const incrementoMin = Math.max(0, -retornoAnteriorMin + deslocEntreMin + base.tempoDescargaMin + retornoNovoMin);
  return {
    incrementoMin,
    distanciaKm: distUltNova,
    tempoCarregamentoMin: 0,
    deslocCarregadoMin: deslocEntreMin,
    tempoDescargaMin: base.tempoDescargaMin,
    deslocVazioMin: retornoNovoMin,
    origemDeslocamento: 'Entrega anterior',
    removeRetornoAnterior: true,
  };
}
// ── Best-insertion helpers ────────────────────────────────────────────────────
// Retorna todas as N+1 posições de inserção possíveis para um novo pedido em uma
// viagem existente, ordenadas pelo menor custo geográfico incremental (menor desvio
// de rota). Posição 0 = antes da 1ª parada; posição N = append ao final (atual).
function posicoesCandidatas(viagem, v, pedido) {
  const paradas = viagem.paradas;
  const n = paradas.length;
  if (!n) return [{ idx: 0, custo: 0 }];
  const terminal = terminaisCad.find(t => t.nome === viagem.terminalOrigem);
  const velC = v.velMediaCarregado || 45;
  const velV = v.velMediaVazio    || 55;
  const cNova = latLonEfetivo(pedido);
  if (isNaN(cNova.lat) || isNaN(cNova.lon)) return [{ idx: n, custo: 0 }];
  const posicoes = [];
  for (let i = 0; i <= n; i++) {
    const prevC = i === 0
      ? { lat: terminal?.lat, lon: terminal?.lon }
      : latLonEfetivo(paradas[i - 1].pedido);
    const nextC = i === n
      ? { lat: terminal?.lat, lon: terminal?.lon }
      : latLonEfetivo(paradas[i].pedido);
    if (isNaN(prevC?.lat) || isNaN(nextC?.lat)) { posicoes.push({ idx: i, custo: Infinity }); continue; }
    const velNext = i === n ? velV : velC;
    const custo =
      tempoDeslocamentoMin(haversine(prevC.lat, prevC.lon, cNova.lat, cNova.lon), velC)
    + tempoDeslocamentoMin(haversine(cNova.lat, cNova.lon, nextC.lat, nextC.lon), velNext)
    - tempoDeslocamentoMin(haversine(prevC.lat, prevC.lon, nextC.lat, nextC.lon), velNext);
    posicoes.push({ idx: i, custo });
  }
  return posicoes.sort((a, b) => a.custo - b.custo);
}
// Calcula os detalhes de inserção para uma posição do meio da viagem
// (não ao final — para o final usa dadosIncrementoParada existente).
function dadosInsercaoEmPosicao(viagem, v, pedido, terminalNome, insertIdx) {
  const terminal = terminaisCad.find(t => t.nome === terminalNome);
  const velC = v.velMediaCarregado || 45;
  const base = dadosCiclo(v, pedido, terminalNome);
  const paradas = viagem.paradas;
  const cNova = latLonEfetivo(pedido);
  const prevC = insertIdx === 0
    ? { lat: terminal?.lat, lon: terminal?.lon }
    : latLonEfetivo(paradas[insertIdx - 1].pedido);
  const nextC = latLonEfetivo(paradas[insertIdx].pedido);
  const dPrevNova  = tempoDeslocamentoMin(haversine(prevC.lat, prevC.lon, cNova.lat, cNova.lon), velC);
  const dNovaNext  = tempoDeslocamentoMin(haversine(cNova.lat, cNova.lon, nextC.lat, nextC.lon), velC);
  const dPrevNext  = tempoDeslocamentoMin(haversine(prevC.lat, prevC.lon, nextC.lat, nextC.lon), velC);
  return {
    insertIdx,
    incrementoMin:        Math.max(0, dPrevNova + base.tempoDescargaMin + dNovaNext - dPrevNext),
    distanciaKm:          haversine(prevC.lat, prevC.lon, cNova.lat, cNova.lon),
    tempoCarregamentoMin: insertIdx === 0 ? base.tempoCarregamentoMin : 0,
    deslocCarregadoMin:   dPrevNova,
    tempoDescargaMin:     base.tempoDescargaMin,
    deslocVazioMin:       0, // parada do meio nunca retorna ao terminal
    origemDeslocamento:   insertIdx === 0 ? 'Terminal' : 'Entrega anterior',
    removeRetornoAnterior: false,
    _deslocNovaNextMin:   dNovaNext, // para atualizar deslocCarregadoMin da próxima parada
  };
}
// Versão de chegadaPrevistaAbsMin para inserção em posição intermediária:
// soma apenas as paradas ANTES de insertIdx.
function chegadaInsercaoAbsMin(viagem, det, insertIdx, viagensVeiculo, idxViagem, jornadaInicioMin, tempoPerdidoMin, numMaxBreaks = 1) {
  let t = inicioViagemAbsMin(viagensVeiculo, idxViagem, jornadaInicioMin, tempoPerdidoMin, numMaxBreaks);
  t += viagem.esperaTerminalMin || 0;
  (viagem.paradas || []).slice(0, insertIdx).forEach(p => {
    t += (p.tempoCarregamentoMin || 0) + (p.waitAfterLoadingMin || 0)
       + (p.deslocCarregadoMin   || 0) + (p.tempoEsperaRestricaoMin || 0)
       + (p.tempoDescargaMin     || 0) + (p.deslocVazioMin          || 0);
  });
  t += (det.tempoCarregamentoMin || 0) + (det.deslocCarregadoMin || 0);
  return t;
}
function contarCandidatosItem(item) {
  return veiculos
    .filter(v => !item.tiposCaminhao.length || item.tiposCaminhao.includes(v.tipo))
    .filter(v => !item.pedido.identidadePetronas || !!v.identidadePetronas)
    .filter(v => veiculoAtendeTerminal(v, item.terminal)).length;
}
function optsTerminais(valorSel) {
  return '<option value="">Selecione o terminal...</option>' +
    nomeTerminais().map(n => `<option value="${n}" ${n===valorSel?'selected':''}>${n}</option>`).join('');
}
function optsClientes(valorSel) {
  return '<option value="">— selecionar cliente —</option>' +
    clientes.map(c => `<option value="${c.id}" ${c.id===valorSel?'selected':''}>${c.codigoSAP ? '['+c.codigoSAP+'] ' : ''}${c.nome}</option>`).join('');
}
// ── Combobox pesquisável ──────────────────────────────────────────────────────
function _norm(s) { return (s||'').toLowerCase().normalize('NFD').replace(/\p{Mn}/gu,''); }
function initCombobox(sel) {
  if (!sel || sel._cmb) return;
  sel._cmb = true;
  sel.style.display = 'none';
  const wrap = document.createElement('div'); wrap.className = 'cmb-wrap';
  sel.parentNode.insertBefore(wrap, sel);
  wrap.appendChild(sel);
  const inp = document.createElement('input');
  inp.type = 'text'; inp.autocomplete = 'off'; inp.className = 'cmb-input';
  inp.placeholder = sel.dataset.ph || 'Digite para buscar...';
  const arr = document.createElement('span'); arr.className = 'cmb-arrow'; arr.textContent = '▾';
  const list = document.createElement('div'); list.className = 'cmb-list';
  wrap.prepend(arr); wrap.prepend(inp); wrap.appendChild(list);
  const allOpts = () => Array.from(sel.options).filter(o => o.value !== '');
  function show(q) {
    const nq = _norm(q);
    const matches = allOpts().filter(o => !nq || _norm(o.text).includes(nq));
    list.innerHTML = matches.length
      ? matches.map(o => `<div class="cmb-item" data-v="${o.value.replace(/"/g,'&quot;')}">${o.text}</div>`).join('')
      : '<div class="cmb-empty">Nenhum resultado</div>';
    list.style.display = 'block';
  }
  function hide() { list.style.display = 'none'; }
  function pick(val, txt) { sel.value = val; inp.value = txt; hide(); sel.dispatchEvent(new Event('change',{bubbles:true})); }
  function sync() { const o = allOpts().find(o=>o.selected); inp.value = o ? o.text : ''; }
  inp.addEventListener('focus', () => show(inp.value));
  inp.addEventListener('input', () => show(inp.value));
  inp.addEventListener('keydown', e => {
    const items = [...list.querySelectorAll('.cmb-item')];
    let idx = items.findIndex(i=>i.classList.contains('cmb-on'));
    if (e.key==='ArrowDown') { e.preventDefault(); idx=Math.min(idx+1,items.length-1); }
    else if (e.key==='ArrowUp') { e.preventDefault(); idx=Math.max(idx-1,0); }
    else if (e.key==='Enter') { e.preventDefault(); if(items[idx]) pick(items[idx].dataset.v, items[idx].textContent); return; }
    else if (e.key==='Escape') { hide(); return; }
    items.forEach(i=>i.classList.remove('cmb-on'));
    if (idx>=0) { items[idx].classList.add('cmb-on'); items[idx].scrollIntoView({block:'nearest'}); }
  });
  list.addEventListener('mousedown', e => { e.preventDefault(); const it=e.target.closest('.cmb-item'); if(it) pick(it.dataset.v, it.textContent); });
  document.addEventListener('mousedown', e => { if(!wrap.contains(e.target)) hide(); });
  sel._cmbSync = sync;
  sync();
}
function refreshCombobox(sel) {
  if (!sel) return;
  if (!sel._cmb) { initCombobox(sel); } else if (sel._cmbSync) { sel._cmbSync(); }
}
function renderTipoBtns(containerId, selecionados=[]) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = TIPOS_CAMINHAO.map(t =>
    `<button type="button" class="tipo-btn ${selecionados.includes(t)?'ativo':''}" onclick="toggleTipo(this)">${t}</button>`
  ).join('');
}
function toggleTipo(btn) { btn.classList.toggle('ativo'); }
function lerTiposSelecionados(containerId) {
  return [...document.querySelectorAll(`#${containerId} .tipo-btn.ativo`)].map(b => b.textContent);
}
// ═══════════════════════════════════════════════════════════════════════════
// TERMINAIS
// ═══════════════════════════════════════════════════════════════════════════
function gerarHorarioGrid(diasAtivos, horarios) {
  const grid = document.getElementById('horario-grid');
  if (!grid) return;
  grid.innerHTML = DIAS_NOMES.map((d,i) => {
    const ativo = diasAtivos.includes(i);
    const ab = horarios?.[i]?.abertura   || '';
    const fe = horarios?.[i]?.fechamento || '';
    return `<div class="horario-col">
      <div class="horario-col-label">${d}</div>
      <input type="time" id="h-ab-${i}" value="${ab}" ${!ativo?'disabled':''} placeholder="ab"/>
      <input type="time" id="h-fe-${i}" value="${fe}" ${!ativo?'disabled':''} placeholder="fe"/>
    </div>`;
  }).join('');
}
function toggleDia(btn) {
  btn.classList.toggle('ativo');
  const i = parseInt(btn.dataset.dia);
  const ab = document.getElementById(`h-ab-${i}`);
  const fe = document.getElementById(`h-fe-${i}`);
  const on = btn.classList.contains('ativo');
  if (ab) ab.disabled = !on;
  if (fe) fe.disabled = !on;
}
function aplicarPadrao(forcar=false) {
  const ab = document.getElementById('t-abertura-padrao').value;
  const fe = document.getElementById('t-fechamento-padrao').value;
  document.querySelectorAll('#dias-semana-btns .dia-btn').forEach(btn => {
    const i = parseInt(btn.dataset.dia), on = btn.classList.contains('ativo');
    const hAb = document.getElementById(`h-ab-${i}`), hFe = document.getElementById(`h-fe-${i}`);
    if (hAb && hFe && on) {
      if (forcar || !hAb.value) hAb.value = ab;
      if (forcar || !hFe.value) hFe.value = fe;
    }
  });
}
function abrirFormTerminal(id=null) {
  editandoTerminalId = id;
  document.querySelectorAll('#dias-semana-btns .dia-btn').forEach(b => b.classList.remove('ativo'));
  let diasAtivos = [0,1,2,3,4,5], horarios = null;
  if (id !== null) {
    const t = terminaisCad.find(x => x.id === id);
    if (t) {
      document.getElementById('form-terminal-title').textContent = 'Editar terminal';
      document.getElementById('t-nome').value   = t.nome;
      document.getElementById('t-cidade').value = t.cidade;
      document.getElementById('t-distribuidora').value = t.distribuidora || '';
      document.getElementById('t-empresa-local').value = t.empresaLocalExpedicao || '';
      document.getElementById('t-lat').value    = t.lat;
      document.getElementById('t-lon').value    = t.lon;
      document.getElementById('t-fuso').value   = t.fuso;
      document.getElementById('t-carregamento').value = t.tempoCarregamentoMedioMin || 60;
      document.getElementById('t-abertura-padrao').value   = t.aberturaPadrao  || '06:00';
      document.getElementById('t-fechamento-padrao').value = t.fechamentoPadrao || '18:00';
      diasAtivos = t.diasAtivos; horarios = t.horarios;
    }
  } else {
    document.getElementById('form-terminal-title').textContent = 'Novo terminal';
    ['t-nome','t-cidade','t-distribuidora','t-empresa-local','t-lat','t-lon'].forEach(fid => { const e=document.getElementById(fid); if(e) e.value=''; });
    document.getElementById('t-fuso').value   = '-3';
    document.getElementById('t-carregamento').value = 60;
    document.getElementById('t-abertura-padrao').value   = '06:00';
    document.getElementById('t-fechamento-padrao').value = '18:00';
    diasAtivos = [0,1,2,3,4,5];
  }
  document.querySelectorAll('#dias-semana-btns .dia-btn').forEach(btn => {
    if (diasAtivos.includes(parseInt(btn.dataset.dia))) btn.classList.add('ativo');
  });
  gerarHorarioGrid(diasAtivos, horarios);
  document.getElementById('form-terminal').classList.remove('hidden');
  document.getElementById('form-terminal').scrollIntoView({behavior:'smooth',block:'start'});
}
function cancelarFormTerminal() {
  editandoTerminalId = null;
  document.getElementById('form-terminal').classList.add('hidden');
}
function lerDiasAtivos() {
  return [...document.querySelectorAll('#dias-semana-btns .dia-btn.ativo')].map(b => parseInt(b.dataset.dia));
}
function lerHorarios(diasAtivos) {
  const h = {};
  diasAtivos.forEach(i => {
    h[i] = {
      abertura:   document.getElementById(`h-ab-${i}`)?.value ?? '',
      fechamento: document.getElementById(`h-fe-${i}`)?.value ?? '',
    };
  });
  return h;
}
function salvarTerminal() {
  const nome  = document.getElementById('t-nome').value.trim();
  const cidade= document.getElementById('t-cidade').value.trim();
  const distribuidora = document.getElementById('t-distribuidora').value.trim();
  const empresaLocalExpedicao = document.getElementById('t-empresa-local').value.trim();
  const lat   = parseFloat(document.getElementById('t-lat').value);
  const lon   = parseFloat(document.getElementById('t-lon').value);
  const fuso  = document.getElementById('t-fuso').value;
  const tempoCarregamentoMedioMin = parseFloat(document.getElementById('t-carregamento').value);
  const abPad = document.getElementById('t-abertura-padrao').value;
  const fePad = document.getElementById('t-fechamento-padrao').value;
  if (!nome) { alert('Informe o nome do terminal.'); return; }
  if (isNaN(lat) || isNaN(lon)) { alert('Informe latitude e longitude válidas.'); return; }
  if (isNaN(tempoCarregamentoMedioMin) || tempoCarregamentoMedioMin <= 0) { alert('Informe o tempo médio de carregamento do terminal.'); return; }
  const diasAtivos = lerDiasAtivos();
  if (!diasAtivos.length) { alert('Selecione ao menos um dia de funcionamento.'); return; }
  const horarios = lerHorarios(diasAtivos);
  if (editandoTerminalId !== null) {
    const idx = terminaisCad.findIndex(t => t.id === editandoTerminalId);
    if (idx !== -1) terminaisCad[idx] = {...terminaisCad[idx], nome, cidade, distribuidora, empresaLocalExpedicao, lat, lon, fuso, tempoCarregamentoMedioMin, aberturaPadrao:abPad, fechamentoPadrao:fePad, diasAtivos, horarios};
  } else {
    if (terminaisCad.find(t => t.nome.toLowerCase() === nome.toLowerCase())) { alert('Já existe um terminal com esse nome.'); return; }
    terminaisCad.push({id:Date.now(), nome, cidade, distribuidora, empresaLocalExpedicao, lat, lon, fuso, tempoCarregamentoMedioMin, aberturaPadrao:abPad, fechamentoPadrao:fePad, diasAtivos, horarios});
  }
  editandoTerminalId = null;
  cancelarFormTerminal();
  renderTerminais();
  atualizarDropdownsTerminais();
}
function removerTerminal(id) {
  if (!confirm('Remover este terminal? Pedidos e veículos vinculados perderão a referência.')) return;
  terminaisCad = terminaisCad.filter(t => t.id !== id);
  renderTerminais();
  atualizarDropdownsTerminais();
}
function renderTerminais() {
  const el = document.getElementById('terminais-list');
  if (!terminaisCad.length) { el.innerHTML = '<div class="empty">Nenhum terminal cadastrado.<br>Clique em "+ Terminal" para começar.</div>'; return; }
  const filtroTerm = valId('f-term-terminal');
  const filtroCidade = valId('f-term-cidade');
  const lista = terminaisCad.filter(t =>
    containsFiltro(t.nome, filtroTerm) &&
    containsFiltro(t.cidade, filtroCidade)
  );
  if (!lista.length) { el.innerHTML = '<div class="empty">Nenhum terminal encontrado para os filtros.</div>'; return; }
  el.innerHTML = lista.map(t => {
    const horTexto = t.diasAtivos.map(i => {
      const h  = t.horarios?.[i];
      const ab = h?.abertura   || t.aberturaPadrao  || '—';
      const fe = h?.fechamento || t.fechamentoPadrao || '—';
      return `${DIAS_NOMES[i]}: ${ab}–${fe}`;
    }).join(' &nbsp;|&nbsp; ');
    const diasTags = DIAS_NOMES.map((d,i) => `<span class="dia-tag ${t.diasAtivos.includes(i)?'on':'off'}">${d}</span>`).join('');
    return `<div class="terminal-card">
      <div class="terminal-card-header">
        <div style="flex:1;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:4px;">
            <span style="font-family:var(--font-cond);font-weight:700;font-size:15px;letter-spacing:0.04em;">${t.nome}</span>
            ${t.cidade?`<span style="font-size:12px;color:#4A6535;">${t.cidade}</span>`:''}
            ${t.distribuidora ? `<span class="tag tag-blue" style="font-size:9px;">CIA: ${t.distribuidora}</span>` : ''}
            ${t.empresaLocalExpedicao ? `<span class="tag tag-purple" style="font-size:9px;">Exp: ${t.empresaLocalExpedicao}</span>` : ''}
            <span class="tag tag-gray" style="font-size:9px;">UTC${t.fuso>=0?'+':''}${t.fuso}</span>
            <span class="tag tag-lime" style="font-size:9px;">Carga média: ${(t.tempoCarregamentoMedioMin||60).toFixed(0)} min</span>
          </div>
          <div style="font-size:11px;color:#4A6535;margin-bottom:6px;">📍 ${t.lat.toFixed(4)}, ${t.lon.toFixed(4)}</div>
          <div class="terminal-dias">${diasTags}</div>
          <div style="font-size:10px;color:#4A6535;margin-top:6px;line-height:1.6;">${horTexto}</div>
        </div>
        <div style="display:flex;gap:5px;flex-shrink:0;">
          <button class="btn btn-sm" onclick="abrirFormTerminal(${t.id})">Editar</button>
          <button class="btn btn-sm btn-danger" onclick="removerTerminal(${t.id})">✕</button>
        </div>
      </div>
    </div>`;
  }).join('');
}
function atualizarDropdownsTerminais() {
  const vt = document.getElementById('v-terminal');
  if (vt) { const cur = vt.value; vt.innerHTML = optsTerminais(cur); }
  const pt = document.getElementById('p-terminal');
  if (pt) { const cur = pt.value; pt.innerHTML = optsTerminais(cur); }
}
function carregarExemplosTerminais() {
  terminaisCad = [
    { id:1, nome:'Terminal Paulínia',  cidade:'Paulínia / SP',  distribuidora:'PETRONAS', lat:-22.7589, lon:-47.1539, fuso:'-3', tempoCarregamentoMedioMin:60, aberturaPadrao:'06:00', fechamentoPadrao:'22:00', diasAtivos:[1,2,3,4,5], horarios:{1:{abertura:'06:00',fechamento:'22:00'},2:{abertura:'06:00',fechamento:'22:00'},3:{abertura:'06:00',fechamento:'22:00'},4:{abertura:'06:00',fechamento:'22:00'},5:{abertura:'06:00',fechamento:'22:00'}} },
    { id:2, nome:'Terminal Santos',    cidade:'Santos / SP',    distribuidora:'PETRONAS', lat:-23.9618, lon:-46.3322, fuso:'-3', tempoCarregamentoMedioMin:75, aberturaPadrao:'07:00', fechamentoPadrao:'19:00', diasAtivos:[1,2,3,4,5,6], horarios:{1:{abertura:'07:00',fechamento:'19:00'},2:{abertura:'07:00',fechamento:'19:00'},3:{abertura:'07:00',fechamento:'19:00'},4:{abertura:'07:00',fechamento:'19:00'},5:{abertura:'07:00',fechamento:'19:00'},6:{abertura:'08:00',fechamento:'14:00'}} },
    { id:3, nome:'Terminal Guarulhos', cidade:'Guarulhos / SP', distribuidora:'PETRONAS', lat:-23.4553, lon:-46.5333, fuso:'-3', tempoCarregamentoMedioMin:45, aberturaPadrao:'00:00', fechamentoPadrao:'23:59', diasAtivos:[0,1,2,3,4,5,6], horarios:{} }
  ];
  renderTerminais();
  atualizarDropdownsTerminais();
}
// ═══════════════════════════════════════════════════════════════════════════
// CLIENTES
// ═══════════════════════════════════════════════════════════════════════════
function abrirFormCliente(id) {
  editandoClienteId = id;
  if (id !== null) {
    const c = clientes.find(x => x.id === id);
    if (!c) return;
    document.getElementById('form-cliente-title').textContent = 'Editar cliente';
    document.getElementById('cl-sap').value      = c.codigoSAP || '';
    document.getElementById('cl-nome').value     = c.nome;
    document.getElementById('cl-cidade').value   = c.cidade;
    document.getElementById('cl-lat').value      = c.lat;
    document.getElementById('cl-lon').value      = c.lon;
    document.getElementById('cl-descarga').value = c.tempoDescargaMediaMin || 45;
    document.getElementById('cl-restricao').value= c.restricaoHorario || '';
    document.getElementById('cl-obs').value      = c.observacoes || '';
    document.getElementById('cl-identidade-petronas').checked = !!c.identidadePetronas;
    renderTipoBtns('cl-tipos-btns', c.tiposCaminhao || []);
  } else {
    document.getElementById('form-cliente-title').textContent = 'Novo cliente';
    ['cl-sap','cl-nome','cl-cidade','cl-lat','cl-lon','cl-restricao','cl-obs'].forEach(fid => {
      const e = document.getElementById(fid); if (e) e.value = '';
    });
    document.getElementById('cl-descarga').value = 45;
    document.getElementById('cl-identidade-petronas').checked = false;
    if (document.getElementById('cl-tipos-btns')) renderTipoBtns('cl-tipos-btns', []);
  }
  document.getElementById('form-cliente').classList.remove('hidden');
  document.getElementById('form-cliente').scrollIntoView({behavior:'smooth', block:'start'});
}
function cancelarFormCliente() {
  editandoClienteId = null;
  document.getElementById('form-cliente').classList.add('hidden');
}
function salvarCliente() {
  const nome   = document.getElementById('cl-nome').value.trim();
  const lat    = parseFloat(document.getElementById('cl-lat').value);
  const lon    = parseFloat(document.getElementById('cl-lon').value);
  const tempoDescargaMediaMin = parseFloat(document.getElementById('cl-descarga').value);
  if (!nome) { alert('Informe o nome do cliente.'); return; }
  if (isNaN(lat) || isNaN(lon)) { alert('Informe latitude e longitude válidas.'); return; }
  if (isNaN(tempoDescargaMediaMin) || tempoDescargaMediaMin <= 0) { alert('Informe o tempo médio de descarga do cliente.'); return; }
  const dados = {
    codigoSAP:       document.getElementById('cl-sap').value.trim(),
    nome,
    cidade:          document.getElementById('cl-cidade').value.trim(),
    lat, lon,
    tempoDescargaMediaMin,
    restricaoHorario: document.getElementById('cl-restricao').value.trim(),
    observacoes:     document.getElementById('cl-obs').value.trim(),
    tiposCaminhao:   lerTiposSelecionados('cl-tipos-btns'),
    identidadePetronas: document.getElementById('cl-identidade-petronas').checked,
  };
  if (editandoClienteId !== null) {
    const idx = clientes.findIndex(c => c.id === editandoClienteId);
    if (idx !== -1) clientes[idx] = {...clientes[idx], ...dados};
  } else {
    clientes.push({id: Date.now(), ...dados});
  }
  editandoClienteId = null;
  cancelarFormCliente();
  renderClientes();
  atualizarDropdownsClientes();
}
function removerCliente(id) {
  if (!confirm('Remover este cliente?')) return;
  clientes = clientes.filter(c => c.id !== id);
  renderClientes();
  atualizarDropdownsClientes();
}
function renderClientes() {
  const el = document.getElementById('clientes-list');
  if (!clientes.length) {
    el.innerHTML = '<div class="empty">Nenhum cliente cadastrado.<br>Clique em "+ Cliente" para começar.</div>';
    return;
  }
  const filtroCliente = valId('f-cli-cliente');
  const filtroCidade = valId('f-cli-cidade');
  const lista = clientes.filter(c =>
    containsFiltro(c.nome, filtroCliente) &&
    containsFiltro(c.cidade, filtroCidade)
  );
  if (!lista.length) {
    el.innerHTML = '<div class="empty">Nenhum cliente encontrado para os filtros.</div>';
    return;
  }
  el.innerHTML = lista.map(c => {
    const tipos = c.tiposCaminhao && c.tiposCaminhao.length
      ? c.tiposCaminhao.map(t => `<span class="tag tag-blue" style="font-size:9px;">${t}</span>`).join(' ')
      : '<span style="font-size:11px;color:#5E9A18;">Todos os tipos</span>';
    return `<div class="card">
      <div style="display:flex;align-items:flex-start;gap:10px;">
        <div style="flex:1;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:5px;">
            ${c.codigoSAP ? `<span class="sap-code">SAP ${c.codigoSAP}</span>` : ''}
            <span style="font-family:var(--font-cond);font-weight:700;font-size:14px;letter-spacing:0.03em;">${c.nome}</span>
            ${c.cidade ? `<span style="font-size:12px;color:#4A6535;">${c.cidade}</span>` : ''}
            ${c.restricaoHorario ? `<span class="tag tag-yellow">${c.restricaoHorario}</span>` : ''}
            ${c.identidadePetronas ? `<span class="tag tag-yellow" style="font-size:9px;">⬡ ID Petronas</span>` : ''}
            <span class="tag tag-lime" style="font-size:9px;">Descarga média: ${(c.tempoDescargaMediaMin||45).toFixed(0)} min</span>
          </div>
          <div style="font-size:11px;color:#4A6535;margin-bottom:6px;">📍 ${c.lat.toFixed(4)}, ${c.lon.toFixed(4)}</div>
          <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:4px;">
            <span style="font-size:10px;color:#4A6535;font-weight:600;text-transform:uppercase;letter-spacing:0.08em;">Caminhões:</span>
            ${tipos}
          </div>
          ${c.observacoes ? `<div style="font-size:11px;color:#4A6535;margin-top:4px;font-style:italic;">${c.observacoes}</div>` : ''}
        </div>
        <div style="display:flex;gap:5px;flex-shrink:0;">
          <button class="btn btn-sm" onclick="abrirFormCliente(${c.id})">Editar</button>
          <button class="btn btn-sm btn-danger" onclick="removerCliente(${c.id})">✕</button>
        </div>
      </div>
    </div>`;
  }).join('');
}
function atualizarDropdownsClientes() {
  const sel = document.getElementById('p-importar-cliente');
  if (sel) { const cur = parseInt(sel.value); sel.innerHTML = optsClientes(cur); refreshCombobox(sel); }
}
function preencherPedidoDeCliente(clienteId) {
  const id = parseInt(clienteId);
  if (!id) return;
  const c = clientes.find(x => x.id === id);
  if (!c) return;
  document.getElementById('p-sap').value      = c.codigoSAP || '';
  document.getElementById('p-cliente').value  = c.nome;
  document.getElementById('p-cidade').value   = c.cidade;
  document.getElementById('p-lat').value      = c.lat;
  document.getElementById('p-lon').value      = c.lon;
  document.getElementById('p-restricao').value = c.restricaoHorario || '';
  document.getElementById('p-identidade-petronas').checked = !!c.identidadePetronas;
  renderTipoBtns('p-tipos-btns', c.tiposCaminhao || []);
}
function carregarExemplosClientes() {
  clientes = [
    { id:101, codigoSAP:'1000123', nome:'Posto Ipiranga Centro',   cidade:'Campinas / SP', lat:-22.9056, lon:-47.0608, tempoDescargaMediaMin:50, restricaoHorario:'06:00–10:00', tiposCaminhao:['Truck','Bitruck'],        observacoes:'Acesso pela Av. Principal. Portaria exige agendamento.', identidadePetronas:true },
    { id:102, codigoSAP:'1000456', nome:'Posto Shell Viracopos',   cidade:'Campinas / SP', lat:-23.0074, lon:-47.1344, tempoDescargaMediaMin:40, restricaoHorario:'',            tiposCaminhao:[],                           observacoes:'', identidadePetronas:false },
    { id:103, codigoSAP:'1000789', nome:'Distribuidora Santos',    cidade:'Santos / SP',   lat:-23.9618, lon:-46.3322, tempoDescargaMediaMin:70, restricaoHorario:'08:00–17:00', tiposCaminhao:['Bitruck','Cavalo Mecânico'], observacoes:'Balança obrigatória na entrada.', identidadePetronas:true },
    { id:104, codigoSAP:'1001012', nome:'Posto Ale Litoral',       cidade:'Guarujá / SP',  lat:-23.9929, lon:-46.2567, tempoDescargaMediaMin:45, restricaoHorario:'',            tiposCaminhao:['Truck'],                    observacoes:'Acesso pela estrada da praia. Caminhões grandes não passam.', identidadePetronas:false },
    { id:105, codigoSAP:'1001345', nome:'Frota Transportes MG',    cidade:'Sorocaba / SP', lat:-23.5015, lon:-47.4526, tempoDescargaMediaMin:35, restricaoHorario:'',            tiposCaminhao:[],                           observacoes:'', identidadePetronas:false },
    { id:106, codigoSAP:'1001678', nome:'Posto BR Bandeirantes',   cidade:'Jundiaí / SP',  lat:-23.1864, lon:-46.8842, tempoDescargaMediaMin:60, restricaoHorario:'07:00–11:00', tiposCaminhao:['Bitruck','Cavalo Mecânico'], observacoes:'Descarregar por compartimento, não simultaneamente.', identidadePetronas:true },
    { id:107, codigoSAP:'1002001', nome:'Rede TeraCom Sorocaba',   cidade:'Sorocaba / SP', lat:-23.5300, lon:-47.4600, tempoDescargaMediaMin:55, restricaoHorario:'',            tiposCaminhao:[],                           observacoes:'Recebe Arla 32 e Diesel no mesmo horário.', identidadePetronas:false },
  ];
  renderClientes();
  atualizarDropdownsClientes();
}
// ═══════════════════════════════════════════════════════════════════════════
// PEDIDOS
// ═══════════════════════════════════════════════════════════════════════════
function addProdutoForm(produto='', volume='', ordemSAP='') {
  numProdForm++;
  const id = numProdForm;
  const div = document.createElement('div');
  div.className = 'prod-item';
  div.id = 'prod-row-' + id;
  div.innerHTML = `
    <div class="row" style="margin-bottom:0;align-items:flex-end;">
      <div class="field"><label>Produto</label>
        <select id="pf-prod-${id}" data-ph="Buscar produto..."><option value="">— selecione o produto —</option>${PRODUTOS.map(p=>`<option value="${p}" ${p===produto?'selected':''}>${p}</option>`).join('')}</select>
      </div>
      <div class="field" style="max-width:130px;"><label>Volume (m³)</label>
        <input type="number" step="0.5" id="pf-vol-${id}" value="${volume}" placeholder="25"/>
      </div>
      <div class="field" style="max-width:140px;"><label>No. Ordem SAP</label>
        <input type="text" id="pf-ordem-${id}" value="${ordemSAP}" placeholder="Ex: 1234567890"/>
      </div>
      ${id>1?`<button class="btn btn-sm btn-danger" style="flex:0;margin-bottom:2px;" onclick="document.getElementById('prod-row-${id}').remove()">✕</button>`:''}
    </div>`;
  document.getElementById('produtos-form').appendChild(div);
  initCombobox(document.getElementById('pf-prod-' + id));
}
function abrirFormPedido(id) {
  editandoPedidoId = id;
  numProdForm = 0;
  document.getElementById('produtos-form').innerHTML = '';
  document.getElementById('p-importar-cliente').innerHTML = optsClientes('');
  if (id !== null) {
    const p = pedidos.find(x => x.id === id);
    if (!p) return;
    document.getElementById('form-pedido-title').textContent = 'Editar pedido';
    document.getElementById('p-sap').value       = p.codigoSAP || '';
    document.getElementById('p-cliente').value   = p.cliente;
    document.getElementById('p-cidade').value    = p.cidade;
    document.getElementById('p-lat').value       = p.lat;
    document.getElementById('p-lon').value       = p.lon;
    document.getElementById('p-restricao').value = p.restricao || '';
    // dataEntregaLogistica vem como DD/MM/YYYY → converte para YYYY-MM-DD para o input date
    const _dlEdit = p.dataEntregaLogistica || '';
    const _dlPts = _dlEdit.split('/');
    document.getElementById('p-data-entrega').value = _dlPts.length === 3
      ? `${_dlPts[2]}-${_dlPts[1].padStart(2,'0')}-${_dlPts[0].padStart(2,'0')}` : '';
    document.getElementById('p-identidade-petronas').checked = !!p.identidadePetronas;
    document.getElementById('p-terminal').innerHTML = optsTerminais(p.terminal || '');
    renderTipoBtns('p-tipos-btns', p.tiposCaminhao || []);
    p.produtos.forEach(pr => addProdutoForm(pr.produto, pr.volume, pr.ordemSAP || ''));
    if (!p.produtos.length) addProdutoForm();
  } else {
    document.getElementById('form-pedido-title').textContent = 'Novo pedido';
    ['p-sap','p-cliente','p-cidade','p-lat','p-lon','p-restricao','p-data-entrega'].forEach(fid => { const e=document.getElementById(fid); if(e) e.value=''; });
    document.getElementById('p-identidade-petronas').checked = false;
    document.getElementById('p-terminal').innerHTML = optsTerminais('');
    if (document.getElementById('p-tipos-btns')) renderTipoBtns('p-tipos-btns', []);
    addProdutoForm();
  }
  refreshCombobox(document.getElementById('p-importar-cliente'));
  refreshCombobox(document.getElementById('p-terminal'));
  document.getElementById('form-pedido').classList.remove('hidden');
  document.getElementById('form-pedido').scrollIntoView({behavior:'smooth',block:'start'});
}
function cancelarFormPedido() {
  editandoPedidoId = null;
  document.getElementById('form-pedido').classList.add('hidden');
}
function salvarPedido() {
  const terminal = document.getElementById('p-terminal').value;
  const cliente  = document.getElementById('p-cliente').value.trim();
  const cidade   = document.getElementById('p-cidade').value.trim();
  const dataEnt  = document.getElementById('p-data-entrega').value;
  const erros = [];
  if (!cliente)  erros.push('• Cliente');
  if (!cidade)   erros.push('• Cidade');
  if (!terminal) erros.push('• Terminal de carga');
  if (!dataEnt)  erros.push('• Data de entrega');
  const produtos = [];
  document.querySelectorAll('[id^="prod-row-"]').forEach(row => {
    const rid  = row.id.replace('prod-row-','');
    const prod = document.getElementById('pf-prod-'+rid);
    const vol  = document.getElementById('pf-vol-'+rid);
    const ord  = document.getElementById('pf-ordem-'+rid);
    if (prod && vol && prod.value && parseFloat(vol.value) > 0)
      produtos.push({ produto: prod.value, volume: parseFloat(vol.value), ordemSAP: ord?.value.trim() || '' });
  });
  if (!produtos.length) erros.push('• Ao menos um produto com quantidade');
  if (erros.length) { alert('Preencha os campos obrigatórios:\n\n' + erros.join('\n')); return; }
  const ordens = [...new Set(produtos.map(pr => pr.ordemSAP).filter(Boolean))];
  const dados = {
    codigoSAP: document.getElementById('p-sap').value.trim(),
    ordens,
    cliente,
    cidade,
    lat:       parseFloat(document.getElementById('p-lat').value) || 0,
    lon:       parseFloat(document.getElementById('p-lon').value) || 0,
    terminal,
    restricao:     document.getElementById('p-restricao').value || null,
    tiposCaminhao: lerTiposSelecionados('p-tipos-btns'),
    identidadePetronas: document.getElementById('p-identidade-petronas').checked,
    dataEntregaLogistica: (() => {
      const v = document.getElementById('p-data-entrega').value; // YYYY-MM-DD
      if (!v) return '';
      const [y, m, d] = v.split('-');
      return `${d}/${m}/${y}`;
    })(),
    produtos,
  };
  if (editandoPedidoId !== null) {
    const idx = pedidos.findIndex(p => p.id === editandoPedidoId);
    if (idx !== -1) pedidos[idx] = {...pedidos[idx], ...dados};
  } else {
    pedidos.push({id: Date.now(), ...dados});
  }
  editandoPedidoId = null;
  cancelarFormPedido();
  renderPedidos();
}
function removerPedido(id) { pedidos = pedidos.filter(p => p.id !== id); renderPedidos(); }
function totalVolPedido(p) { return p.produtos.reduce((s,pr) => s+pr.volume, 0); }
function renderPedidos() {
  const el = document.getElementById('pedidos-list');
  if (!pedidos.length) { el.innerHTML = ''; return; }
  const filtroCliente = valId('f-ped-cliente');
  const filtroCidade = valId('f-ped-cidade');
  const filtroTerminal = valId('f-ped-terminal');
  let lista = pedidos.filter(p =>
    containsFiltro(p.cliente, filtroCliente) &&
    containsFiltro(p.cidade, filtroCidade) &&
    containsFiltro(p.terminal, filtroTerminal)
  );
  if (!lista.length) { el.innerHTML = '<div class="empty">Nenhum pedido encontrado para os filtros.</div>'; return; }
  // Mantém grupos de quebra juntos, ordenados por parte
  lista = lista.slice().sort((a, b) => {
    const gA = a._quebraGrupo || String(a.id);
    const gB = b._quebraGrupo || String(b.id);
    if (gA !== gB) return gA < gB ? -1 : 1;
    return (a._quebraParte || 1) - (b._quebraParte || 1);
  });
  el.innerHTML = lista.map(p => {
    const tiposHtml = p.tiposCaminhao && p.tiposCaminhao.length
      ? p.tiposCaminhao.map(t => `<span class="tag tag-blue" style="font-size:9px;">${t}</span>`).join(' ')
      : '';
    const temQuebra    = !!p._quebraGrupo;
    const parteTag     = p._quebraParte
      ? `<span class="tag tag-yellow" style="font-size:9px;">⌀ Carga ${p._quebraParte}</span>` : '';
    const borderStyle  = temQuebra ? 'border-left:4px solid var(--pet-yellow);' : '';
    const podeQuebrar  = p.produtos.length >= 2;
    return `<div class="card" style="${borderStyle}">
      <div style="display:flex;align-items:flex-start;gap:10px;">
        <div style="flex:1;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:6px;">
            ${p.codigoSAP ? `<span class="sap-code">SAP ${p.codigoSAP}</span>` : ''}
            ${parteTag}
            <span style="font-weight:700;font-size:14px;font-family:var(--font-cond);">${p.cliente}</span>
            <span style="font-size:12px;color:#4A6535;">${p.cidade}</span>
            ${p.terminal ? termTag(p.terminal) : ''}
            ${p.dataEntregaLogistica ? `<span class="tag tag-blue" style="font-size:9px;">📅 ${p.dataEntregaLogistica}</span>` : ''}
            ${p.restricao ? `<span class="tag tag-yellow">${p.restricao}</span>` : ''}
            ${p.identidadePetronas ? `<span class="tag tag-yellow" style="font-size:9px;">⬡ ID Petronas</span>` : ''}
            ${tiposHtml}
          </div>
          <div class="pills">
            ${p.produtos.map(pr => `<span class="pill">${pr.produto}: ${pr.volume} m³${pr.ordemSAP ? ` · OS ${pr.ordemSAP}` : ''}</span>`).join('')}
            <span class="pill" style="color:var(--pet-green);border-color:#C4E87A;">Total: ${totalVolPedido(p).toFixed(1)} m³</span>
          </div>
        </div>
        <div style="display:flex;gap:5px;flex-shrink:0;flex-wrap:wrap;justify-content:flex-end;">
          ${podeQuebrar ? `<button class="btn btn-sm" title="Dividir em duas cargas" onclick="abrirQuebraPedido(${p.id})">⌀ Quebrar</button>` : ''}
          <button class="btn btn-sm" onclick="abrirFormPedido(${p.id})">Editar</button>
          <button class="btn btn-sm btn-danger" onclick="removerPedido(${p.id})">✕</button>
        </div>
      </div>
    </div>`;
  }).join('');
}
// ─── Quebrar Pedido ────────────────────────────────────────────────────────────
var _quebraIdAtual = null;
function abrirQuebraPedido(id) {
  const p = pedidos.find(x => x.id === id);
  if (!p || p.produtos.length < 2) return;
  _quebraIdAtual = id;
  document.getElementById('quebra-pedido-info').textContent =
    p.cliente + ' · ' + p.cidade + ' · ' + totalVolPedido(p).toFixed(1) + ' m³ total';
  document.getElementById('quebra-preview').style.display = 'none';
  document.getElementById('btn-confirmar-quebra').disabled = true;
  document.getElementById('quebra-produtos-list').innerHTML = p.produtos.map((pr, i) =>
    `<div style="border:1px solid var(--border);border-radius:var(--radius-sm);margin-bottom:6px;
                 background:var(--bg);transition:background 0.12s;" id="qrow-${i}">
      <label style="display:flex;align-items:center;gap:10px;padding:9px 11px;cursor:pointer;">
        <input type="checkbox" value="${i}" id="qchk-${i}"
               style="accent-color:var(--pet-green);width:15px;height:15px;cursor:pointer;flex-shrink:0;"
               onchange="onQuebraCheck(${i},${pr.volume})"/>
        <div style="flex:1;">
          <div style="font-size:12px;font-weight:600;">${pr.produto}</div>
          <div style="font-size:11px;color:var(--text-2);">${pr.volume} m³ total</div>
        </div>
      </label>
      <div id="qvol-wrap-${i}" style="display:none;padding:0 11px 10px 36px;display:none;">
        <div style="display:flex;align-items:center;gap:8px;font-size:12px;">
          <span style="color:var(--text-2);">Volume na Carga 2:</span>
          <input type="number" id="qvol-${i}" value="${pr.volume}"
                 min="0.1" max="${pr.volume}" step="0.5"
                 style="width:80px;padding:4px 7px;border:1.5px solid var(--border);border-radius:5px;
                        font-size:12px;font-family:var(--font);text-align:right;"
                 onfocus="this.style.borderColor='var(--pet-green)'"
                 onblur="this.style.borderColor='var(--border)'"
                 oninput="atualizarPreviewQuebra()"/>
          <span style="color:var(--text-2);">m³</span>
          <span id="qvol-resto-${i}" style="color:var(--text-3);font-size:11px;"></span>
        </div>
      </div>
    </div>`
  ).join('');
  document.getElementById('modal-quebra-pedido').classList.add('show');
}
function onQuebraCheck(i, volMax) {
  const wrap = document.getElementById('qvol-wrap-' + i);
  const chk  = document.getElementById('qchk-' + i);
  const row  = document.getElementById('qrow-' + i);
  wrap.style.display = chk.checked ? 'block' : 'none';
  row.style.background = chk.checked ? 'var(--pet-green-bg)' : 'var(--bg)';
  const inp = document.getElementById('qvol-' + i);
  if (chk.checked && inp) inp.value = volMax;
  atualizarPreviewQuebra();
}
function atualizarPreviewQuebra() {
  const p = pedidos.find(x => x.id === _quebraIdAtual);
  if (!p) return;
  const btn  = document.getElementById('btn-confirmar-quebra');
  const prev = document.getElementById('quebra-preview');
  // Lê o estado de cada produto
  const linhas = p.produtos.map((pr, i) => {
    const chk    = document.getElementById('qchk-' + i);
    const volInp = document.getElementById('qvol-' + i);
    const checked = chk?.checked || false;
    const volPt2  = checked ? Math.min(Math.max(parseFloat(volInp?.value) || pr.volume, 0.1), pr.volume) : 0;
    const volPt1  = pr.volume - volPt2;
    const resto   = document.getElementById('qvol-resto-' + i);
    if (resto && checked) resto.textContent = volPt1 > 0.001 ? `(${volPt1.toFixed(1)} m³ ficam na Carga 1)` : '(tudo vai para Carga 2)';
    return { pr, checked, volPt1, volPt2 };
  });
  const pt1 = linhas.filter(l => l.volPt1 > 0.001).map(l => ({ produto: l.pr.produto, volume: l.volPt1 }));
  const pt2 = linhas.filter(l => l.checked && l.volPt2 > 0.001).map(l => ({ produto: l.pr.produto, volume: l.volPt2 }));
  const temSel = pt2.length > 0;
  const temRest = pt1.length > 0;
  if (!temSel || !temRest) { btn.disabled = true; prev.style.display = 'none'; return; }
  const fmt = arr => arr.map(pr => {
    const nome = pr.produto.split(' - ')[0];
    return `${nome}: ${pr.volume.toFixed(1)} m³`;
  }).join(', ');
  const v1 = pt1.reduce((s, pr) => s + pr.volume, 0);
  const v2 = pt2.reduce((s, pr) => s + pr.volume, 0);
  prev.style.display = 'block';
  prev.innerHTML = `<strong>Carga 1 (${v1.toFixed(1)} m³):</strong> ${fmt(pt1)}<br>
                    <strong>Carga 2 (${v2.toFixed(1)} m³):</strong> ${fmt(pt2)}`;
  btn.disabled = false;
}
function confirmarQuebraPedido() {
  const p = pedidos.find(x => x.id === _quebraIdAtual);
  if (!p) return;
  const linhas = p.produtos.map((pr, i) => {
    const chk    = document.getElementById('qchk-' + i);
    const volInp = document.getElementById('qvol-' + i);
    const checked = chk?.checked || false;
    const volPt2  = checked ? Math.min(Math.max(parseFloat(volInp?.value) || pr.volume, 0.1), pr.volume) : 0;
    const volPt1  = pr.volume - volPt2;
    return { pr, checked, volPt1, volPt2 };
  });
  const prodsPt1 = linhas
    .filter(l => l.volPt1 > 0.001)
    .map(l => ({ ...l.pr, volume: l.volPt1 }));
  const prodsPt2 = linhas
    .filter(l => l.checked && l.volPt2 > 0.001)
    .map(l => ({ ...l.pr, volume: l.volPt2 }));
  if (!prodsPt1.length || !prodsPt2.length) return;
  const grupo  = p._quebraGrupo || (String(p.id) + '_grp');
  const parte1 = p._quebraParte || 1;
  p.produtos     = prodsPt1;
  p._quebraGrupo = grupo;
  p._quebraParte = parte1;
  pedidos.push({
    id:                   Date.now() + Math.floor(Math.random() * 999),
    codigoSAP:            p.codigoSAP,
    ordens:               [...(p.ordens || [])],
    cliente:              p.cliente,
    cidade:               p.cidade,
    lat:                  p.lat,
    lon:                  p.lon,
    terminal:             p.terminal,
    restricao:            p.restricao,
    tiposCaminhao:        [...(p.tiposCaminhao || [])],
    identidadePetronas:   p.identidadePetronas,
    dataEntregaLogistica: p.dataEntregaLogistica,
    produtos:             prodsPt2,
    _quebraGrupo:         grupo,
    _quebraParte:         parte1 + 1,
  });
  fecharQuebraPedido();
  renderPedidos();
}
function fecharQuebraPedido(event) {
  if (event && event.target !== document.getElementById('modal-quebra-pedido')) return;
  document.getElementById('modal-quebra-pedido').classList.remove('show');
  _quebraIdAtual = null;
}
// ─── Pedidos Liberados (Excel) ───────────────────────────────────────────────
function xlsxMapPedidosLiberadosRows(rows) {
  const groups = {};
  const order  = [];
  for (const r of rows) {
    const erpId    = String(r['Clientes Id ERP']      ?? r['Cliente Id ERP']      ?? '').trim();
    const terminal = String(r['Terminal']              ?? '').trim();
    const cliente  = String(r['Clientes Razão Social'] ?? r['Cliente Razão Social'] ?? '').trim();
    if (!cliente) continue;
    const cidade   = String(r['Cidade Entrega']        ?? r['Cidade']               ?? '').trim();
    const material = String(r['Material']              ?? '').trim();
    const prodNome = String(r['Produto']               ?? '').trim();
    const volume   = parseFloat(r['Volume Pedido (m³)'] ?? r['Volume Pedido'] ?? r['Volume'] ?? 0);
    const ordemSAP = String(r['No. Ordem SAP'] ?? r['No Ordem SAP'] ?? r['Ordem SAP'] ?? '').trim();
    const dataEntregaRaw = r['Data Entrega Logística'] ?? r['Data Entrega Logistica'] ?? r['Data Entrega'] ?? '';
    const dataEntrega = dataEntregaRaw instanceof Date
      ? dataEntregaRaw.toLocaleDateString('pt-BR')
      : String(dataEntregaRaw).trim().replace(/^(\d{4})-(\d{2})-(\d{2}).*/, '$3/$2/$1');
    const key = erpId + '||' + terminal;
    if (!groups[key]) {
      const cliCad = clientes.find(c => c.codigoSAP === erpId);
      groups[key] = {
        codigoSAP: erpId,
        ordens: [],       // Nos. Ordem SAP acumulados por linha do Excel
        cliente,
        cidade,
        lat: cliCad?.lat ?? 0,
        lon: cliCad?.lon ?? 0,
        terminal,
        restricao: cliCad?.restricaoHorario ?? null,
        tiposCaminhao: cliCad?.tiposCaminhao ?? [],
        identidadePetronas: cliCad?.identidadePetronas ?? false,
        dataEntregaLogistica: dataEntrega,
        produtos: [],
      };
      order.push(key);
    }
    if (ordemSAP && !groups[key].ordens.includes(ordemSAP)) groups[key].ordens.push(ordemSAP);
    if (volume > 0) {
      const prodLabel = PRODUTOS.find(p => p.startsWith(material))
        ?? (material ? `${material} - ${prodNome}` : prodNome);
      // ordemSAP vai junto ao produto para ser exibido por compartimento
      if (prodLabel) groups[key].produtos.push({ produto: prodLabel, volume, ordemSAP });
    }
  }
  return order.map((key, i) => {
    const g = groups[key];
    g.id = Date.now() + i;
    return g;
  }).filter(g => g.produtos.length > 0);
}
async function carregarPedidosLiberados() {
  if (typeof XLSX === 'undefined') return false;
  const candidatos = [
    'pedidos_liberados.xlsx',
    'Pedidos Liberados.xlsx',
    'Pedidos_Liberados.xlsx',
    'pedidos liberados.xlsx',
  ];
  for (const arquivo of candidatos) {
    try {
      const resp = await fetch(arquivo + '?t=' + Date.now());
      if (!resp.ok) continue;
      const wb   = XLSX.read(await resp.arrayBuffer(), { type: 'array', cellDates: true });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
      if (!rows.length) continue;
      const novos = xlsxMapPedidosLiberadosRows(rows);
      if (!novos.length) continue;
      pedidos = novos;
      renderPedidos();
      return true;
    } catch (e) {
      console.warn(`[${arquivo}]`, e.message);
    }
  }
  return false;
}
// ─── Upload de pedidos pelo usuário ──────────────────────────────────────────
function uploadPedidosLiberados(input) {
  if (!input.files || !input.files[0]) return;
  if (typeof XLSX === 'undefined') { alert('SheetJS não carregado.'); return; }
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
      if (!rows.length) { alert('Planilha vazia ou sem dados reconhecidos.'); return; }
      const novos = xlsxMapPedidosLiberadosRows(rows);
      if (!novos.length) { alert('Nenhum pedido reconhecido. Verifique se o arquivo segue o modelo correto.'); return; }
      pedidos = novos;
      renderPedidos();
      showTab('pedidos');
    } catch (err) {
      alert('Erro ao processar o arquivo: ' + err.message);
    } finally {
      input.value = '';
    }
  };
  reader.readAsArrayBuffer(input.files[0]);
}
function limparTodosPedidos() {
  if (!pedidos.length) return;
  if (!confirm('Remover todos os ' + pedidos.length + ' pedido(s) carregados?')) return;
  pedidos = [];
  renderPedidos();
}
function baixarModeloPedidos() {
  if (typeof XLSX === 'undefined') { alert('SheetJS não carregado.'); return; }
  const headers = [
    'Clientes Id ERP', 'Clientes Razão Social', 'Cidade Entrega',
    'Terminal', 'Material', 'Produto', 'Volume Pedido (m³)',
    'No. Ordem SAP', 'Data Entrega Logística',
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers]);
  ws['!cols'] = headers.map(() => ({ wch: 26 }));
  XLSX.utils.book_append_sheet(wb, ws, 'Pedidos');
  XLSX.writeFile(wb, 'modelo_pedidos_liberados.xlsx');
}
// ═══════════════════════════════════════════════════════════════════════════
// CSV
// ═══════════════════════════════════════════════════════════════════════════
function handleCSV(event) {
  const file = event.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const lines = e.target.result.split('\n').filter(l => l.trim());
    const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
    for (let i = 1; i < lines.length; i++) {
      const vals = lines[i].split(',').map(v => v.trim());
      const obj  = {}; headers.forEach((h,idx) => obj[h] = vals[idx] || '');
      const produtos = [];
      for (let k = 1; k <= 5; k++) {
        const pKey = ['produto'+k,'produto_'+k, k===1?'produto':''].find(key => obj[key] !== undefined && obj[key] !== '');
        const vKey = ['volume'+k+'_m3','volume_'+k+'_m3', k===1?'volume_m3':''].find(key => obj[key] !== undefined && obj[key] !== '');
        if (pKey && vKey && parseFloat(obj[vKey]) > 0)
          produtos.push({produto: obj[pKey], volume: parseFloat(obj[vKey])});
      }
      if (!produtos.length) continue;
      pedidos.push({
        id: Date.now()+i,
        codigoSAP: obj['codigo_sap'] || '',
        cliente:   obj['cliente'] || `Cliente ${i}`,
        cidade:    obj['cidade'] || '-',
        lat:       parseFloat(obj['lat']) || 0,
        lon:       parseFloat(obj['lon']) || 0,
        terminal:  obj['terminal'] || '',
        restricao: obj['restricao_horario'] || null,
        tiposCaminhao: [],
        produtos,
      });
    }
    renderPedidos(); showTab('pedidos');
  };
  reader.readAsText(file);
}
// ═══════════════════════════════════════════════════════════════════════════
// VEÍCULOS
// ═══════════════════════════════════════════════════════════════════════════
function aplicarTurnoVeiculo() {
  const turno = document.getElementById('v-turno')?.value;
  const inicio = document.getElementById('v-jornada-inicio');
  const fim = document.getElementById('v-jornada-fim');
  if (!inicio || !fim) return;
  const presets = {
    manha: ['06:00','14:00'],
    tarde: ['14:00','22:00'],
    noite: ['22:00','06:00'],
    integral: ['00:00','23:59'],
  };
  if (presets[turno]) {
    inicio.value = presets[turno][0];
    fim.value = presets[turno][1];
  }
}
function limparFormVeiculo() {
  numComps = 0;
  editandoVeiculoId = null;
  document.getElementById('form-veiculo-title').textContent = 'Novo veículo';
  document.getElementById('v-compartimentos').innerHTML = '<div class="sec-title">Compartimentos (m³)</div>';
  document.getElementById('v-placa').value = '';
  document.getElementById('v-transportadora').value = '';
  document.getElementById('v-tipo').value = 'Truck';
  document.getElementById('v-terminal').innerHTML = optsTerminais('');
  document.getElementById('v-cidade-base').value = '';
  document.getElementById('v-vel-carregado').value = 45;
  document.getElementById('v-vel-vazio').value = 55;
  document.getElementById('v-turno').value = 'personalizado';
  document.getElementById('v-jornada-inicio').value = '06:00';
  document.getElementById('v-jornada-fim').value = '18:00';
  document.getElementById('v-tempo-perdido').value = '01:00';
  document.getElementById('v-identidade-petronas').checked = false;
  document.getElementById('v-disponibilidade').value = 'Disponível';
  document.getElementById('v-contrato').value = 'Dedicado';
  document.getElementById('v-motorista-diurno').value = '';
  document.getElementById('v-motorista-noturno').value = '';
}
function abrirFormVeiculo(id=null) {
  limparFormVeiculo();
  if (id !== null) {
    const v = veiculos.find(x => x.id === id);
    if (!v) return;
    editandoVeiculoId = id;
    document.getElementById('form-veiculo-title').textContent = 'Editar veículo';
    document.getElementById('v-placa').value = v.placa || '';
    document.getElementById('v-implemento').value = v.implemento || '';
    document.getElementById('v-transportadora').value = v.transportadora || '';
    document.getElementById('v-tipo').value = v.tipo || 'Bitrem';
    document.getElementById('v-terminal').innerHTML = optsTerminais(v.terminal || '');
    document.getElementById('v-cidade-base').value = v.cidadeBase || v.cidade || cidadeDoTerminal(v.terminal) || '';
    document.getElementById('v-vel-carregado').value = v.velMediaCarregado || 45;
    document.getElementById('v-vel-vazio').value = v.velMediaVazio || 55;
    document.getElementById('v-turno').value = v.turno || 'personalizado';
    document.getElementById('v-jornada-inicio').value = v.jornadaInicio || '06:00';
    document.getElementById('v-jornada-fim').value = v.jornadaFim || '18:00';
    const _tp = v.tempoPerdidoMin || 0;
    document.getElementById('v-tempo-perdido').value = `${String(Math.floor(_tp/60)).padStart(2,'0')}:${String(_tp%60).padStart(2,'0')}`;
    document.getElementById('v-identidade-petronas').checked = !!v.identidadePetronas;
    document.getElementById('v-disponibilidade').value = v.disponibilidade || 'Disponível';
    document.getElementById('v-contrato').value = v.contrato || 'Dedicado';
    document.getElementById('v-motorista-diurno').value = v.motoristaDiurno || '';
    document.getElementById('v-motorista-noturno').value = v.motoristaNt || '';
    (v.compartimentos || []).forEach(c => addCompartimento(c.cap, c.produto || ''));
    if (!(v.compartimentos || []).length) addCompartimento();
  } else {
    addCompartimento();
  }
  document.getElementById('form-veiculo').classList.remove('hidden');
  document.getElementById('form-veiculo').scrollIntoView({behavior:'smooth', block:'start'});
}
function cancelarFormVeiculo() {
  editandoVeiculoId = null;
  document.getElementById('form-veiculo').classList.add('hidden');
}
function addCompartimento(capacidade='', produtoFixado='') {
  numComps++;
  const id = numComps;
  const div = document.createElement('div');
  div.className = 'row';
  div.id = 'comp-row-' + id;
  div.innerHTML = `
    <div class="field"><label>Compartimento ${id} (m³)</label>
      <input type="number" step="0.5" id="comp-${id}" value="${capacidade!=='' ? capacidade : ''}" placeholder="Ex: 15"/>
    </div>
    <div class="field" style="max-width:150px;"><label>Produto fixo?</label>
      <select id="comprod-${id}"><option value="">Qualquer</option>${PRODUTOS.map(p=>`<option ${p===produtoFixado?'selected':''}>${p}</option>`).join('')}</select>
    </div>
    ${id>1?`<button class="btn btn-sm btn-danger" style="flex:0;margin-top:20px;" onclick="document.getElementById('comp-row-${id}').remove()">x</button>`:''}`;
  document.getElementById('v-compartimentos').appendChild(div);
}
function salvarVeiculo() {
  const comps = [];
  document.querySelectorAll('[id^="comp-row-"]').forEach(row => {
    const rid = row.id.replace('comp-row-','');
    const el  = document.getElementById('comp-'+rid);
    const pr  = document.getElementById('comprod-'+rid);
    if (el && el.value) comps.push({cap: parseFloat(el.value), produto: pr ? pr.value : ''});
  });
  if (!comps.length) { alert('Adicione ao menos um compartimento.'); return; }
  const terminal = document.getElementById('v-terminal').value;
  let cidadeBase = (document.getElementById('v-cidade-base').value || '').trim();
  if (!cidadeBase && terminal) cidadeBase = cidadeDoTerminal(terminal);
  if (!terminal && !cidadeBase) { alert('Informe terminal de origem ou cidade base do veículo.'); return; }
  const velMediaCarregado = parseFloat(document.getElementById('v-vel-carregado').value);
  const velMediaVazio = parseFloat(document.getElementById('v-vel-vazio').value);
  const turno = document.getElementById('v-turno').value || 'personalizado';
  const jornadaInicio = document.getElementById('v-jornada-inicio').value;
  const jornadaFim = document.getElementById('v-jornada-fim').value;
  const jornadaMin = duracaoJornadaMin(jornadaInicio, jornadaFim);
  const tempoPerdidoMin = parseHoraMin(document.getElementById('v-tempo-perdido').value) || 0;
  if (isNaN(velMediaCarregado) || velMediaCarregado <= 0) { alert('Informe a velocidade média carregado.'); return; }
  if (isNaN(velMediaVazio) || velMediaVazio <= 0) { alert('Informe a velocidade média vazio.'); return; }
  if (!jornadaInicio || !jornadaFim || jornadaMin <= 0) { alert('Informe início e fim válidos para a jornada do veículo.'); return; }
  const dados = {
    placa: document.getElementById('v-placa').value || 'SEM-PLACA',
    implemento: document.getElementById('v-implemento').value.trim(),
    transportadora: document.getElementById('v-transportadora').value.trim(),
    tipo:  document.getElementById('v-tipo').value,
    terminal,
    cidadeBase,
    turno,
    jornadaInicio,
    jornadaFim,
    jornadaMin,
    tempoPerdidoMin,
    velMediaCarregado,
    velMediaVazio,
    compartimentos: comps,
    capacidade: comps.reduce((s,c) => s+c.cap, 0),
    identidadePetronas: document.getElementById('v-identidade-petronas').checked,
    disponibilidade:  document.getElementById('v-disponibilidade').value || 'Disponível',
    contrato:         document.getElementById('v-contrato').value || 'Dedicado',
    motoristaDiurno:  document.getElementById('v-motorista-diurno').value.trim(),
    motoristaNt:      document.getElementById('v-motorista-noturno').value.trim(),
  };
  if (editandoVeiculoId !== null) {
    const idx = veiculos.findIndex(v => v.id === editandoVeiculoId);
    if (idx !== -1) veiculos[idx] = {...veiculos[idx], ...dados};
  } else {
    veiculos.push({id: Date.now(), ...dados});
  }
  renderVeiculos();
  cancelarFormVeiculo();
}
function removerVeiculo(id) {
  veiculos = veiculos.filter(v => v.id !== id);
  renderVeiculos();
}
function renderVeiculos() {
  // Sempre re-sincroniza motoristas e disponibilidade do painel do dia antes de renderizar.
  // Usa a data do seletor do roteirizador (rot-data-operacao) como prioridade,
  // pois é a data para a qual a programação está sendo feita.
  // Fallback para dateStr(S.dateOffset) caso o input não exista.
  const inputData = document.getElementById('rot-data-operacao')?.value;
  const ds = inputData || dateStr(S.dateOffset);
  sincronizarDisponibilidadeVeiculos(ds).then(_renderVeiculosInterno).catch(_renderVeiculosInterno);
}
function _renderVeiculosInterno() {
  atualizarFiltroMapaTransportadora();
  const el = document.getElementById('veiculos-list');
  if (!veiculos.length) { el.innerHTML = ''; return; }
  const filtroPlaca    = valId('f-vei-placa');
  const filtroTransp   = valId('f-vei-transp');
  const filtroTerminal = valId('f-vei-terminal');
  const filtroCidade   = valId('f-vei-cidade');
  const filtroTipo     = valId('f-vei-tipo');
  const lista = veiculos.filter(v =>
    containsFiltro(v.placa, filtroPlaca) &&
    containsFiltro(v.transportadora, filtroTransp) &&
    containsFiltro(baseVeiculoLabel(v), filtroTerminal) &&
    containsFiltro(cidadeBaseVeiculo(v), filtroCidade) &&
    (!filtroTipo || v.tipo === filtroTipo)
  );
  if (!lista.length) { el.innerHTML = '<div class="empty">Nenhum veículo encontrado para os filtros.</div>'; return; }
  el.innerHTML = lista.map(v => `
    <div class="card">
      <div style="display:flex;align-items:flex-start;gap:10px;">
        <div style="flex:1;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:6px;">
            <span style="font-weight:700;font-size:15px;font-family:var(--font-cond);letter-spacing:0.05em;">${v.placa}</span>
            ${v.implemento ? `<span class="tag tag-gray" style="font-size:10px;letter-spacing:.04em;">⊞ ${v.implemento}</span>` : ''}
            ${v.transportadora ? `<span class="tag tag-lime" style="font-size:9px;">${v.transportadora}</span>` : ''}
            <span class="tag tag-gray">${v.tipo}</span>
            ${(v.disponibilidade||'Disponível') === 'Indisponível'
              ? (() => {
                  const motivo = v._motivoIndisponivel;
                  if (motivo === 'inativo') return `<span class="tag" style="font-size:9px;background:#FEE2E2;color:#B91C1C;border:1px solid #FECACA;" title="Veículo inativo no cadastro Firestore">⛔ Inativo</span>`;
                  if (motivo === 'sem_registro') return `<span class="tag" style="font-size:9px;background:#FFF7ED;color:#92400E;border:1px solid #FDE68A;" title="Sem registro de disponibilidade para este dia">⚠ Sem registro no painel</span>`;
                  const labels = {manutencao:'Manutenção',folga:'Folga',programado:'Em viagem',indisponivel:'Indisponível'};
                  return `<span class="tag" style="font-size:9px;background:#FEE2E2;color:#B91C1C;border:1px solid #FECACA;" title="Status do Painel de Disponibilidade">⛔ ${labels[motivo]||'Indisponível'}</span>`;
                })()
              : `<span class="tag" style="font-size:9px;background:#D1FAE5;color:#065F46;border:1px solid #6EE7B7;" title="Status do Painel de Disponibilidade">✓ Disponível</span>`}
            ${(v.contrato||'Dedicado') === 'Spot'
              ? `<span class="tag" style="font-size:9px;background:#FEF3C7;color:#92400E;border:1px solid #FCD34D;">Spot</span>`
              : `<span class="tag" style="font-size:9px;background:#EEF3FF;color:#1E40AF;border:1px solid #93C5FD;">Dedicado</span>`}
            ${v.identidadePetronas ? `<span class="tag tag-yellow" style="font-size:9px;">⬡ ID Petronas</span>` : ''}
            ${v.terminal ? termTag(v.terminal) : `<span class="tag tag-gray">${baseVeiculoLabel(v)}</span>`}
            <span style="font-size:12px;color:#4A6535;">${v.capacidade.toFixed(1)} m³/viagem</span>
            <span class="tag tag-blue" style="font-size:9px;">${v.jornadaInicio||'06:00'}-${v.jornadaFim||'18:00'} (${Math.round((v.jornadaMin||duracaoJornadaMin(v.jornadaInicio||'06:00', v.jornadaFim||'18:00'))/60)}h)</span>
            <span class="tag tag-purple" style="font-size:9px;">V/C: ${(v.velMediaVazio||55).toFixed(0)} / ${(v.velMediaCarregado||45).toFixed(0)} km/h</span>
          </div>
          ${(v.motoristaDiurno || v.motoristaNt) ? `
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:6px;padding:5px 10px;background:rgba(74,101,53,0.1);border-radius:6px;border:1px solid rgba(74,101,53,0.2);">
            <span style="font-size:10px;color:#4A6535;font-weight:600;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">🧑‍✈️ Motorista</span>
            ${v.motoristaDiurno ? `<span style="font-size:11px;color:#1F2937;font-weight:500;">☀️ ${v.motoristaDiurno}</span>` : ''}
            ${v.motoristaDiurno && v.motoristaNt ? `<span style="color:#9CA3AF;font-size:10px;">·</span>` : ''}
            ${v.motoristaNt ? `<span style="font-size:11px;color:#1F2937;font-weight:500;">🌙 ${v.motoristaNt}</span>` : ''}
            ${v._motoristaPainel ? `<span style="font-size:9px;background:#D1FAE5;color:#065F46;border-radius:4px;padding:1px 5px;margin-left:2px;">painel</span>` : `<span style="font-size:9px;background:#F3F4F6;color:#6B7280;border-radius:4px;padding:1px 5px;margin-left:2px;">cadastro</span>`}
          </div>` : `
          <div style="display:flex;align-items:center;gap:6px;margin-top:6px;padding:4px 10px;background:#FFF7ED;border-radius:6px;border:1px solid #FDE68A;">
            <span style="font-size:10px;color:#92400E;">⚠ Motorista não informado no painel para este dia</span>
          </div>`}
          <div class="pills">
            ${v.compartimentos.map((c,i)=>`<span class="pill">C${i+1}: ${c.cap}m³${c.produto?' - '+c.produto:''}</span>`).join('')}
          </div>
        </div>
        <div style="display:flex;gap:5px;flex-shrink:0;">
          <button class="btn btn-sm" onclick="abrirFormVeiculo(${v.id})">Editar</button>
          <button class="btn btn-sm btn-danger" onclick="removerVeiculo(${v.id})">x</button>
        </div>
      </div>
    </div>`).join('');
} // fim _renderVeiculosInterno
// EXEMPLOS
function carregarExemplos() {
  if (!terminaisCad.length) carregarExemplosTerminais();
  if (!clientes.length)     carregarExemplosClientes();
  pedidos = [
    {id:1,codigoSAP:'1000123',cliente:'Posto Ipiranga Centro',  cidade:'Campinas', lat:-22.9056,lon:-47.0608,terminal:'Paulínia TRANSO',  restricao:'06:00–10:00',tiposCaminhao:[],identidadePetronas:true, produtos:[{produto:'2000031 - PETRONAS GASOLINA COMUM',volume:15},{produto:'2000033 - PETRONAS ETANOL COMUM',volume:10}]},
    {id:2,codigoSAP:'1000456',cliente:'Posto Shell Viracopos',  cidade:'Campinas', lat:-23.0074,lon:-47.1344,terminal:'Paulínia TRANSO',  restricao:null,          tiposCaminhao:[],identidadePetronas:false,produtos:[{produto:'2000035 - PETRONAS DIESEL S10',volume:30}]},
    {id:3,codigoSAP:'1000789',cliente:'Distribuidora Santos',   cidade:'Santos',   lat:-23.9618,lon:-46.3322,terminal:'Cubatão IPIRANGA Nexta',    restricao:'08:00–17:00', tiposCaminhao:[],identidadePetronas:true, produtos:[{produto:'2000035 - PETRONAS DIESEL S10',volume:25},{produto:'2000036 - PETRONAS DIESEL S500',volume:20}]},
    {id:4,codigoSAP:'1001012',cliente:'Posto Ale Litoral',      cidade:'Guarujá',  lat:-23.9929,lon:-46.2567,terminal:'Cubatão IPIRANGA Nexta',    restricao:null,          tiposCaminhao:[],identidadePetronas:false,produtos:[{produto:'2000032 - PETRONAS PRIMAX GASOLINA',volume:15}]},
    {id:5,codigoSAP:'1001345',cliente:'Frota Transportes MG',   cidade:'Sorocaba', lat:-23.5015,lon:-47.4526,terminal:'Paulínia TRANSO',  restricao:null,          tiposCaminhao:[],identidadePetronas:false,produtos:[{produto:'2000036 - PETRONAS DIESEL S500',volume:20}]},
    {id:6,codigoSAP:'1001678',cliente:'Posto BR Bandeirantes',  cidade:'Jundiaí',  lat:-23.1864,lon:-46.8842,terminal:'São Caetano VIBRA SIM', restricao:'07:00–11:00', tiposCaminhao:[],identidadePetronas:true, produtos:[{produto:'2000033 - PETRONAS ETANOL COMUM',volume:8},{produto:'2000031 - PETRONAS GASOLINA COMUM',volume:12}]},
    {id:7,codigoSAP:'1002001',cliente:'Rede TeraCom Sorocaba',  cidade:'Sorocaba', lat:-23.5300,lon:-47.4600,terminal:'Paulínia TRANSO',  restricao:null,          tiposCaminhao:[],identidadePetronas:false,produtos:[{produto:'2000047 - PETRONAS DYNAMIC DIESEL S10',volume:18},{produto:'2000035 - PETRONAS DIESEL S10',volume:5}]},
  ];
  renderPedidos();
  renderVeiculos();
}
// ═══════════════════════════════════════════════════════════════════════════
// OTIMIZAÇÃO
// ═══════════════════════════════════════════════════════════════════════════
async function otimizar(modo = 'padrao', dataCarregamento = null) {
  // ══════════════════════════════════════════════════════════════════════════
  // MOTOR DE OTIMIZAÇÃO
  // Regras:
  //  • Respeitar horários do terminal (abertura e fechamento)
  //  • Respeitar janela de entrega do cliente, tipo de veículo e ID Petronas
  //  • Compartimentos: all-or-nothing por produto (nunca parcial)
  //  • Jornada = apenas tempo produtivo (carga + trajeto + descarga + retorno)
  //    Esperas (terminal, janela, overnight) NÃO contam como jornada
  //  • Overnight: quando janela exige entrega antes da abertura do terminal,
  //    carrega antes do fechamento do dia anterior e aguarda na base
  //  • Prioridade: 1) ID Petronas  2) maior volume do pedido
  //  • Minimizar veículos; agrupar pedidos da mesma cidade quando possível
  // ══════════════════════════════════════════════════════════════════════════
  if (!pedidos.length)  { showTab('pedidos');  alert('Adicione ao menos um pedido.'); return; }
  if (!veiculos.length) { showTab('veiculos'); alert('Adicione ao menos um veículo.'); return; }
  Object.keys(_motoristasOverride).forEach(k => delete _motoristasOverride[k]);
  const todosBtn = document.querySelectorAll('.btn-otimizar-ded');
  todosBtn.forEach(b => { b.disabled = true; b.style.opacity = '0.7'; });
  const btn = document.querySelector('.btn-otimizar-ded');
  const txtOriginal = btn ? btn.textContent : '';
  if (btn) btn.textContent = 'Calculando...';
  // Backup da lista completa de veículos — restaurado no finally
  // para que a aba Veículos & Turnos continue mostrando todos após roteirizar
  let _veiculosTodos = veiculos.slice();
  try {
  // ── Ressincroniza disponibilidade com a data EXATA do carregamento ──────────
  // Garante que o status do painel seja sempre o do dia correto, independente
  // de qual data estava selecionada no header quando a aba foi aberta.
  // Sem isso, se o usuário roteirizar para amanhã mas o painel estava em hoje,
  // veículos com 'programado' ou outros status não-disponível passariam.
  {
    const _dataIso = (() => {
      const d = dataCarregamento || new Date();
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    })();
    try { await sincronizarDisponibilidadeVeiculos(_dataIso); } catch(e) {
      console.warn('[otimizar] Falha ao ressincronizar disponibilidade:', e);
    }
  }
  // ── Remove veículos não-disponíveis antes do motor rodar ───────────────────
  // Regras:
  //   • 'Disponível'  → sempre entra (respeitando _horarioDisponivelAPartirDe se houver)
  //   • 'Manutenção'  → entra SOMENTE se previsão de retorno (_horarioDisponivelAPartirDe)
  //                     for antes ou igual ao jornadaFim do veículo (há janela de uso no dia)
  //   • Demais (Programado/Em viagem, Folga, sem_registro) → sempre bloqueado
  const veiculosIndisponiveis = veiculos.filter(v => {
    if ((v.disponibilidade || 'Indisponível') !== 'Disponível') return true; // bloqueia
    // Veículo marcado como disponível mas em manutenção com horário de retorno:
    // verifica se ainda há janela útil no dia (retorno antes do fim da jornada)
    if (v._emManutencao && v._horarioDisponivelAPartirDe) {
      const retornoMin = parseHoraMin(v._horarioDisponivelAPartirDe);
      const fimMin     = parseHoraMin(v.jornadaFim || '18:00');
      // Se o retorno for igual ou depois do fim da jornada, não há janela útil
      if (!isNaN(retornoMin) && !isNaN(fimMin) && retornoMin >= fimMin) return true;
    }
    return false;
  });
  veiculos = veiculos.filter(v => !veiculosIndisponiveis.includes(v));
  if (veiculosIndisponiveis.length) {
    console.log(`[Otimizador] ${veiculosIndisponiveis.length} veículo(s) removido(s) por indisponibilidade:`,
      veiculosIndisponiveis.map(v => `${v.placa} (${v._motivoIndisponivel || 'indisponível'})`).join(', '));
    // Remove também do ultimoResultado para que terminais travados não restaurem veículos indisponíveis
    if (ultimoResultado) {
      const idsIndisp = new Set(veiculosIndisponiveis.map(v => v.id));
      idsIndisp.forEach(id => { delete ultimoResultado[id]; });
    }
  }
  // Ressincroniza controleTempo antes de usar como base para terminais travados,
  // corrigindo qualquer drift causado por edições manuais pós-otimização.
  recalcularControleTempo();
  // ── Estado ────────────────────────────────────────────────────────────────
  const resultado     = {};  // veiculo.id → [Viagem]
  const controleTempo = {};  // veiculo.id → { limiteMin, usadoMin }
                             //   usadoMin = APENAS tempo produtivo
  // IDs de pedidos já alocados em terminais travados (não entram na fila)
  const pedidosLocadosIds = new Set();
  veiculos.forEach(v => {
    const jornadaMin = v.jornadaMin || duracaoJornadaMin(v.jornadaInicio || '06:00', v.jornadaFim || '18:00') || 0;
    const termV = baseVeiculoLabel(v);
    if (lockedTerminals.has(termV) && ultimoResultado && (v.disponibilidade || 'Disponível') !== 'Indisponível') {
      // Terminal travado: restaura exatamente o estado atual
      resultado[v.id] = (ultimoResultado[v.id] || [])
        .filter(vi => !vi._vazio)
        .map(vi => JSON.parse(JSON.stringify(vi)));
      const ctAnterior = ultimoControleTempo?.[v.id];
      controleTempo[v.id] = ctAnterior
        ? { limiteMin: jornadaMin, usadoMin: ctAnterior.usadoMin }
        : { limiteMin: jornadaMin, usadoMin: 0 };
      resultado[v.id].forEach(vi =>
        vi.paradas.forEach(pa => { if (pa.pedido?.id) pedidosLocadosIds.add(pa.pedido.id); })
      );
    } else {
      resultado[v.id]     = [];
      controleTempo[v.id] = { limiteMin: jornadaMin, usadoMin: 0 };
    }
  });
  // ── Helpers internos ──────────────────────────────────────────────────────
  // Minuto de início da jornada do veículo
  // jIni(v): retorna o minuto de início efetivo da jornada do veículo.
  // Se o painel registrou horário de disponibilidade (_horarioDisponivelAPartirDe),
  // usa o MAIOR entre jornadaInicio e esse horário — o veículo só pode sair
  // quando ambos os critérios forem satisfeitos.
  const jIni = v => {
    const jornadaMin = parseHoraMin(v.jornadaInicio || '06:00');
    const base = isNaN(jornadaMin) ? 360 : jornadaMin;
    if (v._horarioDisponivelAPartirDe) {
      const dispMin = parseHoraMin(v._horarioDisponivelAPartirDe);
      if (!isNaN(dispMin) && dispMin > base) return dispMin;
    }
    return base;
  };
  // Clock absoluto atual do veículo = jornada_início + sum(tempoConsumidoMin das viagens)
  const clockV = v => jIni(v)
    + Math.min(resultado[v.id].length, doisTurnos(v) ? 2 : 1) * (v.tempoPerdidoMin || 0)
    + resultado[v.id].reduce((s, vi) => s + (vi.tempoConsumidoMin || 0), 0);
  // Base de datas: data de carregamento informada pelo usuário no modal.
  // Fallback: menor data de entrega do lote (comportamento legado sem modal).
  const datasEntregaValidas = pedidos
    .map(p => parseDataBr(p?.dataEntregaLogistica))
    .filter(Boolean);
  const _minDataLote = datasEntregaValidas.length
    ? new Date(Math.min(...datasEntregaValidas.map(d => d.getTime())))
    : null;
  const baseDataEntrega = dataCarregamento || _minDataLote || new Date();
  const diaAlvoPedido = (pedido) => {
    const d = parseDataBr(pedido?.dataEntregaLogistica);
    if (!d) return null;
    return Math.round((d.getTime() - baseDataEntrega.getTime()) / 86400000);
  };
  // Veículo atende as restrições do pedido (tipo, Petronas, terminal)?
  const veicOk = (v, p) =>
    (!p.tiposCaminhao?.length || p.tiposCaminhao.includes(v.tipo)) &&
    (!p.identidadePetronas || !!v.identidadePetronas) &&
    veiculoAtendeTerminal(v, p.terminal);
  const totalVolProdutos = (produtos=[]) => produtos.reduce((s, pr) => s + (pr?.volume || 0), 0);
  const produtosPendentesPedido = (pedido) => {
    const alocadoPorProduto = {};
    Object.values(resultado).forEach(viagens => {
      (viagens || []).forEach(vi => {
        (vi.paradas || []).forEach(pa => {
          if (pa?.pedido?.id !== pedido.id) return;
          (pa.itens || []).forEach(it => {
            const k = it.produto || '';
            alocadoPorProduto[k] = (alocadoPorProduto[k] || 0) + (it.volume || 0);
          });
        });
      });
    });
    const saldoPorProduto = { ...alocadoPorProduto };
    const pend = [];
    for (const pr of (pedido.produtos || [])) {
      const k = pr.produto || '';
      const disponivel = saldoPorProduto[k] || 0;
      const consumido = Math.min(pr.volume || 0, disponivel);
      saldoPorProduto[k] = Math.max(0, disponivel - consumido);
      const restante = Math.max(0, (pr.volume || 0) - consumido);
      if (restante > 0.0001) pend.push({ ...pr, volume: restante });
    }
    return pend;
  };
  // Simula se todos os produtos do pedido cabem EXATAMENTE nos compartimentos
  // (cópia — sem efeito colateral)
  const podeFitar = (produtos, compsDisp) => {
    const copia = compsDisp.map(c => ({ ...c }));
    for (const pr of produtos) {
      const el  = copia.filter(c => compElegivelProduto(c, pr.produto) && c.disponivel > 0);
      // Tenta exact primeiro (ideal); se não houver combinação exata, aceita >= (compartimento maior)
      let aloc = selecionarCompartimentos(el, pr.volume, 'exact');
      if (!aloc) aloc = selecionarCompartimentos(el, pr.volume, 'geq');
      if (!aloc) return false;
      aloc.forEach(c => { c.disponivel = 0; });
    }
    return true;
  };
  // veicOkCompartimentos: extensão de veicOk que também verifica se os produtos
  // do pedido cabem nos compartimentos do veículo.
  // Declarado após podeFitar e produtosPendentesPedido para evitar referência antecipada.
  const veicOkCompartimentos = (v, p) => {
    if (!veicOk(v, p)) return false;
    const prods = produtosPendentesPedido(p);
    if (!prods.length) return true;
    const compsSimul = criarCompsDisp(v);
    if (!compsSimul.length) return true; // veículo sem compartimentos fixos não bloqueia
    return podeFitar(prods, compsSimul);
  };

  // Aloca todos os produtos do pedido em compartimentos reais (commit)
  // Retorna false se algum produto não couber (não deveria ocorrer após podeFitar)
  const commitPedido = (viagem, pedido, detalheParada, custoRelogio, custoProdutivo, v, produtosSelecionados = pedido.produtos) => {
    let primeiro = true;
    for (const pr of (produtosSelecionados || [])) {
      const el  = viagem.compsDisp.filter(c => compElegivelProduto(c, pr.produto) && c.disponivel > 0);
      let aloc = selecionarCompartimentos(el, pr.volume, 'exact');
      if (!aloc) aloc = selecionarCompartimentos(el, pr.volume, 'geq'); // aceita compartimento maior
      if (!aloc) return false;
      const item = { pedidoId: pedido.id, pedido, produto: pr.produto, volume: pr.volume, restante: 0, ordemSAP: pr.ordemSAP || '' };
      // Primeiro produto do pedido → cria a parada com timing; demais → adiciona à parada existente
      alocarItem(viagem, item, pr.volume, primeiro ? detalheParada : null, 0, aloc);
      primeiro = false;
    }
    viagem.tempoConsumidoMin = (viagem.tempoConsumidoMin || 0) + custoRelogio;
    controleTempo[v.id].usadoMin += custoProdutivo;
    return true;
  };
  // Cria nova viagem vazia para o veículo
  const novaViagem = (v, terminal, esperaTerminal) => ({
    compsDisp: criarCompsDisp(v),
    paradas: [], quebras: [],
    tempoConsumidoMin: 0,
    terminalOrigem: terminal,
    esperaTerminalMin: esperaTerminal,
  });
  // Candidatos compatíveis para um pedido, ordenados:
  //   1. Veículos já indo à mesma cidade (reaproveitamento de rota)
  //   2. Menor capacidade suficiente (best-fit — preserva trucks grandes)
  //   3. Mais jornada produtiva disponível
  const candidatos = (pedido, volMin=null, permitirCapacidadeMenor=false) => {
    const vol     = (volMin === null || volMin === undefined) ? totalVolPedido(pedido) : volMin;
    const cidadeP = (pedido.cidade || '').toLowerCase().trim();
    return veiculos
      .filter(v => !lockedTerminals.has(baseVeiculoLabel(v)) && (v.disponibilidade || 'Disponível') !== 'Indisponível' && veicOkCompartimentos(v, pedido) && (permitirCapacidadeMenor || v.capacidade >= vol - 0.001))
      .sort((a, b) => {
        // 1. Dedicado antes de Spot
        const aDedicado = (a.contrato || 'Dedicado') !== 'Spot' ? 1 : 0;
        const bDedicado = (b.contrato || 'Dedicado') !== 'Spot' ? 1 : 0;
        if (bDedicado !== aDedicado) return bDedicado - aDedicado;
        // 2. Veículo já carrega outra parte deste mesmo pedido (quebra)
        //    → consolida no mesmo caminhão em vez de dispersar para outra cidade
        const aTemPedido = resultado[a.id].some(vi => vi.paradas.some(pa => pa.pedido?.id === pedido.id)) ? 1 : 0;
        const bTemPedido = resultado[b.id].some(vi => vi.paradas.some(pa => pa.pedido?.id === pedido.id)) ? 1 : 0;
        if (bTemPedido !== aTemPedido) return bTemPedido - aTemPedido;
        // 3. Mesma cidade
        const aCidade = resultado[a.id].some(vi => vi.paradas.some(p => (p.pedido?.cidade||'').toLowerCase().trim() === cidadeP)) ? 1 : 0;
        const bCidade = resultado[b.id].some(vi => vi.paradas.some(p => (p.pedido?.cidade||'').toLowerCase().trim() === cidadeP)) ? 1 : 0;
        if (bCidade !== aCidade) return bCidade - aCidade;
        // 4+5. Score composto: best-fit + bônus por veículo ativo.
        //   Veículo ativo recebe desconto de ATIVO_BONUS m³ no score,
        //   ou seja: tolera-se usar um ativo até ATIVO_BONUS m³ maior que o ideal
        //   antes de ativar um novo veículo de encaixe mais preciso.
        //   Exemplo (BONUS=4, pedido 15 m³):
        //     ativo 20 m³ → score 16  vs  inativo 15 m³ → score 15  → inativo ganha
        //     ativo 20 m³ → score 16  vs  inativo 20 m³ → score 20  → ativo ganha
        const ATIVO_BONUS = 4;
        const aAtivo = resultado[a.id].length > 0 ? 1 : 0;
        const bAtivo = resultado[b.id].length > 0 ? 1 : 0;
        // Modo dedicado: bônus negativo para dedicados → prefere ativar os ociosos (spreading)
        const aBonus = (modo === 'dedicado' && (a.contrato || 'Dedicado') !== 'Spot') ? -8 : ATIVO_BONUS;
        const bBonus = (modo === 'dedicado' && (b.contrato || 'Dedicado') !== 'Spot') ? -8 : ATIVO_BONUS;
        const aScore = a.capacidade - aAtivo * aBonus;
        const bScore = b.capacidade - bAtivo * bBonus;
        if (Math.abs(aScore - bScore) > 0.001) return aScore - bScore;
        // Desempate quando scores iguais: preferir ativo (consolida antes de ativar)
        if (bAtivo !== aAtivo) return bAtivo - aAtivo;
        // 6. Mais jornada disponível
        return (controleTempo[b.id].limiteMin - controleTempo[b.id].usadoMin)
             - (controleTempo[a.id].limiteMin - controleTempo[a.id].usadoMin);
      });
  };
  // Tenta encaixar pedido em viagem JÁ EXISTENTE do veículo v.
  // Usa best-insertion: avalia todas as posições possíveis (início, meio, fim)
  // e tenta da mais barata geograficamente para a mais cara, retornando a
  // primeira que satisfaz janela de entrega e limite de jornada.
  const tentarEncaixe = (pedido, v, viagem, produtosSelecionados = pedido.produtos, opts = {}) => {
    const permitirExcederJornada = !!opts.permitirExcederJornada;
    // Pedido sem terminal → aceita em qualquer viagem (usa o terminal de origem da viagem)
    if (pedido.terminal && viagem.terminalOrigem !== pedido.terminal) return null;
    if (!podeFitar(produtosSelecionados, viagem.compsDisp)) return null;
    const idxV = resultado[v.id].indexOf(viagem);
    const n    = viagem.paradas.length;
    // Cap de jornada: normal = 100%, fallback de sobra = 130%
    const limiteEfetivo = permitirExcederJornada
      ? controleTempo[v.id].limiteMin * 1.05
      : controleTempo[v.id].limiteMin;
    for (const pos of posicoesCandidatas(viagem, v, pedido)) {
      const isEnd = pos.idx >= n;
      const det   = isEnd
        ? dadosIncrementoParada(viagem, v, pedido, viagem.terminalOrigem)
        : dadosInsercaoEmPosicao(viagem, v, pedido, viagem.terminalOrigem, pos.idx);
      const _nmb = doisTurnos(v) ? 2 : 1;
      const chegAbs = isEnd
        ? chegadaPrevistaAbsMin(viagem, det, resultado[v.id], idxV, jIni(v), v.tempoPerdidoMin || 0, _nmb)
        : chegadaInsercaoAbsMin(viagem, det, pos.idx, resultado[v.id], idxV, jIni(v), v.tempoPerdidoMin || 0, _nmb);
      const aval = avaliarRestricaoPedido(pedido, chegAbs, diaAlvoPedido(pedido));
      if (!aval.ok) continue;
      // Inserção no meio: verifica se o atraso introduzido empurra stops subsequentes
      // para além das suas janelas de entrega (o motor só checa o stop inserido, não os demais).
      if (!isEnd && det.incrementoMin > 0.001) {
        let tSub = inicioViagemAbsMin(resultado[v.id], idxV, jIni(v), v.tempoPerdidoMin || 0, _nmb)
                 + (viagem.esperaTerminalMin || 0);
        let violou = false;
        for (let j = 0; j < n; j++) {
          const pj = viagem.paradas[j];
          if (j === 0) tSub += (pj.tempoCarregamentoMin || 0);
          tSub += (pj.deslocCarregadoMin || 0);
          if (j >= pos.idx) {
            const avalSub = avaliarRestricaoPedido(pj.pedido, tSub + det.incrementoMin, diaAlvoPedido(pj.pedido));
            if (!avalSub.ok) { violou = true; break; }
          }
          tSub += (pj.tempoEsperaRestricaoMin || 0) + (pj.tempoDescargaMin || 0) + (pj.deslocVazioMin || 0);
        }
        if (violou) continue;
      }
      let prod;
      if (isEnd) {
        // Para append ao final: desconta retorno anterior para evitar dupla contagem
        const retornoAnteriorMin = (n ? (viagem.paradas[n - 1]?.deslocVazioMin || 0) : 0);
        prod = Math.max(0, -retornoAnteriorMin + (det.deslocCarregadoMin || 0) + (det.tempoDescargaMin || 0) + (det.deslocVazioMin || 0));
      } else {
        prod = det.incrementoMin;
      }
      const excesso2 = (controleTempo[v.id].usadoMin + prod) - limiteEfetivo;
      if (!permitirExcederJornada && excesso2 > 0.001) continue;
      det.tempoEsperaRestricaoMin = aval.esperaMin || 0;
      if (excesso2 > 0.001) det._jornadaExcedenteMin = excesso2; // registra excesso para alerta
      return { detalhe: det, custoRelogio: det.incrementoMin + (aval.esperaMin || 0), custoProdutivo: prod };
    }
    return null;
  };
  // Tenta iniciar NOVA VIAGEM para o pedido no veículo v (normal ou overnight)
  // Retorna { viagem, detalhe, custoRelogio, custoProdutivo } ou null
  const tentarNovaViagem = (pedido, v, produtosSelecionados = pedido.produtos, opts = {}) => {
    const permitirExcederJornada = !!opts.permitirExcederJornada;
    // Quando o pedido não tem terminal, usa o terminal/cidade-base do veículo como origem
    const terminalEfetivo = pedido.terminal || v.terminal || '';
    const ck       = clockV(v);
    const espTerm  = calcEsperaTerminal(terminalEfetivo, ck);
    if (!podeFitar(produtosSelecionados, criarCompsDisp(v))) return null;
    const vRef     = { paradas: [], terminalOrigem: terminalEfetivo, esperaTerminalMin: espTerm };
    const det      = dadosIncrementoParada(vRef, v, { ...pedido, terminal: terminalEfetivo }, terminalEfetivo);
    const chegAbs  = chegadaPrevistaAbsMin(vRef, det, resultado[v.id], resultado[v.id].length, jIni(v), v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1);
    const _diaAlvoBase = diaAlvoPedido(pedido);
    const aval = avaliarRestricaoPedido(pedido, chegAbs, _diaAlvoBase);
    // Tenta dia seguinte SOMENTE quando:
    // 1. Pedido tem janela de entrega definida (restricao não nulo)
    // 2. A chegada passou do FIM da janela de hoje (chegou tarde demais)
    // 3. O diaAlvo base é 0 (hoje) — não empurra pedidos futuros mais para frente
    const _temJanela = !!(pedido?.restricao);
    const _diaAlvoAjust = (!aval.ok && _temJanela && (_diaAlvoBase ?? 0) === 0)
      ? 1  // tenta entrega no dia seguinte
      : (_diaAlvoBase ?? 0);
    const avalFinal = (!aval.ok && _diaAlvoAjust !== (_diaAlvoBase ?? 0))
      ? avaliarRestricaoPedido(pedido, chegAbs, _diaAlvoAjust)
      : aval;
    const _diaAlvo = _diaAlvoAjust;
    // Para entregas com data futura (N dias à frente do carregamento), o veículo dispõe
    // de N+1 dias de jornada produtiva — escala o limite pelo horizonte real da viagem.
    const diasViagem = Math.max(1, _diaAlvo + 1);
    const limiteEfetivo = permitirExcederJornada
      ? controleTempo[v.id].limiteMin * diasViagem * 1.05
      : controleTempo[v.id].limiteMin * diasViagem;
    // Pausa produtiva: conta na jornada ao abrir a 2ª viagem (turno único)
    // ou a 3ª viagem (2 turnos — a 1ª pausa é do noturno, só relógio).
    // A pausa NÃO entra no custoRelogio/tempoConsumidoMin: é somada pelo
    // inicioViagemAbsMin para manter o relógio correto sem dupla contagem.
    const prodRefeicao = resultado[v.id].length === (doisTurnos(v) ? 2 : 1)
      ? (v.tempoPerdidoMin || 0) : 0;
    // ── Entrega no mesmo dia ou dia seguinte (espera curta no cliente, ≤ 32h) ──
    if (avalFinal.ok && (avalFinal.esperaMin || 0) <= 32 * 60) {
      const prod = det.incrementoMin + prodRefeicao;
      const excesso3a = (controleTempo[v.id].usadoMin + prod) - limiteEfetivo;
      if (!permitirExcederJornada && excesso3a > 0.001) return null;
      det.tempoEsperaRestricaoMin = avalFinal.esperaMin || 0;
      if (excesso3a > 0.001) det._jornadaExcedenteMin = excesso3a;
      if (_diaAlvoAjust > (_diaAlvoBase ?? 0)) det.overnight = true; // marcado para alerta de dia+1
      const vi = novaViagem(v, terminalEfetivo, espTerm);
      return { vi, detalhe: det,
               custoRelogio: espTerm + det.incrementoMin + (avalFinal.esperaMin || 0),
               custoProdutivo: prod };
    }
    // ── Overnight: carrega no terminal, aguarda e parte no horário certo ─────
    // Também cobre caso em que espera no cliente seria > 8h (viagem longa/futura).
    const itemRef = { pedido, produto: null, restante: 0 };
    // Tenta overnight para o diaAlvo calculado; se falhar, tenta dia seguinte
    let ov = calcOvernightViagem(v, itemRef, ck, terminalEfetivo, _diaAlvo);
    if (!ov && _diaAlvo === 0) ov = calcOvernightViagem(v, itemRef, ck, terminalEfetivo, 1);
    if (ov && controleTempo[v.id].usadoMin + ov.productiveMin + prodRefeicao <= limiteEfetivo + 0.001) {
      const ovDet = {
        ...ov,
        tempoEsperaRestricaoMin: 0,
        overnight: true,
        waitAfterLoadingMin: ov.waitAfterLoadingMin,
        origemDeslocamento: 'Terminal',
      };
      const vi = novaViagem(v, terminalEfetivo, ov.waitBeforeLoadingMin);
      return { vi, detalhe: ovDet, custoRelogio: ov.totalElapsedMin, custoProdutivo: ov.productiveMin + prodRefeicao };
    }
    // ── Fallback: espera longa (overnight indisponível) ─────────────────────
    if (avalFinal.ok) {
      const prod = det.incrementoMin + prodRefeicao;
      const excesso4 = (controleTempo[v.id].usadoMin + prod) - limiteEfetivo;
      if (!permitirExcederJornada && excesso4 > 0.001) return null;
      det.tempoEsperaRestricaoMin = avalFinal.esperaMin || 0;
      if (excesso4 > 0.001) det._jornadaExcedenteMin = excesso4;
      if (_diaAlvoAjust > (_diaAlvoBase ?? 0)) det.overnight = true;
      const vi = novaViagem(v, terminalEfetivo, espTerm);
      return { vi, detalhe: det,
               custoRelogio: espTerm + det.incrementoMin + (avalFinal.esperaMin || 0),
               custoProdutivo: prod };
    }
    return null;
  };
  const gerarSubconjuntosProdutos = (produtos=[]) => {
    const arr = Array.isArray(produtos) ? produtos : [];
    const n = arr.length;
    if (!n) return [];
    const out = [];
    const total = 1 << n;
    for (let mask = 1; mask < total; mask++) {
      const subset = [];
      let vol = 0;
      for (let i = 0; i < n; i++) {
        if ((mask & (1 << i)) === 0) continue;
        subset.push(arr[i]);
        vol += arr[i].volume || 0;
      }
      if (vol > 0.0001) out.push({ subset, vol });
    }
    out.sort((a, b) => (b.vol - a.vol) || (a.subset.length - b.subset.length));
    return out;
  };
  const melhorSubconjuntoQueCabe = (produtosPendentes, compsDisp) => {
    const subconjuntos = gerarSubconjuntosProdutos(produtosPendentes);
    for (const sc of subconjuntos) {
      if (podeFitar(sc.subset, compsDisp)) return sc;
    }
    return null;
  };
  // Fallback secundário: permite quebrar o pedido em múltiplas viagens/placas
  // após falhar a alocação all-or-nothing.
  const tentarQuebraPedido = (pedido, opts = {}) => {
    const maxPartes = Number.isFinite(opts.maxPartes) ? opts.maxPartes : Number.POSITIVE_INFINITY;
    const permitirExcederJornada = !!opts.permitirExcederJornada;
    let pendentes = Array.isArray(opts.produtosIniciais) && opts.produtosIniciais.length
      ? opts.produtosIniciais.map(pr => ({ ...pr }))
      : [...(pedido.produtos || [])];
    let partesAlocadas = 0;
    let alocadoAlgum = false;
    const viagensComQuebra = [];
    let guard = 0;
    while (pendentes.length && guard < 100 && partesAlocadas < maxPartes) {
      guard++;
      let melhor = null;
      const cands = candidatos(pedido, 0, true);
      for (const v of cands) {
        for (const vi of (resultado[v.id] || [])) {
          const sub = melhorSubconjuntoQueCabe(pendentes, vi.compsDisp);
          if (!sub) continue;
          const t = tentarEncaixe(pedido, v, vi, sub.subset, { permitirExcederJornada });
          if (!t) continue;
          const cand = { tipo: 'encaixe', v, vi, t, sub };
          if (!melhor || cand.sub.vol > melhor.sub.vol || (Math.abs(cand.sub.vol - melhor.sub.vol) < 0.001 && melhor.tipo === 'nova')) {
            melhor = cand;
          }
        }
        const subNova = melhorSubconjuntoQueCabe(pendentes, criarCompsDisp(v));
        if (!subNova) continue;
        const tNova = tentarNovaViagem(pedido, v, subNova.subset, { permitirExcederJornada });
        if (!tNova) continue;
        const candNova = { tipo: 'nova', v, vi: tNova.vi, t: tNova, sub: subNova };
        if (!melhor || candNova.sub.vol > melhor.sub.vol) {
          melhor = candNova;
        }
      }
      if (!melhor) break;
      const ok = commitPedido(
        melhor.vi,
        pedido,
        melhor.t.detalhe,
        melhor.t.custoRelogio,
        melhor.t.custoProdutivo,
        melhor.v,
        melhor.sub.subset
      );
      if (!ok) break;
      if (melhor.tipo === 'nova') resultado[melhor.v.id].push(melhor.vi);
      alocadoAlgum = true;
      partesAlocadas++;
      viagensComQuebra.push(melhor.vi);
      const usados = new Set(melhor.sub.subset);
      pendentes = pendentes.filter(pr => !usados.has(pr));
    }
    if (alocadoAlgum && (partesAlocadas > 1 || pendentes.length)) {
      const viMarcador = viagensComQuebra[0];
      if (viMarcador && !(viMarcador.quebras || []).some(q => q.pedidoId === pedido.id)) {
        viMarcador.quebras.push({
          pedidoId: pedido.id,
          cliente: pedido.cliente,
          partes: partesAlocadas,
          restanteM3: totalVolProdutos(pendentes),
        });
      }
    }
    return {
      alocadoAlgum,
      completo: pendentes.length === 0,
      restante: totalVolProdutos(pendentes),
      pendentes,
    };
  };
  // ── Fila priorizada ───────────────────────────────────────────────────────
  // 1) ID Petronas primeiro
  // 2) Dentro do mesmo terminal: janelas noturnas (fim ≤ 06:00) primeiro
  //    → garante que o menor truck compatível seja alocado para a noturna,
  //    e não "desperdiçado" em pedidos diurnos que aparecem antes na ordem alfa
  // 3) Agrupamento por terminal + cidade — pedidos da mesma cidade consecutivos
  // 4) Maior volume (dentro de cada grupo)
  const _fimJanela = (restricao) => {
    if (!restricao) return 9999;
    const m = restricao.match(/\d{2}:\d{2}-(\d{2}):(\d{2})/);
    return m ? parseInt(m[1]) * 60 + parseInt(m[2]) : 9999;
  };
  const _isNoturna = (restricao) => _fimJanela(restricao) <= 360; // termina até 06:00
  const fila = [...pedidos]
    .filter(p => !pedidosLocadosIds.has(p.id) && p.produtos?.length && totalVolPedido(p) > 0.001)
    .sort((a, b) => {
      if (b.identidadePetronas !== a.identidadePetronas)
        return (b.identidadePetronas ? 1 : 0) - (a.identidadePetronas ? 1 : 0);
      // Dentro do mesmo terminal: noturnas primeiro (evita big truck ser alocado em diurno)
      const termA = a.terminal || '', termB = b.terminal || '';
      if (termA === termB) {
        const nA = _isNoturna(a.restricao) ? 1 : 0;
        const nB = _isNoturna(b.restricao) ? 1 : 0;
        if (nA !== nB) return nB - nA;
      }
      const chaveA = termA + '||' + (a.cidade || '').toLowerCase().trim();
      const chaveB = termB + '||' + (b.cidade || '').toLowerCase().trim();
      if (chaveA < chaveB) return -1;
      if (chaveA > chaveB) return  1;
      return totalVolPedido(b) - totalVolPedido(a);
    });
  console.log('[Otimizador] fila:', fila.length, 'pedidos | grupos cidade+terminal:',
    [...new Set(fila.map(p => (p.terminal||'') + '|' + (p.cidade||'')))].length);
  // ── Loop principal ────────────────────────────────────────────────────────
  console.log('[Otimizador] fila:', fila.length, 'pedidos | veículos:', veiculos.length);
  for (const pedido of fila) {
    let alocado = false;
    // PASSO 1: Encaixar em viagem existente (mesmo terminal, cabem os produtos)
    outer1:
    for (const v of candidatos(pedido)) {
      for (const vi of resultado[v.id]) {
        const t = tentarEncaixe(pedido, v, vi);
        if (!t) continue;
        if (commitPedido(vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v)) {
          alocado = true;
          break outer1; // sucesso → para tudo
        }
        // commit falhou (compartimento) → tenta próxima viagem/veículo
      }
      if (alocado) break;
    }
    if (alocado) continue;
    // PASSO 1.5 (modo dedicado): Spreading proativo — divide produtos de um pedido
    // entre EXATAMENTE 2 dedicados ociosos quando nenhum truck sozinho resolve.
    // Guardas que evitam fragmentação excessiva (problema aprendido com Americana):
    //  a) Só ativa se o maior truck ocioso não consegue ≥70% do volume sozinho
    //     (quando existe um caminhão grande o suficiente, Passo 2 + 2ª viagem é melhor)
    //  b) Limita a exatamente 2 trucks — mais que isso produz partes minúsculas
    //  c) Cada parte deve ocupar ≥40% da capacidade do truck receptor
    if (modo === 'dedicado' && !alocado && pedido.produtos.length >= 2) {
      const ociososCompat = veiculos
        .filter(v =>
          !lockedTerminals.has(baseVeiculoLabel(v)) &&
          (v.disponibilidade || 'Disponível') !== 'Indisponível' &&
          (v.contrato || 'Dedicado') !== 'Spot' &&
          resultado[v.id].length === 0 &&
          veicOk(v, pedido))
        .sort((a, b) => a.capacidade - b.capacidade); // menores primeiro
      const volTotal15 = totalVolPedido(pedido);
      const maiorCapOcioso = ociososCompat.length ? Math.max(...ociososCompat.map(v => v.capacidade)) : 0;
      // Guarda (a): pula split se QUALQUER veículo compatível (ocioso ou não) consegue ≥85% do volume.
      // Isso evita dividir pedidos quando há um truck grande disponível em 2ª viagem.
      const maiorCapQualquer = veiculos
        .filter(v => !lockedTerminals.has(baseVeiculoLabel(v)) &&
          (v.disponibilidade || 'Disponível') !== 'Indisponível' &&
          veicOk(v, pedido))
        .reduce((mx, v) => Math.max(mx, v.capacidade), 0);
      const umTruckResolveMaior = maiorCapQualquer >= volTotal15 * 0.85;
      if (ociososCompat.length >= 2 && !umTruckResolveMaior) {
        const resSnap15 = JSON.parse(JSON.stringify(resultado));
        const ctSnap15  = JSON.parse(JSON.stringify(controleTempo));
        let prodRestantes = pedido.produtos.map((p, i) => ({ ...p, _si: i }));
        const partes15 = [];
        for (const vOcioso of ociososCompat) {
          if (!prodRestantes.length || partes15.length >= 2) break; // guarda (b)
          const prodsSem = prodRestantes.map(p => ({ produto: p.produto, volume: p.volume, ordemSAP: p.ordemSAP }));
          const sub = melhorSubconjuntoQueCabe(prodsSem, criarCompsDisp(vOcioso));
          if (!sub) continue;
          const tNova = tentarNovaViagem(pedido, vOcioso, sub.subset);
          if (!tNova) continue;
          partes15.push({ v: vOcioso, t: tNova, sub });
          const siUsados = new Set();
          for (const pr of sub.subset) {
            const idx = prodRestantes.findIndex(p => p.produto === pr.produto && !siUsados.has(p._si));
            if (idx >= 0) siUsados.add(prodRestantes[idx]._si);
          }
          prodRestantes = prodRestantes.filter(p => !siUsados.has(p._si));
        }
        // Guarda (c): cada parte deve ocupar ≥40% do truck receptor
        const ocupacaoOk = partes15.every(({ v, sub }) => sub.vol / v.capacidade >= 0.40);
        if (prodRestantes.length === 0 && partes15.length === 2 && ocupacaoOk) {
          let tudo15 = true;
          for (const { v, t, sub } of partes15) {
            if (!commitPedido(t.vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v, sub.subset)) {
              tudo15 = false; break;
            }
            resultado[v.id].push(t.vi);
          }
          if (tudo15) {
            alocado = true;
            console.log('[Dedicado] ⊞ Split proativo:', pedido.cliente,
              '→', partes15.map(p => `${p.v.placa} (${(p.sub.vol/p.v.capacidade*100).toFixed(0)}%)`).join(' + '));
          } else {
            Object.keys(resSnap15).forEach(k => resultado[k] = resSnap15[k]);
            Object.keys(ctSnap15 ).forEach(k => controleTempo[k] = ctSnap15[k]);
          }
        }
      }
    }
    if (alocado) continue; // ← Passo 1.5 alocou — não continuar para Passo 2/3
    // PASSO 2: Abrir nova viagem — pré-verifica compartimentos antes de criar
    {
      const cands2 = candidatos(pedido);
      const motivosSkip = [];
      console.log(`[P2] ${pedido.cliente} vol=${totalVolPedido(pedido).toFixed(1)}m³ candidatos=${cands2.length}`);
      for (const v of cands2) {
        const fitOk = podeFitar(pedido.produtos, criarCompsDisp(v));
        if (!fitOk) {
          const comps = criarCompsDisp(v).map(c=>c.cap).join('+');
          const prods = pedido.produtos.map(p=>`${p.produto}:${p.volume}m³`).join(', ');
          console.log(`  → ${v.placa} REJEITADO: compartimento incompatível. comps=[${comps}] prods=[${prods}]`);
          motivosSkip.push({ placa: v.placa, contrato: v.contrato||'Dedicado', motivo: 'compartimento incompatível' });
          continue;
        }
        const t = tentarNovaViagem(pedido, v);
        if (!t) {
          const ct = controleTempo[v.id];
          console.log(`  → ${v.placa} REJEITADO: timing/jornada. usadoMin=${ct.usadoMin} limiteMin=${ct.limiteMin} clockV=${clockV(v)}`);
          motivosSkip.push({ placa: v.placa, contrato: v.contrato||'Dedicado', motivo: 'timing/jornada' });
          continue;
        }
        if (commitPedido(t.vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v)) {
          resultado[v.id].push(t.vi);
          alocado = true;
          console.log(`  → ${v.placa} ALOCADO P2`);
          if ((v.contrato||'Dedicado') === 'Spot') {
            const dedSkip = motivosSkip.filter(m => m.contrato !== 'Spot');
            if (dedSkip.length)
              console.warn('[⚠ Spot > Dedicado]', pedido.cliente, '|', v.placa,
                '(Spot) alocado. Dedicados recusados antes:', dedSkip.map(m => `${m.placa}: ${m.motivo}`).join(', '));
          }
          break;
        }
      }
    }
    if (!alocado) {
      // PASSO 2.5: Tenta encaixe em viagem existente OU nova viagem com jornada excedida.
      // Pedidos que não cabem dentro da jornada normal são alocados com excesso registrado.
      // O alerta "Jornada excedida: +X h" é gerado em renderResultado a partir de usadoMin > limiteMin.
      const cands25 = candidatos(pedido);
      // 2.5a: encaixe em viagem existente com jornada excedida
      for (const v of cands25) {
        if (alocado) break;
        for (const vi of (resultado[v.id] || []).filter(x => !x._vazio && x.paradas?.length)) {
          const t = tentarEncaixe(pedido, v, vi, pedido.produtos, { permitirExcederJornada: true });
          if (!t) continue;
          if (commitPedido(vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v)) {
            alocado = true;
            console.log('[Otimizador] PASSO 2.5a (encaixe jornada excedida):', pedido.cliente, '→', v.placa);
            break;
          }
        }
      }
      // 2.5b: nova viagem com jornada excedida
      if (!alocado) {
        for (const v of cands25) {
          const fitOk = podeFitar(pedido.produtos, criarCompsDisp(v));
          if (!fitOk) continue;
          const t = tentarNovaViagem(pedido, v, pedido.produtos, { permitirExcederJornada: true });
          if (!t) continue;
          if (commitPedido(t.vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v)) {
            resultado[v.id].push(t.vi);
            alocado = true;
            console.log('[Otimizador] PASSO 2.5b (nova viagem jornada excedida):', pedido.cliente, '→', v.placa);
            break;
          }
        }
      }
    }
    if (!alocado) {
      // PASSO 3 (primário): quebra controlada — ÚLTIMO RECURSO.
      // Só chega aqui se nenhum veículo com capacidade total aceitou o pedido inteiro.
      // Modo padrão: 1 parte — evita consumir a frota em um único cliente.
      // Modo dedicado: máx 2 partes — evita fragmentar pedidos em 3+ trucks.
      const qb = tentarQuebraPedido(pedido, { maxPartes: modo === 'dedicado' ? 2 : 1 });
      if (qb.alocadoAlgum) {
        alocado = true;
        if (!qb.completo) {
          console.log('[Otimizador] ALOCAÇÃO PARCIAL (quebra controlada):', pedido.cliente,
            '| restante:', qb.restante.toFixed(1)+'m³');
        }
      }
    }
    if (!alocado) {
      const cands = candidatos(pedido);
      console.log('[Otimizador] NÃO ALOCADO:', pedido.cliente,
        '| vol:', totalVolPedido(pedido).toFixed(1)+'m³',
        '| terminal:', pedido.terminal,
        '| candidatos:', cands.length);
      cands.slice(0, 5).forEach(v => {
        const fitOk = podeFitar(pedido.produtos, criarCompsDisp(v));
        const nt    = fitOk ? tentarNovaViagem(pedido, v) : null;
        console.log('  →', v.placa, v.tipo, v.capacidade+'m³',
          `jornada:${controleTempo[v.id].usadoMin}/${controleTempo[v.id].limiteMin}min`,
          `comps:${fitOk?'OK':'FALHA'}`,
          nt ? `timing:OK prod:${nt.custoProdutivo}min` : 'timing:FALHA');
      });
    }
  }
  // PASSO 4 (secundário): com todos os pedidos já "semeados", tenta usar sobra
  // de capacidade/jornada para completar remanescentes de quebras.
  for (const pedido of fila) {
    const pendentes = produtosPendentesPedido(pedido);
    if (totalVolProdutos(pendentes) <= 0.0001) continue;
    // Primeiro tenta alocar os pendentes INTEIROS em nova viagem antes de quebrar mais
    let alocadoInteiro = false;
    const cands4 = candidatos(pedido);
    for (const v of cands4) {
      const fitOk = podeFitar(pendentes, criarCompsDisp(v));
      if (!fitOk) continue;
      const t = tentarNovaViagem({ ...pedido, produtos: pendentes }, v, pendentes, { permitirExcederJornada: true });
      if (!t) continue;
      if (commitPedido(t.vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v)) {
        resultado[v.id].push(t.vi);
        alocadoInteiro = true;
        console.log('[Otimizador] PASSO 4 (remanescente inteiro):', pedido.cliente, '→', v.placa);
        break;
      }
    }
    if (alocadoInteiro) continue;
    const qb2 = tentarQuebraPedido(pedido, { produtosIniciais: pendentes, permitirExcederJornada: true });
    if (qb2.alocadoAlgum && !qb2.completo) {
      console.log('[Otimizador] REMANESCENTE APÓS SOBRA:', pedido.cliente,
        '| restante:', qb2.restante.toFixed(1)+'m³');
    }
  }
  // ── PASSO 5: Consolidação — compacta frota eliminando veículos subutilizados ──
  // Tenta esvaziar cada veículo (do menos para o mais utilizado), realocando seus
  // pedidos nos demais. Permite estouro de jornada de até 15%.
  {
    const FATOR_OVERAGE = 1.05;
    const ordemConsolid = [...veiculos]
      .filter(v => resultado[v.id]?.length > 0 && !lockedTerminals.has(baseVeiculoLabel(v))
        // Modo dedicado: nunca esvaziar dedicados (custo fixo — manter todos ativos)
        && (modo === 'padrao' || (v.contrato || 'Dedicado') === 'Spot'))
      .sort((a, b) => {
        // Preferir esvaziar Spot antes de Dedicado (preservar uso da frota dedicada)
        const aSpot = (a.contrato || 'Dedicado') === 'Spot' ? 1 : 0;
        const bSpot = (b.contrato || 'Dedicado') === 'Spot' ? 1 : 0;
        if (aSpot !== bSpot) return bSpot - aSpot;
        const uA = controleTempo[a.id].usadoMin / (controleTempo[a.id].limiteMin || 1);
        const uB = controleTempo[b.id].usadoMin / (controleTempo[b.id].limiteMin || 1);
        return uA - uB;
      });
    for (const vFonte of ordemConsolid) {
      if (!resultado[vFonte.id]?.length) continue;
      // Pula veículos que têm quebras de pedido (alocação parcial) — lógica complexa
      const temQuebra = resultado[vFonte.id].some(vi => vi.quebras?.length > 0);
      if (temQuebra) continue;
      // Coleta pedidos únicos deste veículo na ordem das paradas
      const pedidosFonte = [];
      const vistos = new Set();
      for (const vi of resultado[vFonte.id]) {
        for (const pa of vi.paradas) {
          if (pa.pedido && !vistos.has(pa.pedido.id)) {
            vistos.add(pa.pedido.id);
            pedidosFonte.push(pa.pedido);
          }
        }
      }
      if (!pedidosFonte.length) continue;
      // Snapshot completo do estado antes de tentar
      const resSnap = JSON.parse(JSON.stringify(resultado));
      const ctSnap  = JSON.parse(JSON.stringify(controleTempo));
      // Eleva limite de jornada dos outros veículos em 15%
      veiculos.forEach(v => {
        if (v.id !== vFonte.id)
          controleTempo[v.id].limiteMin = ctSnap[v.id].limiteMin * FATOR_OVERAGE;
      });
      // Esvazia veículo fonte temporariamente
      resultado[vFonte.id]              = [];
      controleTempo[vFonte.id].usadoMin = 0;
      // Tenta realocar cada pedido em outros veículos
      let todosMigrados = true;
      for (const pedido of pedidosFonte) {
        let alocado = false;
        // Usa apenas os produtos ainda pendentes para este pedido
        // (exclui o que já está alocado em outros veículos)
        const pendProds = produtosPendentesPedido(pedido);
        if (totalVolProdutos(pendProds) <= 0.0001) {
          // Pedido já 100% coberto por outros veículos — não precisa migrar
          continue;
        }
        // Candidatos: exclui vFonte, prioriza Dedicado antes de Spot e depois os já ativos
        // Regra crítica: NUNCA mover trips de Dedicado para Spot
        const vFonteEhSpot = (vFonte.contrato || 'Dedicado') === 'Spot';
        const cands = veiculos
          .filter(v => {
            if (v.id === vFonte.id) return false;
            if (lockedTerminals.has(baseVeiculoLabel(v))) return false;
            if ((v.disponibilidade || 'Disponível') === 'Indisponível') return false;
            if (!veicOk(v, pedido)) return false;
            if (!vFonteEhSpot && (v.contrato || 'Dedicado') === 'Spot') return false;
            return true;
          })
          .sort((a, b) => {
            const aDedicado = (a.contrato || 'Dedicado') !== 'Spot' ? 1 : 0;
            const bDedicado = (b.contrato || 'Dedicado') !== 'Spot' ? 1 : 0;
            if (bDedicado !== aDedicado) return bDedicado - aDedicado;
            return (resultado[b.id].length > 0 ? 1 : 0) - (resultado[a.id].length > 0 ? 1 : 0);
          });
        // 1. Encaixar em viagem existente (só produtos pendentes)
        for (const v of cands) {
          for (const vi of resultado[v.id]) {
            const t = tentarEncaixe(pedido, v, vi, pendProds);
            if (t && commitPedido(vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v, pendProds)) {
              alocado = true; break;
            }
          }
          if (alocado) break;
        }
        // 2. Nova viagem (só produtos pendentes)
        if (!alocado) {
          for (const v of cands) {
            const t = tentarNovaViagem(pedido, v, pendProds);
            if (t && commitPedido(t.vi, pedido, t.detalhe, t.custoRelogio, t.custoProdutivo, v, pendProds)) {
              resultado[v.id].push(t.vi);
              alocado = true; break;
            }
          }
        }
        if (!alocado) { todosMigrados = false; break; }
      }
      // Restaura limites originais independente do resultado
      veiculos.forEach(v => { controleTempo[v.id].limiteMin = ctSnap[v.id].limiteMin; });
      if (todosMigrados) {
        console.log('[Consolidação] ✓ Esvaziado:', vFonte.placa,
          '(' + pedidosFonte.length + ' pedido(s) migrado(s))');
      } else {
        // Reverte todo o estado ao snapshot anterior
        Object.keys(resSnap).forEach(k => resultado[k] = resSnap[k]);
        Object.keys(ctSnap ).forEach(k => controleTempo[k] = ctSnap[k]);
        console.log('[Consolidação] ✗ Não foi possível esvaziar:', vFonte.placa);
      }
    }
  }
  // Itens ainda pendentes após tentativa secundária de sobra.
  const pendenciasFinais = [];
  for (const pedido of fila) {
    const pend = produtosPendentesPedido(pedido);
    for (const pr of pend) {
      pendenciasFinais.push({
        pedido,
        produto: pr.produto,
        restante: pr.volume,
        motivoNaoAlocado: 'quebra',
      });
    }
  }
  // ── Reconstrói contador a partir do histórico real antes de atribuir IDs ────
  // Garante que rodadas repetidas sem salvar não inflem a sequência.
  // Não bloqueia se a pasta não estiver disponível — apenas avisa via toast.
  try { await _reconstruirPetSeqDoHistorico({ obrigatorio: false }); } catch(e) {
    console.warn('[otimizar] Não foi possível reconstruir sequência do histórico:', e);
  }
  atribuirPetIds(resultado, baseDataEntrega || new Date());
  // Sugestões de quebra manual (só no modo dedicado): pedidos alocados sozinhos
  // em veículo com <55% de utilização temporal E há dedicado ocioso compatível.
  _sugestoesSplitDedicado = [];
  if (modo === 'dedicado') {
    const dedicOciosos = veiculos.filter(v =>
      !(resultado[v.id]?.some(vi => vi.paradas?.length > 0)) &&
      (v.contrato || 'Dedicado') !== 'Spot');
    if (dedicOciosos.length > 0) {
      for (const v of veiculos) {
        const viagens = resultado[v.id] || [];
        const util = (controleTempo[v.id]?.usadoMin || 0) / (controleTempo[v.id]?.limiteMin || 1);
        if (util > 0.55) continue;
        for (const vi of viagens) {
          if (!vi.paradas?.length) continue;
          const ids = new Set(vi.paradas.map(pa => pa.pedido?.id));
          if (ids.size !== 1) continue;
          const pedido = vi.paradas[0].pedido;
          if (!pedido || pedido.produtos.length < 2) continue;
          const ociosoOk = dedicOciosos.find(vO =>
            veiculoAtendeTerminal(vO, pedido.terminal) &&
            (!pedido.tiposCaminhao?.length || pedido.tiposCaminhao.includes(vO.tipo)) &&
            (!pedido.identidadePetronas || vO.identidadePetronas)
          );
          if (!ociosoOk) continue;
          _sugestoesSplitDedicado.push({
            cliente:    pedido.cliente,
            placa:      v.placa,
            terminal:   pedido.terminal,
            volTotal:   totalVolPedido(pedido),
            numProd:    pedido.produtos.length,
            utilPct:    Math.round(util * 100),
          });
        }
      }
    }
  }
  ultimoResultado                   = resultado;
  ultimoResultado._baseDataEntrega  = baseDataEntrega; // âncora de datas para o display
  ultimoControleTempo               = controleTempo;
  ultimoItensOtimizacao             = pendenciasFinais;
  resultadoOriginal     = JSON.parse(JSON.stringify(resultado));
  historicoManual       = [];
  atualizarBarraManual();
  estadoAtual = {
    pedidos:       JSON.parse(JSON.stringify(pedidos)),
    terminais:     JSON.parse(JSON.stringify(terminaisCad)),
    veiculos:      JSON.parse(JSON.stringify(veiculos)),
    resultado:     JSON.parse(JSON.stringify(resultado)),
    controleTempo: JSON.parse(JSON.stringify(controleTempo)),
  };
  document.querySelectorAll('.sel-roteirizacao').forEach(s => s.value = '');
  renderResultado(resultado, controleTempo);
  renderMapaGeral();
  renderTemplateOperacao();
  showTab('resultado');
  } catch(e) {
    console.error('[otimizar] Erro inesperado:', e);
    alert('Erro ao roteirizar: ' + (e.message || e));
  } finally {
    // Restaura lista completa de veículos (inclui indisponíveis)
    // para que a aba Veículos & Turnos continue mostrando todos
    veiculos = _veiculosTodos;
    todosBtn.forEach(b => { b.disabled = false; b.style.opacity = ''; });
    if (btn) btn.textContent = txtOriginal;
  }
}
function alocarItem(viagem, item, entrega, detalheParada=null, tempoAdicionalMin=0, compsSelecionados=null) {
  if (!viagem.terminalOrigem) viagem.terminalOrigem = item.terminal || '';
  let vol = entrega;
  const cc = (compsSelecionados && compsSelecionados.length)
    ? compsSelecionados
    : viagem.compsDisp.filter(c => c.disponivel>0 && (!c.produto||c.produto===item.produto));
  const alocacoesCpt = [];
  for (const comp of cc) {
    if (vol <= 0) break;
    const uso = Math.min(comp.disponivel, vol);
    if (uso <= 0) continue;
    comp.disponivel -= uso;
    vol -= uso;
    alocacoesCpt.push({
      cpt: comp.cpt || 0,
      produto: item.produto,
      volume: uso,
    });
  }
  const ex = viagem.paradas.find(p => p.pedido.id === item.pedidoId);
  if (ex) {
    ex.itens.push({
      produto:item.produto,
      volume:entrega,
      completo:entrega>=item.volume,
      ordemSAP: item.ordemSAP || '',
      alocacoesCpt,
    });
    ex.volumeTotal+=entrega;
  }
  else {
    const insertIdx = (detalheParada?.insertIdx !== undefined
      && detalheParada.insertIdx < viagem.paradas.length)
      ? detalheParada.insertIdx
      : null; // null = append ao final (comportamento padrão)
    if (insertIdx !== null) {
      // Inserção no meio da viagem: atualiza paradas adjacentes
      if (insertIdx === 0 && viagem.paradas.length > 0) {
        // A nova parada herda o carregamento; a antiga 1ª parada perde
        viagem.paradas[0].tempoCarregamentoMin = 0;
        viagem.paradas[0].origemDeslocamento   = 'Entrega anterior';
      }
      if (detalheParada._deslocNovaNextMin !== undefined) {
        // Atualiza o deslocamento da parada que ficará após a nova
        viagem.paradas[insertIdx].deslocCarregadoMin = detalheParada._deslocNovaNextMin;
      }
    } else if (detalheParada?.removeRetornoAnterior && viagem.paradas.length) {
      // Append ao final: zera retorno da última parada (ela não é mais a última)
      viagem.paradas[viagem.paradas.length - 1].deslocVazioMin = 0;
    }
    const novaParada = {
    pedido:item.pedido,
    itens:[{produto:item.produto,volume:entrega,completo:entrega>=item.volume,ordemSAP:item.ordemSAP||'',alocacoesCpt}],
    volumeTotal:entrega,
    cicloMin: detalheParada?.incrementoMin || 0,
    distanciaKm: detalheParada?.distanciaKm || 0,
    tempoCarregamentoMin: detalheParada?.tempoCarregamentoMin || 0,
    deslocCarregadoMin: detalheParada?.deslocCarregadoMin || 0,
    tempoDescargaMin: detalheParada?.tempoDescargaMin || 0,
    deslocVazioMin: detalheParada?.deslocVazioMin || 0,
    tempoEsperaRestricaoMin: detalheParada?.tempoEsperaRestricaoMin || 0,
    waitAfterLoadingMin: detalheParada?.waitAfterLoadingMin || 0,
    overnight: !!detalheParada?.overnight,
    origemDeslocamento: detalheParada?.origemDeslocamento || 'Terminal',
    };
    if (insertIdx !== null) {
      viagem.paradas.splice(insertIdx, 0, novaParada);
    } else {
      viagem.paradas.push(novaParada);
    }
  }
  viagem.tempoConsumidoMin = (viagem.tempoConsumidoMin || 0) + (tempoAdicionalMin || 0);
}
// Hue 25 (laranja) → 48 (amarelo) conforme pct sobe de 51 a 90
function _utilHue(pct) { return Math.round(25 + Math.max(0, Math.min(1, (pct - 51) / 39)) * 23); }
function utilTag(pct) { return pct > 90 ? 'tag-green' : pct > 50 ? 'tag-yellow' : 'tag-red'; }
// Cor inline para o badge de % (sobrepõe a classe no intervalo amarelo→laranja)
function utilTagStyle(pct) {
  if (pct > 90 || pct <= 50) return '';
  const h = _utilHue(pct);
  return `background:hsl(${h},80%,86%);color:hsl(${h},80%,22%);border-color:hsl(${h},70%,62%);`;
}
// Estilo dinâmico do route-header baseado na % de ocupação
function utilEstilo(pct) {
  if (pct > 90) return 'border-left:5px solid #5E9A18;background:linear-gradient(to right,#86EFAC,#4ADE80);';
  if (pct > 50) {
    const h = _utilHue(pct);
    return `border-left:5px solid hsl(${h},85%,45%);background:linear-gradient(to right,hsl(${h},90%,68%),hsl(${h},90%,56%));`;
  }
  return 'border-left:5px solid #DC2626;background:linear-gradient(to right,#FCA5A5,#F87171);';
}
// ── Exportação Herrlog ────────────────────────────────────────────────────────
async function exportarHrrlog() {
  if (!ultimoResultado || !veiculos.length) { alert('Execute a otimização primeiro.'); return; }
  if (typeof XLSX === 'undefined') { alert('SheetJS não carregado.'); return; }
  const rowsViagens = [];
  const rowsEventos = [];
  // Formata minuto absoluto + data base → "YYYY-MM-DD HH:MM:00"
  const fmtHrr = (baseDate, absMin) => {
    const m = Math.round(absMin); // arredonda para evitar casas decimais
    const d = new Date(baseDate.getTime() + Math.floor(m / 1440) * 86400000);
    const yy = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    const dy = String(d.getDate()).padStart(2, '0');
    const hh = String(Math.floor((m % 1440) / 60)).padStart(2, '0');
    const mn = String(m % 60).padStart(2, '0');
    return `${yy}-${mo}-${dy} ${hh}:${mn}:00`;
  };
  veiculos.forEach(v => {
    const todasViagens = (ultimoResultado[v.id] || []).filter(vi => !vi._vazio && (vi.paradas || []).length);
    if (!todasViagens.length) return;
    const _jIniRawCt = parseHoraMin(v.jornadaInicio || '06:00');
    let jIniMin = isNaN(_jIniRawCt) ? 360 : _jIniRawCt;
    if (v._horarioDisponivelAPartirDe) { const _dmCt = parseHoraMin(v._horarioDisponivelAPartirDe); if (!isNaN(_dmCt) && _dmCt > jIniMin) jIniMin = _dmCt; }
    const jornadaInicioMin = jIniMin; // alias para compatibilidade
    todasViagens.forEach((vi, idx) => {
      // Data base da viagem (mesma lógica do template operação)
      const _fp = vi.paradas[0];
      const _dl = _fp?.pedido?.dataEntregaLogistica;
      let baseDate = new Date();
      if (_dl) {
        const pts = _dl.split('/');
        if (pts.length >= 3) {
          const d = new Date(parseInt(pts[2]), parseInt(pts[1]) - 1, parseInt(pts[0]));
          if (_fp.overnight) d.setDate(d.getDate() - 1);
          baseDate = d;
        }
      }
      // Clock acumulado até esta viagem
      let relogioMin = inicioViagemAbsMin(todasViagens, idx, jIniMin, v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1);
      relogioMin += vi.esperaTerminalMin || 0;
      const temOverrideCarga = vi.horarioCargaManualMin !== undefined && !isNaN(vi.horarioCargaManualMin);
      if (temOverrideCarga) {
        const clockAnterior = idx > 0
          ? inicioViagemAbsMin(todasViagens, idx, jIniMin, v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1)
          : jIniMin;
        const baseDay = Math.floor(clockAnterior / 1440) * 1440;
        let alvo = baseDay + vi.horarioCargaManualMin;
        if (alvo < clockAnterior - 0.001) alvo += 1440;
        relogioMin = alvo;
      }
      // Início de carga: usa o valor já calculado pelo render da tela se disponível,
      // senão usa relogioMin puro (sem atrasoP0 — tempoEsperaRestricaoMin é espera no
      // cliente, não atraso de carregamento no terminal).
      const p0 = vi.paradas[0];
      const inicioCargaMin = temOverrideCarga ? relogioMin : (vi._inicioCargaMin ?? relogioMin);
      // Terminal desta viagem (fallback ao terminal do pedido se terminalOrigem ausente)
      const _termPedidoHrr = vi.paradas?.find(p => p.pedido?.terminal)?.pedido?.terminal || '';
      const termNome = vi.terminalOrigem || _termPedidoHrr || v.terminal || '';
      const term = terminaisCad.find(t => t.nome === termNome);
      const codigoTerminal = term?.empresaLocalExpedicao || term?.nome || termNome;
      // Produtos únicos (nome resumido)
      const produtos = [...new Set(
        vi.paradas.flatMap(p => (p.itens || []).map(it => resumoProduto(it.produto)))
      )].join('/');
      // Volume total em litros
      const volLitros = Math.round(vi.paradas.reduce((s, p) => s + (p.volumeTotal || 0), 0) * 1000);
      // KM total da viagem — usa o km ajustado manualmente no mapa da viagem
      // (quando o usuário desviou a rota), senão cai no cálculo original do otimizador
      const kmTotal = Math.round(
        vi._kmAjustado != null
          ? vi._kmAjustado
          : vi.paradas.reduce((s, p) => s + (p.distanciaKm || 0), 0)
      );
      const codViagem = vi.petId || `VIA${v.placa}${idx + 1}`;
      // ── Linha Viagens ────────────────────────────────────────────────────────
      rowsViagens.push({
        'Código da Viagem*': codViagem,
        'Status': 'Programado',
        'Início Previsto*': fmtHrr(baseDate, inicioCargaMin),
        'Placa*': v.placa || '',
        'Trailers*': v.implemento || v.placa || '',
        'Tipo de atendimento': 'CIF',
        'KM': kmTotal || '',
        'Produto': produtos,
        'Observação': '',
        'Valor (R$)': '',
        'Total Planejado (Kg)': volLitros,
        'Total Carregado (KG)': volLitros,
        'Peso total (KG)': '',
        'Prioridade': 'Normal',
        'Ordem de Retirada': '',
      });
      // ── Evento: Carregamento (terminal) ──────────────────────────────────────
      rowsEventos.push({
        'Código da Viagem*': codViagem,
        'Codigo do Cliente*': codigoTerminal,
        'Tipo Evento*': 'Carregamento',
        'Data Planejada*': fmtHrr(baseDate, inicioCargaMin),
      });
      // ── Eventos: Descarga (cada parada) ──────────────────────────────────────
      let clock = inicioCargaMin;
      vi.paradas.forEach((p, idxP) => {
        const espOrig = p.tempoEsperaRestricaoMin || 0;
        const wal = p.overnight ? (p.waitAfterLoadingMin || 0) : 0;
        const atraso = (!temOverrideCarga && !p.overnight && idxP === 0 && espOrig > 0 && (p.tempoCarregamentoMin || 0) > 0) ? espOrig : 0;
        const fimCarga = clock + atraso + (p.tempoCarregamentoMin || 0);
        const chegada  = fimCarga + wal + (p.deslocCarregadoMin || 0);
        const espVis   = p.overnight ? 0 : (espOrig - atraso);
        const fimDesc  = chegada + espVis + (p.tempoDescargaMin || 0);
        clock = fimDesc + (p.deslocVazioMin || 0);
        rowsEventos.push({
          'Código da Viagem*': codViagem,
          'Codigo do Cliente*': p.pedido?.codigoSAP || '',
          'Tipo Evento*': 'Descarga',
          'Data Planejada*': fmtHrr(baseDate, chegada),
        });
      });
    });
  });
  const hV = ['Código da Viagem*','Status','Início Previsto*','Placa*','Trailers*','Tipo de atendimento','KM','Produto','Observação','Valor (R$)','Total Planejado (Kg)','Total Carregado (KG)','Peso total (KG)','Prioridade','Ordem de Retirada'];
  const hE = ['Código da Viagem*','Codigo do Cliente*','Tipo Evento*','Data Planejada*'];
  const wb = XLSX.utils.book_new();
  const wsV = XLSX.utils.json_to_sheet(rowsViagens, { header: hV });
  const wsE = XLSX.utils.json_to_sheet(rowsEventos, { header: hE });
  wsV['!cols'] = [22,14,20,12,12,18,8,40,16,12,18,18,14,12,16].map(w => ({ wch: w }));
  wsE['!cols'] = [22,20,22,20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsV, 'Viagens');
  XLSX.utils.book_append_sheet(wb, wsE, 'Eventos de Viagem');
  const hoje = new Date();
  const fname = `Herrlog_${String(hoje.getDate()).padStart(2,'0')}${String(hoje.getMonth()+1).padStart(2,'0')}${hoje.getFullYear()}_${String(hoje.getHours()).padStart(2,'0')}${String(hoje.getMinutes()).padStart(2,'0')}.xlsx`;
  XLSX.writeFile(wb, fname);
  // Salva automaticamente no histórico do dashboard após exportar
  try {
    await salvarNoHistorico(true);
    showToast('Herrlog exportado e salvo no histórico ✅', true);
  } catch(e) {
    showToast('Herrlog exportado. Falha ao salvar histórico: ' + e.message, false);
  }
}
// Cor da placa-badge baseada na % de ocupação
function utilPlacaBg(pct) {
  if (pct > 90) return '#1A2A0D';
  if (pct > 50) return `hsl(${_utilHue(pct)},75%,25%)`;
  return               '#991B1B';
}
function utilViagBg(pct) {
  if (pct > 90) return 'background:#DCFCE7;border-left:3px solid #4ADE80;';
  if (pct > 50) {
    const h = _utilHue(pct);
    return `background:hsl(${h},85%,87%);border-left:3px solid hsl(${h},80%,60%);`;
  }
  return 'background:#FEE2E2;border-left:3px solid #FCA5A5;';
}
function utilStopBg(pct) {
  if (pct > 90) return 'background:#F0FFF4;';
  if (pct > 50) return `background:hsl(${_utilHue(pct)},90%,94%);`;
  return               'background:#FFF5F5;';
}
function renderTemplateOperacao() {
  const el = document.getElementById('operacao-content');
  if (!el) return;
  if (!ultimoResultado) {
    el.innerHTML = '<div class="empty">Execute a otimização para gerar o template operacional.</div>';
    return;
  }
  const filtroCliente = valId('f-op-cliente');
  const filtroCidade = valId('f-op-cidade');
  const filtroTerminal = valId('f-op-terminal');
  const filtroPlaca = valId('f-op-placa');
  const filtroTransp = valId('f-op-transp');
  const temFiltroParada = !!(txt(filtroCliente) || txt(filtroCidade) || txt(filtroTerminal));
  const dataAgora = new Date();
  const dataBr = `${String(dataAgora.getDate()).padStart(2,'0')}/${String(dataAgora.getMonth()+1).padStart(2,'0')}/${dataAgora.getFullYear()}`;
  const blocos = [];
  veiculos.forEach(v => {
    const viagens = (ultimoResultado[v.id] || []).filter(vi => !vi._vazio && (vi.paradas || []).length);
    viagens.forEach((vi, idx) => {
      const base = vi.terminalOrigem || v.terminal || cidadeBaseVeiculo(v) || '-';
      const paradasOrig = vi.paradas || [];
      const paradasFiltradas = paradasOrig.filter(p =>
        containsFiltro(p.pedido?.cliente, filtroCliente) &&
        containsFiltro(p.pedido?.cidade, filtroCidade) &&
        containsFiltro(p.pedido?.terminal, filtroTerminal)
      );
      const paradasUsar = temFiltroParada ? paradasFiltradas : paradasOrig;
      const passaVeiculo =
        containsFiltro(v.placa, filtroPlaca) &&
        containsFiltro(v.transportadora, filtroTransp) &&
        (!txt(filtroTerminal) ||
          containsFiltro(base, filtroTerminal) ||
          containsFiltro(v.terminal, filtroTerminal) ||
          paradasFiltradas.length > 0);
      if (!passaVeiculo) return;
      if (!paradasUsar.length) return;
      blocos.push({ v, viOriginal: vi, paradasUsar, idx, viagensVeiculo: viagens });
    });
  });
  if (!blocos.length) {
    el.innerHTML = '<div class="empty">Nenhum template operacional para os filtros informados.</div>';
    return;
  }
  el.innerHTML = blocos.map(({ v, viOriginal, paradasUsar, idx, viagensVeiculo }) => {
    // terminalOrigem pode estar vazio em viagens criadas manualmente → fallback ao terminal do pedido
    const _termPedido = viOriginal.paradas?.find(p => p.pedido?.terminal)?.pedido?.terminal || '';
    const base = viOriginal.terminalOrigem || _termPedido || v.terminal || cidadeBaseVeiculo(v) || '-';
    const cia = distribuidoraDoTerminal(viOriginal.terminalOrigem || v.terminal) || '-';
    const capLitros = Math.round((v.capacidade || 0) * 1000);
    const volViagemM3 = (paradasUsar || []).reduce((s, p) => s + (p.volumeTotal || 0), 0);
    const ocupPct = (v.capacidade || 0) > 0 ? Math.round((volViagemM3 / v.capacidade) * 100) : 0;
    const ocupClasse = ocupPct >= 90 ? 'op-ocup-verde' : (ocupPct < 50 ? 'op-ocup-vermelho' : 'op-ocup-amarelo');
    const _jIniRawRo = parseHoraMin(v.jornadaInicio || '06:00');
    const jornadaInicioMin = (() => { const base = isNaN(_jIniRawRo)?360:_jIniRawRo; if(v._horarioDisponivelAPartirDe){const dm=parseHoraMin(v._horarioDisponivelAPartirDe);if(!isNaN(dm)&&dm>base)return dm;} return base; })();
    let relogioMin = inicioViagemAbsMin(
      viagensVeiculo || [],
      idx,
      isNaN(jornadaInicioMin) ? 360 : jornadaInicioMin,
      v.tempoPerdidoMin || 0,
      doisTurnos(v) ? 2 : 1
    );
    // Data base para este ciclo — tenta a 1ª parada, depois qualquer parada da viagem,
    // depois qualquer viagem do veículo, para não perder a data quando o pedido não a tem.
    const _op_fp = viOriginal.paradas[0];
    let opBaseDate = null;
    const _parseDateBr = (dl, overnight) => {
      if (!dl) return null;
      const pts = dl.split('/');
      if (pts.length < 3) return null;
      const d = new Date(parseInt(pts[2]), parseInt(pts[1]) - 1, parseInt(pts[0]));
      if (overnight) d.setDate(d.getDate() - 1);
      return d;
    };
    if (ultimoResultado._baseDataEntrega) {
      const _bd = new Date(ultimoResultado._baseDataEntrega);
      if (!isNaN(_bd.getTime())) opBaseDate = _bd;
    }
    if (!opBaseDate) {
      opBaseDate = _parseDateBr(_op_fp?.pedido?.dataEntregaLogistica, _op_fp?.overnight);
      if (!opBaseDate) {
        for (const p of (viOriginal.paradas || [])) {
          opBaseDate = _parseDateBr(p.pedido?.dataEntregaLogistica, p.overnight);
          if (opBaseDate) break;
        }
      }
      if (!opBaseDate) {
        for (const vi of (viagensVeiculo || [])) {
          for (const p of (vi.paradas || [])) {
            opBaseDate = _parseDateBr(p.pedido?.dataEntregaLogistica, p.overnight);
            if (opBaseDate) break;
          }
          if (opBaseDate) break;
        }
      }
    }
    const fmtOpDT = (absMin) => {
      if (!opBaseDate) return fmtHora(absMin);
      const d = new Date(opBaseDate.getTime() + Math.floor(absMin / 1440) * 86400000);
      return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')} ${fmtHora(absMin)}`;
    };
    // Aplica espera terminal (overnight)
    const _op_espTerm = viOriginal.esperaTerminalMin || 0;
    relogioMin += _op_espTerm;
    const _opTemOverride = viOriginal.horarioCargaManualMin !== undefined;
    // Se o usuário definiu horário manual de carga, aplica ao relogioMin do bloco transportador
    // para que "Carregamento" no resumo reflita o horário editado, não o calculado automaticamente.
    if (_opTemOverride && !isNaN(viOriginal.horarioCargaManualMin)) {
      // O horarioCargaManualMin é sempre relativo ao dia base (dia 0 = dia de carregamento).
      // Não usamos Math.floor(relogioMin/1440) para evitar jogar para o dia seguinte
      // quando o relogioMin acumulado já passou de 1440 (overnight anterior).
      // O dia base é sempre 0 (jornada começa no dia do carregamento).
      const _alvoManual = viOriginal.horarioCargaManualMin;
      relogioMin = _alvoManual;
    }
    if (!_opTemOverride && Number.isFinite(viOriginal._inicioCargaMin)) {
      relogioMin = viOriginal._inicioCargaMin;
    }
    let inicioCargaCicloMin = relogioMin;
    let retornoBaseCicloMin = relogioMin;
    (viOriginal.paradas || []).forEach((p, idxParada) => {
      const esperaOriginalMin = p.tempoEsperaRestricaoMin || 0;
      const waitAfterLoad    = p.overnight ? (p.waitAfterLoadingMin || 0) : 0;
      const atrasoCargaMin   = (!p.overnight && idxParada === 0 && esperaOriginalMin > 0 && (p.tempoCarregamentoMin || 0) > 0 && !_opTemOverride)
        ? esperaOriginalMin : 0;
      const inicioCargaParadaMin = relogioMin + atrasoCargaMin;
      if (idxParada === 0) inicioCargaCicloMin = inicioCargaParadaMin;
      const fimCargaMin        = inicioCargaParadaMin + (p.tempoCarregamentoMin || 0);
      const chegadaEntregaMin  = fimCargaMin + waitAfterLoad + (p.deslocCarregadoMin || 0);
      const esperaVisivelMin   = p.overnight ? 0 : (esperaOriginalMin - atrasoCargaMin);
      const inicioDescargaMin  = chegadaEntregaMin + esperaVisivelMin;
      const fimDescargaMin     = inicioDescargaMin + (p.tempoDescargaMin || 0);
      const retornoTerminalMin = fimDescargaMin + (p.deslocVazioMin || 0);
      relogioMin = retornoTerminalMin;
      retornoBaseCicloMin = retornoTerminalMin;
    });
    const paradasVisiveisRef = new Set(paradasUsar || []);
    const cptTotal = (v.compartimentos || []).length;
    let cptAtual = cptTotal;
    const linhasDados = [];
    // Re-simula alocação de compartimentos respeitando as regras operacionais:
    //   CARREGAMENTO: sempre do C1 em diante (ascendente) — os últimos ficam vazios.
    //   DESCARGA:     do último carregado para o C1 (descendente).
    //   MULTI-STOP:   processa paradas em ORDEM REVERSA (última entrega primeiro).
    //                 Última entrega ocupa C1, C2... | Primeira entrega ocupa os CPTs mais altos.
    //                 Template (CPT desc) exibe as entregas na ordem correta da rota.
    (() => {
      // Ordena ascendente: C1 primeiro. Para capacidades iguais o sort estável preserva a
      // ordem crescente de CPT → selecionarCompartimentos (sort interno por disponivel DESC,
      // estável) preferirá CPTs menores quando empatar em capacidade.
      const compsSimul = criarCompsDisp(v)
        .sort((a, b) => (a.cap || 0) - (b.cap || 0) || a.cpt - b.cpt);
      // Processa paradas em ORDEM REVERSA: última entrega consome os menores CPTs,
      // primeira entrega recebe os CPTs mais altos →
      // template (CPT desc) exibe as entregas na sequência correta da rota.
      const paradasRev = [...(viOriginal.paradas || [])].reverse();
      for (const p of paradasRev) {
        for (const it of (p.itens || [])) {
          const el = compsSimul.filter(c => c.disponivel > 0 && compElegivelProduto(c, it.produto));
          const aloc = selecionarCompartimentos(el, it.volume, 'exact')
                    || selecionarCompartimentos(el, it.volume, 'maxLE');
          if (!aloc) continue;
          it.alocacoesCpt = [];
          let restante = it.volume;
          for (const c of aloc) {
            const uso = Math.min(c.disponivel, restante);
            if (uso <= 0) continue;
            c.disponivel -= uso;
            restante -= uso;
            it.alocacoesCpt.push({ cpt: c.cpt, produto: it.produto, volume: uso });
          }
        }
      }
    })();
    (viOriginal.paradas || []).forEach(p => {
      const cidadeEntrega = (p.pedido?.cidade || '').toString().trim();
      const postoCidade = cidadeEntrega ? `${p.pedido?.cliente || ''} (${cidadeEntrega.toUpperCase()})` : (p.pedido?.cliente || '');
      // Mapa produto → ordemSAP individual (vem de xlsxMapPedidosLiberadosRows)
      const produtoOrdemMap = {};
      (p.pedido?.produtos || []).forEach(pr => {
        if (pr.ordemSAP) produtoOrdemMap[pr.produto] = pr.ordemSAP;
      });
      const itens = (p.itens || []).length ? p.itens : [{ produto: '-', volume: p.volumeTotal || 0 }];
      const _prodOcorrencia = {}; // rastreia quantas vezes cada produto apareceu nos itens
      const linhasParada = itens.flatMap(it => {
        const prodKey = it.produto || '';
        _prodOcorrencia[prodKey] = (_prodOcorrencia[prodKey] || 0) + 1;
        // ordemSAP: (1) gravado no item, (2) Nª ocorrência do produto em pedido.produtos, (3) mapa prod→ordem, (4) ordem única, (5) vazio
        let ordemSAPIt = it.ordemSAP || '';
        if (!ordemSAPIt) {
          let occ = 0;
          for (const pr of (p.pedido?.produtos || [])) {
            if (pr.produto === prodKey) {
              occ++;
              if (occ === _prodOcorrencia[prodKey]) { ordemSAPIt = pr.ordemSAP || ''; break; }
            }
          }
        }
        if (!ordemSAPIt) ordemSAPIt = produtoOrdemMap[prodKey] || (p.pedido?.ordens?.length === 1 ? p.pedido.ordens[0] : '') || '';
        const alocacoes = (it.alocacoesCpt || []).filter(a => (a.volume || 0) > 0);
        if (alocacoes.length) {
          return alocacoes.map(a => ({
            postoCidade, ordemSAP: ordemSAPIt,
            produto: a.produto || prodKey,
            volumeL: Math.round((a.volume || 0) * 1000),
            cptOrig: a.cpt || '',
          }));
        }
        return [{
          postoCidade, ordemSAP: ordemSAPIt,
          produto: prodKey,
          volumeL: Math.round((it.volume || 0) * 1000),
          cptOrig: '',
        }];
      });
      linhasParada.forEach(ld => {
        // Prefere o compartimento real da alocação; só usa countdown quando não há info de compartimento.
        const cptAtrib = ld.cptOrig || (cptAtual > 0 ? cptAtual-- : '');
        if (!paradasVisiveisRef.has(p)) return;
        linhasDados.push({ ...ld, cpt: cptAtrib });
      });
    });
    linhasDados.sort((a, b) => (parseInt(b.cpt) || 0) - (parseInt(a.cpt) || 0));
    const linhas = linhasDados.map(ld => `
      <tr>
        <td>${ld.ordemSAP || ''}</td>
        <td>${cia}</td>
        <td>${base}</td>
        <td>${ld.postoCidade}</td>
        <td style="text-align:center;font-weight:700;">${ld.cpt || ''}</td>
        <td>${ld.produto || ''}</td>
        <td style="text-align:right;white-space:nowrap;">${ld.volumeL.toLocaleString('pt-BR')}</td>
      </tr>`).join('');
    const linhasFinal = linhas;
    const _nexta_svg = `<svg style="height:20px;width:auto;" viewBox="0 0 242 45" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M149.631 0.870612H138.515L127.579 14.6805L116.586 0.870612H105.356L121.715 21.6422L104.148 43.9406H114.999L127.078 28.421L139.156 43.7829L139.279 43.9406H150.833L133.011 21.5161L149.631 0.870612ZM67.2515 43.9437H97.073V35.4805H65.8352V26.5725H86.6225V18.1756H65.8352V9.26758H97.073V0.870612H56.9272V43.9437H67.2515ZM196.417 0.870612H155.382V9.26758H171.479V43.9406H180.324V9.27074H196.42V0.870612H196.417ZM224.4 0.870612H211.811L195.01 43.9437H204.426L208.498 33.1273H227.523L231.658 43.9437H241.2L224.4 0.870612ZM224.453 24.7304H211.618L217.892 8.75657L224.457 24.7304H224.453ZM36.8748 0.870612V34.8938C36.8748 35.2786 36.5719 35.6666 36.1019 35.6666C35.9127 35.6666 35.6729 35.6004 35.4963 35.3764L17.8759 4.82621C17.0305 3.35942 15.8508 2.0882 14.3777 1.25545C13.0718 0.517321 11.4504 1.75258e-06 9.54517 1.75258e-06C4.25526 -0.00315263 0 4.25211 0 9.86061V43.9437H8.93006V9.92054C8.93006 9.53571 9.23288 9.14772 9.70289 9.14772C9.89215 9.14772 10.1319 9.21396 10.3085 9.43792L26.5599 37.9031C27.2097 39.0387 27.8595 40.0828 28.503 40.9755C29.7774 42.7514 32.4586 44.8207 36.2691 44.8207C41.559 44.8207 45.8111 40.5654 45.8111 34.9569V0.870612H36.8811H36.8748Z" fill="#2D6A1B"/></svg>`;
    return `
      <div class="op-bloco" data-bloco-id="bloco-${v.placa}-${idx}">
        <div class="op-head">
          <!-- Barra superior: logo + ID da viagem + botões de exportação -->
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;padding-bottom:9px;border-bottom:1.5px solid #C8E0B0;">
            <div style="display:flex;align-items:center;gap:10px;">
              ${_nexta_svg}
              <span style="font-size:9px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#4A6535;opacity:.7;">Roteirização Outbound</span>
            </div>
            <div style="display:flex;align-items:center;gap:8px;">
              <div style="font-size:14px;font-weight:800;letter-spacing:.06em;color:#1A3A0A;">
                Viagem ${idx + 1}${viOriginal.petId ? ` &nbsp;·&nbsp; <span style="font-family:var(--font-cond);letter-spacing:.08em;">${viOriginal.petId}</span>` : ''}
              </div>
              <button onclick="exportarBlocoPDF('bloco-${v.placa}-${idx}', '${(viOriginal.petId || `V${idx+1}_${v.placa}`).replace(/'/g,'')}', event)"
                title="Exportar esta programação como PDF"
                style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:12px;border:1px solid rgba(239,68,68,0.35);background:rgba(239,68,68,0.07);color:#DC2626;font-size:10px;font-weight:700;cursor:pointer;font-family:var(--font-cond);letter-spacing:.05em;white-space:nowrap;transition:all .15s;"
                onmouseover="this.style.background='rgba(239,68,68,0.15)'" onmouseout="this.style.background='rgba(239,68,68,0.07)'">
                📄 PDF
              </button>
              <button onclick="exportarBlocoPNG('bloco-${v.placa}-${idx}', '${(viOriginal.petId || `V${idx+1}_${v.placa}`).replace(/'/g,'')}', event)"
                title="Exportar esta programação como imagem PNG"
                style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:12px;border:1px solid rgba(79,70,229,0.35);background:rgba(79,70,229,0.07);color:#4338CA;font-size:10px;font-weight:700;cursor:pointer;font-family:var(--font-cond);letter-spacing:.05em;white-space:nowrap;transition:all .15s;"
                onmouseover="this.style.background='rgba(79,70,229,0.15)'" onmouseout="this.style.background='rgba(79,70,229,0.07)'">
                🖼 PNG
              </button>
            </div>
          </div>
          <!-- Grid de dados -->
          <div style="display:grid;grid-template-columns:repeat(4,auto);gap:5px 16px;font-size:12px;align-items:baseline;">
            <div class="op-head-lbl">Carregamento</div><div><b>${fmtOpDT(inicioCargaCicloMin)}</b></div>
            <div class="op-head-lbl">Transportador</div><div>${v.transportadora || '-'}</div>
            <div class="op-head-lbl">Retorno base</div><div>${fmtOpDT(retornoBaseCicloMin)}</div>
            <div class="op-head-lbl">Cavalo</div><div>${v.placa || '-'}</div>
            <div class="op-head-lbl">Motorista</div><div>${motoristaDaViagem(v, inicioCargaCicloMin, viOriginal) || '—'}</div>
            ${v.implemento ? `<div class="op-head-lbl">Implemento</div><div>${v.implemento}</div>` : ''}
            <div class="op-head-lbl">Cap. (L)</div><div>${capLitros.toLocaleString('pt-BR')}</div>
            <div class="op-head-lbl">Base</div><div>${base}</div>
            <div class="op-head-lbl">Ocupação</div><div><span class="op-ocup ${ocupClasse}">${ocupPct}%</span><span style="margin-left:5px;font-size:11px;">(${volViagemM3.toFixed(1)} / ${(v.capacidade||0).toFixed(1)} m³)</span></div>
          </div>
          ${(() => {
            const paradasComJanela = (viOriginal.paradas || []).filter(p => {
              const obs = encontrarClienteDoPedido(p.pedido)?.observacoes || p.pedido?.observacoes || '';
              return p.pedido?.restricao || obs;
            });
            if (!paradasComJanela.length) return '';
            const linhasJanela = paradasComJanela.map(p => {
              const dataEnt = p.pedido.dataEntregaLogistica ? `${p.pedido.dataEntregaLogistica} ` : '';
              const obs = encontrarClienteDoPedido(p.pedido)?.observacoes || p.pedido?.observacoes || '';
              return `<div style="display:flex;gap:8px;align-items:baseline;flex-wrap:wrap;">
                <span style="font-weight:600;font-size:12px;">${p.pedido.cliente}</span>
                ${p.pedido.restricao ? `<span style="font-size:11px;color:#92400E;background:#FEF3C7;border:1px solid #FCD34D;border-radius:4px;padding:1px 7px;">${dataEnt}${p.pedido.restricao}</span>` : ''}
                ${obs ? `<span style="font-size:11px;color:#4A6535;font-style:italic;">${obs}</span>` : ''}
              </div>`;
            }).join('');
            return `<div style="padding:8px 14px;border-top:1px solid var(--border);background:#FFFDF0;">
              <div style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#92400E;margin-bottom:5px;">Janelas de recebimento</div>
              <div style="display:flex;flex-direction:column;gap:4px;">${linhasJanela}</div>
            </div>`;
          })()}
        </div>
        <table class="op-table">
          <colgroup>
            <col style="width:90px"/>
            <col style="width:50px"/>
            <col style="width:130px"/>
            <col/>
            <col style="width:44px"/>
            <col style="width:220px"/>
            <col style="width:76px"/>
          </colgroup>
          <thead>
            <tr>
              <th>No. Ordem SAP</th>
              <th>CIA</th>
              <th>Base</th>
              <th>Posto de entrega</th>
              <th style="text-align:center;">CPT</th>
              <th>Cód. Produto — Descrição</th>
              <th style="text-align:right;white-space:nowrap;">Volume (L)</th>
            </tr>
          </thead>
          <tbody>
            ${linhasFinal || '<tr><td colspan="6" style="color:#8AAA70;">Sem itens</td></tr>'}
          </tbody>
        </table>
      </div>
    `;
  }).join('');
}
// ── Ajuste manual (drag-and-drop de viagens entre placas) ────────────────────
// Marca veículo como vazio usando viagem fantasma dentro do próprio ultimoResultado
// Assim o estado fica auto-contido e undo/redo funciona sem tracking externo
function marcarVeiculoVazio(veiculoId) {
  if (!ultimoResultado[veiculoId]) ultimoResultado[veiculoId] = [];
  if (ultimoResultado[veiculoId].length === 0)
    ultimoResultado[veiculoId] = [{ _vazio: true, paradas: [], quebras: [] }];
}
function limparFantasma(veiculoId) {
  if (!ultimoResultado[veiculoId]) return;
  ultimoResultado[veiculoId] = ultimoResultado[veiculoId].filter(vi => !vi._vazio);
}
function atualizarBarraManual() {
  const bar  = document.getElementById('ajuste-manual-bar');
  const info = document.getElementById('ajuste-manual-info');
  const btn  = document.getElementById('btn-desfazer-manual');
  if (!bar) return;
  const temHistorico = historicoManual.length > 0;
  bar.style.display = temHistorico ? 'flex' : 'none';
  if (info) info.textContent = temHistorico ? `${historicoManual.length} movimentação(ões) pendente(s)` : '';
  if (btn)  btn.disabled = !temHistorico;
}
function iniciarDragViagem(veiculoId, tripIndex, el) {
  dragInfo = { tipo: 'viagem', veiculoId, tripIndex };
  setTimeout(() => el.classList.add('viagem-draggando'), 0);
}
function fimDragViagem(el) {
  el.classList.remove('viagem-draggando');
  dragInfo = null;
  document.querySelectorAll('.route-card.drop-alvo,.route-card.drop-invalido')
    .forEach(c => c.classList.remove('drop-alvo','drop-invalido'));
  document.querySelectorAll('.viagem-header.drop-viagem-reorder')
    .forEach(h => h.classList.remove('drop-viagem-reorder'));
}
function dragEnterCard(event, veiculoId) {
  if (!dragInfo) return;
  if (dragInfo.tipo === 'viagem') {
    event.preventDefault();
    event.currentTarget.classList.remove('drop-invalido');
    event.currentTarget.classList.toggle('drop-invalido', dragInfo.veiculoId === veiculoId);
    event.currentTarget.classList.toggle('drop-alvo',     dragInfo.veiculoId !== veiculoId);
  } else if (dragInfo.tipo === 'parada' && dragInfo.veiculoId !== veiculoId) {
    event.preventDefault();
    event.currentTarget.classList.add('drop-alvo');
  }
}
function dragLeaveCard(event) {
  if (!event.currentTarget.contains(event.relatedTarget))
    event.currentTarget.classList.remove('drop-alvo', 'drop-invalido');
}
function dragOverCard(event, veiculoId) {
  if (!dragInfo) return;
  if (dragInfo.tipo === 'viagem') {
    event.preventDefault();
    event.dataTransfer.dropEffect = dragInfo.veiculoId === veiculoId ? 'none' : 'move';
  } else if (dragInfo.tipo === 'parada' && dragInfo.veiculoId !== veiculoId) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
  }
}
// Helper: aloca parada em um veículo, respeitando capacidade
// Se a viagem preferida (tripIndexPref) estiver cheia, tenta outra; se nenhuma couber, cria nova viagem.
function _alocarParada(parada, veiculoId, tripIndexPref) {
  const vDest   = veiculos.find(x => x.id === veiculoId);
  const cap     = vDest?.capacidade ?? Infinity;
  const viagens = ultimoResultado[veiculoId] || [];
  const vol     = parada.volumeTotal || 0;
  // Tenta a viagem preferida primeiro
  if (tripIndexPref != null) {
    const vi = viagens[tripIndexPref];
    if (vi) {
      const ocupado = vi.paradas.reduce((s,p)=>s+(p.volumeTotal||0), 0);
      if (ocupado + vol <= cap + 0.001) { vi.paradas.push(parada); return false; }
    }
  }
  // Tenta qualquer viagem com espaço
  const viLivre = viagens.find(vi => {
    const ocupado = vi.paradas.reduce((s,p)=>s+(p.volumeTotal||0), 0);
    return ocupado + vol <= cap + 0.001;
  });
  if (viLivre) { viLivre.paradas.push(parada); return false; }
  // Nenhuma viagem comporta → cria viagem nova com terminal do pedido da parada
  const termOrig = parada?.pedido?.terminal || viagens[0]?.terminalOrigem || vDest?.terminal || '';
  if (!ultimoResultado[veiculoId]) ultimoResultado[veiculoId] = [];
  const novaVi = { paradas: [parada], quebras: [], terminalOrigem: termOrig, esperaTerminalMin: 0, tempoConsumidoMin: 0 };
  ultimoResultado[veiculoId].push(novaVi);
  if (vDest) recalcularTimingViagem(novaVi, vDest);
  return true; // criou nova viagem
}
function dropNaCard(event, veiculoDestId) {
  event.currentTarget.classList.remove('drop-alvo', 'drop-invalido');
  if (!dragInfo) return;
  if (dragInfo.tipo === 'viagem') {
    if (dragInfo.veiculoId === veiculoDestId) { dragInfo = null; return; }
    const { veiculoId: origId, tripIndex } = dragInfo;
    dragInfo = null;
    historicoManual.push(JSON.parse(JSON.stringify(ultimoResultado)));
    limparFantasma(origId);
    const [viagem] = (ultimoResultado[origId] || []).splice(tripIndex, 1);
    if ((ultimoResultado[origId] || []).length === 0) marcarVeiculoVazio(origId);
    limparFantasma(veiculoDestId);
    if (!ultimoResultado[veiculoDestId]) ultimoResultado[veiculoDestId] = [];
    ultimoResultado[veiculoDestId].push(viagem);
    // Recalcula tempos da viagem no contexto do novo veículo
    const _vDest = veiculos.find(x => x.id === veiculoDestId);
    if (_vDest) recalcularTimingViagem(viagem, _vDest);
  } else if (dragInfo.tipo === 'parada') {
    if (dragInfo.veiculoId === veiculoDestId) { dragInfo = null; return; }
    const { veiculoId: origId, tripIndex: origTripIdx, stopIndex } = dragInfo;
    dragInfo = null;
    historicoManual.push(JSON.parse(JSON.stringify(ultimoResultado)));
    limparFantasma(origId);
    const [parada] = ultimoResultado[origId][origTripIdx].paradas.splice(stopIndex, 1);
    if (ultimoResultado[origId][origTripIdx].paradas.length === 0) {
      ultimoResultado[origId].splice(origTripIdx, 1);
      if (ultimoResultado[origId].length === 0) marcarVeiculoVazio(origId);
    }
    limparFantasma(veiculoDestId);
    _alocarParada(parada, veiculoDestId, null);
  } else { dragInfo = null; return; }
  atualizarBarraManual();
  renderResultado(ultimoResultado, ultimoControleTempo || {});
}
function desfazerManual() {
  if (!historicoManual.length) return;
  ultimoResultado = historicoManual.pop();
  atualizarBarraManual();
  renderResultado(ultimoResultado, ultimoControleTempo || {});
}
// ── Drag de parada (entrega individual) ──────────────────────────────────────
function iniciarDragParada(veiculoId, tripIndex, stopIndex, el) {
  dragInfo = { tipo: 'parada', veiculoId, tripIndex, stopIndex };
  setTimeout(() => el.classList.add('stop-draggando'), 0);
}
function fimDragParada(el) {
  el.classList.remove('stop-draggando');
  dragInfo = null;
  document.querySelectorAll('.viagem-header.drop-viagem-alvo,.viagem-header.drop-viagem-invalido')
    .forEach(h => h.classList.remove('drop-viagem-alvo','drop-viagem-invalido'));
  document.querySelectorAll('.route-card.drop-alvo,.route-card.drop-invalido')
    .forEach(c => c.classList.remove('drop-alvo','drop-invalido'));
  document.querySelectorAll('.stop.stop-reorder-alvo')
    .forEach(s => s.classList.remove('stop-reorder-alvo'));
}
function dragEnterStop(event, veiculoId, tripIdx, stopIdx) {
  if (!dragInfo || dragInfo.tipo !== 'parada') return;
  if (dragInfo.veiculoId === veiculoId && dragInfo.tripIndex === tripIdx && dragInfo.stopIndex === stopIdx) return;
  event.preventDefault();
  event.stopPropagation();
  event.currentTarget.classList.add('stop-reorder-alvo');
}
function dragLeaveStop(event) {
  if (!event.currentTarget.contains(event.relatedTarget))
    event.currentTarget.classList.remove('stop-reorder-alvo');
}
function dragOverStop(event, veiculoId, tripIdx, stopIdx) {
  if (!dragInfo || dragInfo.tipo !== 'parada') return;
  if (dragInfo.veiculoId === veiculoId && dragInfo.tripIndex === tripIdx && dragInfo.stopIndex === stopIdx) return;
  event.preventDefault();
  event.stopPropagation();
  event.dataTransfer.dropEffect = 'move';
}
function dropNaStop(event, veiculoDestId, tripDestIdx, stopDestIdx) {
  event.stopPropagation();
  event.currentTarget.classList.remove('stop-reorder-alvo');
  if (!dragInfo || dragInfo.tipo !== 'parada') return;
  const { veiculoId: origId, tripIndex: origTripIdx, stopIndex: origStopIdx } = dragInfo;
  dragInfo = null;
  if (origId === veiculoDestId && origTripIdx === tripDestIdx && origStopIdx === stopDestIdx) return;
  historicoManual.push(JSON.parse(JSON.stringify(ultimoResultado)));
  limparFantasma(origId);
  if (origId === veiculoDestId && origTripIdx === tripDestIdx) {
    // Mesma viagem: reordena
    const paradas = ultimoResultado[origId][origTripIdx].paradas;
    const [parada] = paradas.splice(origStopIdx, 1);
    const insertIdx = origStopIdx < stopDestIdx ? stopDestIdx - 1 : stopDestIdx;
    paradas.splice(insertIdx, 0, parada);
  } else {
    // Viagem ou veículo diferente: remove da origem e insere na posição destino
    const [parada] = ultimoResultado[origId][origTripIdx].paradas.splice(origStopIdx, 1);
    if (ultimoResultado[origId][origTripIdx].paradas.length === 0) {
      ultimoResultado[origId].splice(origTripIdx, 1);
      if (ultimoResultado[origId].length === 0) marcarVeiculoVazio(origId);
    }
    limparFantasma(veiculoDestId);
    const vDest = veiculos.find(x => x.id === veiculoDestId);
    const cap = vDest?.capacidade ?? Infinity;
    const destVi = (ultimoResultado[veiculoDestId] || [])[tripDestIdx];
    if (destVi) {
      const volDest = destVi.paradas.reduce((s,p) => s + (p.volumeTotal||0), 0);
      if (volDest + (parada.volumeTotal||0) <= cap + 0.001) {
        destVi.paradas.splice(stopDestIdx, 0, parada); // insere na posição exata
      } else {
        _alocarParada(parada, veiculoDestId, null); // cria nova viagem se cheio
      }
    } else {
      _alocarParada(parada, veiculoDestId, null);
    }
  }
  recalcularTodasViagens(origId, veiculoDestId);
  atualizarBarraManual();
  renderResultado(ultimoResultado, ultimoControleTempo || {});
}
function dragEnterViagem(event, veiculoId, tripIndex) {
  if (!dragInfo) return;
  if (dragInfo.tipo === 'parada') {
    event.preventDefault();
    event.stopPropagation();
    const mesma = dragInfo.veiculoId === veiculoId && dragInfo.tripIndex === tripIndex;
    event.currentTarget.classList.toggle('drop-viagem-alvo',    !mesma);
    event.currentTarget.classList.toggle('drop-viagem-invalido', mesma);
  } else if (dragInfo.tipo === 'viagem' && dragInfo.veiculoId === veiculoId) {
    // Mesma placa: reordenar viagens
    if (dragInfo.tripIndex === tripIndex) return;
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.add('drop-viagem-reorder');
  }
  // Viagem de placa diferente: deixa borbulhar para o route-card
}
function dragLeaveViagem(event) {
  if (!event.currentTarget.contains(event.relatedTarget))
    event.currentTarget.classList.remove('drop-viagem-alvo','drop-viagem-invalido','drop-viagem-reorder');
}
function dragOverViagem(event, veiculoId, tripIndex) {
  if (!dragInfo) return;
  if (dragInfo.tipo === 'parada') {
    event.preventDefault();
    event.stopPropagation();
    const mesma = dragInfo.veiculoId === veiculoId && dragInfo.tripIndex === tripIndex;
    event.dataTransfer.dropEffect = mesma ? 'none' : 'move';
  } else if (dragInfo.tipo === 'viagem' && dragInfo.veiculoId === veiculoId) {
    if (dragInfo.tripIndex === tripIndex) return;
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = 'move';
  }
}
function dropNaViagem(event, veiculoId, tripIndex) {
  event.stopPropagation();
  event.currentTarget.classList.remove('drop-viagem-alvo','drop-viagem-invalido','drop-viagem-reorder');
  if (!dragInfo) return;
  // Reordenar viagens dentro da mesma placa
  if (dragInfo.tipo === 'viagem' && dragInfo.veiculoId === veiculoId) {
    const origTripIdx = dragInfo.tripIndex;
    dragInfo = null;
    if (origTripIdx === tripIndex) return;
    historicoManual.push(JSON.parse(JSON.stringify(ultimoResultado)));
    const viagens = ultimoResultado[veiculoId];
    const [vi] = viagens.splice(origTripIdx, 1);
    const insertIdx = origTripIdx < tripIndex ? tripIndex - 1 : tripIndex;
    viagens.splice(insertIdx, 0, vi);
    recalcularTodasViagens(veiculoId);
    atualizarBarraManual();
    renderResultado(ultimoResultado, ultimoControleTempo || {});
    return;
  }
  if (dragInfo.tipo !== 'parada') return;
  const { veiculoId: origVeicId, tripIndex: origTripIdx, stopIndex } = dragInfo;
  dragInfo = null;
  if (origVeicId === veiculoId && origTripIdx === tripIndex) return;
  historicoManual.push(JSON.parse(JSON.stringify(ultimoResultado)));
  // Remove parada da origem
  limparFantasma(origVeicId);
  const [parada] = ultimoResultado[origVeicId][origTripIdx].paradas.splice(stopIndex, 1);
  // Se a viagem origem ficou vazia, remove-a e ajusta o índice destino
  if (ultimoResultado[origVeicId][origTripIdx].paradas.length === 0) {
    ultimoResultado[origVeicId].splice(origTripIdx, 1);
    if (origVeicId === veiculoId && origTripIdx < tripIndex) tripIndex--;
    if (ultimoResultado[origVeicId].length === 0) marcarVeiculoVazio(origVeicId);
  }
  // Aloca com check de capacidade: se a viagem alvo estiver cheia, cria nova viagem
  limparFantasma(veiculoId);
  _alocarParada(parada, veiculoId, tripIndex);
  recalcularTodasViagens(origVeicId, veiculoId);
  atualizarBarraManual();
  renderResultado(ultimoResultado, ultimoControleTempo || {});
}
// ── Troca de placa (veículo) ──────────────────────────────────────────────────
var _placaDropOrigId = null;
function abrirTrocaPlaca(veiculoId, btn, event) {
  event.stopPropagation();
  const menu = document.getElementById('placa-dropdown-menu');
  if (!menu) return;
  // Fecha se já estava aberto para o mesmo veículo
  if (_placaDropOrigId === veiculoId && menu.style.display !== 'none') {
    menu.style.display = 'none';
    _placaDropOrigId = null;
    return;
  }
  _placaDropOrigId = veiculoId;
  const vAtual  = veiculos.find(x => x.id === veiculoId);
  const termAtual = baseVeiculoLabel(vAtual);
  const outros = veiculos
    .filter(vv => vv.id !== veiculoId)
    .sort((a, b) => (baseVeiculoLabel(a) === termAtual ? 0 : 1) - (baseVeiculoLabel(b) === termAtual ? 0 : 1));
  menu.innerHTML = `
    <div style="padding:7px 12px;font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text-2);border-bottom:1px solid var(--border);">
      Trocar ${vAtual?.placa || ''} por:
    </div>
    ${outros.map(vv => {
      const mesmaBase  = baseVeiculoLabel(vv) === termAtual;
      const viagens    = (ultimoResultado[vv.id] || []).filter(vi => !vi._vazio && vi.paradas?.length > 0);
      const emUso      = viagens.length > 0;
      const volUsado   = viagens.reduce((s, vi) => s + vi.paradas.reduce((ss, p) => ss + (p.volumeTotal || 0), 0), 0);
      const statusHtml = emUso
        ? `<span style="color:#D97706;font-size:10px;font-weight:700;white-space:nowrap;">● em uso · ${volUsado.toFixed(0)} m³</span>`
        : `<span style="color:var(--pet-green);font-size:10px;font-weight:700;white-space:nowrap;">● disponível</span>`;
      return `<div onclick="confirmarTrocaPlaca(${vv.id})"
        style="padding:9px 12px;cursor:pointer;display:flex;align-items:center;gap:8px;border-bottom:1px solid var(--border);"
        onmouseenter="this.style.background='var(--bg)'" onmouseleave="this.style.background=''">
        <span style="font-family:var(--font-cond);font-weight:800;font-size:13px;letter-spacing:.06em;">${vv.placa}</span>
        <span style="color:var(--text-2);font-size:11px;flex:1;">${vv.transportadora ? vv.transportadora + ' · ' : ''}${vv.tipo} · ${vv.capacidade.toFixed(0)} m³</span>
        ${statusHtml}
        ${mesmaBase ? `<span style="color:var(--pet-green);font-size:10px;font-weight:700;">mesma base</span>` : `<span style="color:var(--text-3);font-size:10px;">${baseVeiculoLabel(vv)}</span>`}
      </div>`;
    }).join('')}`;
  const rect = btn.getBoundingClientRect();
  menu.style.top   = (rect.bottom + 4) + 'px';
  menu.style.left  = rect.left + 'px';
  menu.style.right = 'auto';
  menu.style.display = 'block';
  // Desliza para a esquerda só o necessário para não sair da tela
  requestAnimationFrame(() => {
    const mr = menu.getBoundingClientRect();
    const overflow = mr.right - (window.innerWidth - 8);
    if (overflow > 0) {
      menu.style.left = Math.max(8, rect.left - overflow) + 'px';
    }
  });
}
function confirmarTrocaPlaca(veiculoDestId) {
  const menu = document.getElementById('placa-dropdown-menu');
  if (menu) menu.style.display = 'none';
  if (!_placaDropOrigId || _placaDropOrigId === veiculoDestId) { _placaDropOrigId = null; return; }
  historicoManual.push(JSON.parse(JSON.stringify(ultimoResultado)));
  const origId = _placaDropOrigId;
  // Troca as viagens entre os dois veículos
  const tripOrig = ultimoResultado[origId]       ? [...ultimoResultado[origId]]       : [];
  const tripDest = ultimoResultado[veiculoDestId] ? [...ultimoResultado[veiculoDestId]] : [];
  ultimoResultado[origId]        = tripDest;
  ultimoResultado[veiculoDestId] = tripOrig;
  // Marca como vazio se ficou sem viagens reais
  if (!ultimoResultado[origId].some(vi => !vi._vazio && vi.paradas?.length))
    marcarVeiculoVazio(origId);
  if (!ultimoResultado[veiculoDestId].some(vi => !vi._vazio && vi.paradas?.length))
    marcarVeiculoVazio(veiculoDestId);
  _placaDropOrigId = null;
  // Recalcula timing e ressincroniza controleTempo para ambos os veículos
  recalcularTodasViagens(origId, veiculoDestId);
  atualizarBarraManual();
  renderResultado(ultimoResultado, ultimoControleTempo || {});
}
function toggleLockTerminal(terminal, btn) {
  if (lockedTerminals.has(terminal)) {
    lockedTerminals.delete(terminal);
  } else {
    lockedTerminals.add(terminal);
  }
  if (ultimoResultado) renderResultado(ultimoResultado, ultimoControleTempo || {});
}
// Fecha dropdown ao clicar fora
document.addEventListener('click', () => {
  const menu = document.getElementById('placa-dropdown-menu');
  if (menu) menu.style.display = 'none';
  _placaDropOrigId = null;
});
function confirmarReset(btn) {
  const popup = document.getElementById('resetar-confirm');
  if (!popup) return;
  popup.style.display = popup.style.display === 'none' ? 'block' : 'none';
}
function fecharConfirmReset() {
  const popup = document.getElementById('resetar-confirm');
  if (popup) popup.style.display = 'none';
}
function resetarOtimizacao() {
  fecharConfirmReset();
  // usa resultadoOriginal se disponível; fallback para o primeiro estado salvo no histórico
  const base = resultadoOriginal || (historicoManual.length > 0 ? historicoManual[0] : null);
  if (!base) return;
  try {
    ultimoResultado = JSON.parse(JSON.stringify(base));
    historicoManual = [];
    lockedTerminals.clear();
    atualizarBarraManual();
    renderResultado(ultimoResultado, ultimoControleTempo || {});
  } catch(e) {
    console.error('Erro ao resetar:', e);
  }
}
// ── Resumo do nome do produto: remove código SAP e prefixo de marca ──────────
function resumoProduto(nome) {
  if (!nome) return '';
  let s = nome.toString().trim();
  s = s.replace(/^\d+\s*-\s*/, '');       // remove "2000034 - "
  s = s.replace(/^PETRONAS\s+/i, '');     // remove prefixo PETRONAS redundante
  return s;
}
// ── Mapa completo do veículo (todas as viagens) ────────────────────────────────
function _freteNum(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const s = String(v).trim().replace(/[^\d,.-]/g, '');
  const n = s.includes(',')
    ? parseFloat(s.replace(/\./g, '').replace(',', '.'))
    : parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function _freteMoeda(v) {
  return 'R$ ' + (Number(v) || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function _freteHoras(min) {
  return ((Number(min) || 0) / 60).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'h';
}
function _freteDataBase(viagem) {
  const iso = document.getElementById('rot-data-operacao')?.value;
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso || '')) {
    const [y, m, d] = iso.split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  const br = viagem?.paradas?.[0]?.pedido?.dataEntregaLogistica || viagem?.paradas?.[0]?.pedido?.dataEntrega;
  const mt = String(br || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mt) return new Date(Number(mt[3]), Number(mt[2]) - 1, Number(mt[1]));
  return new Date();
}
function _freteDiasMes(viagem) {
  const d = _freteDataBase(viagem);
  return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
}
function _freteTempoViagemMin(viagem, v, kmMapa) {
  const kmRealMapa = Number(kmMapa) || 0;
  if (kmRealMapa > 0 && v) {
    const velCarregado = Number(v.velMediaCarregado) || 45;
    const velVazio = Number(v.velMediaVazio) || 55;
    const paradas = viagem?.paradas || [];
    const kmCarregadoBase = paradas.reduce((s, p) => s + (((p.deslocCarregadoMin || 0) / 60) * velCarregado), 0);
    const kmVazioBase = paradas.reduce((s, p) => s + (((p.deslocVazioMin || 0) / 60) * velVazio), 0);
    const kmBaseTotal = kmCarregadoBase + kmVazioBase;
    const propCarregado = kmBaseTotal > 0 ? kmCarregadoBase / kmBaseTotal : 0.85;
    const propVazio = 1 - propCarregado;
    const tempoMapaMin = ((kmRealMapa * propCarregado) / velCarregado) * 60
      + ((kmRealMapa * propVazio) / velVazio) * 60;
    const servicoMin = paradas.reduce((s, p, i) => s
      + (i === 0 ? (p.tempoCarregamentoMin || 0) : 0)
      + (p.waitAfterLoadingMin || 0)
      + (p.tempoEsperaRestricaoMin || 0)
      + (p.tempoDescargaMin || 0), 0);
    return tempoMapaMin + servicoMin;
  }
  const direto = Number(viagem?.tempoConsumidoMin);
  if (direto > 0) return direto;
  return (viagem?.paradas || []).reduce((s, p, i) => s
    + (i === 0 ? (p.tempoCarregamentoMin || 0) : 0)
    + (p.waitAfterLoadingMin || 0)
    + (p.deslocCarregadoMin || 0)
    + (p.tempoEsperaRestricaoMin || 0)
    + (p.tempoDescargaMin || 0)
    + (p.deslocVazioMin || 0), 0);
}
function _freteParcelaFixaMensal(v, viagem, fixoMensal, kmMapa) {
  const diasMes = _freteDiasMes(viagem) || 30;
  const fixoDiario = (Number(fixoMensal) || 0) / diasMes;
  const jornadaDispMin = Number(v?.jornadaMin) || duracaoJornadaMin(v?.jornadaInicio || '06:00', v?.jornadaFim || '18:00') || 720;
  const jornadaUsadaMin = _freteTempoViagemMin(viagem, v, kmMapa);
  const fatorJornada = jornadaDispMin > 0 ? Math.min(1, Math.max(0, jornadaUsadaMin / jornadaDispMin)) : 0;
  const fixoViagem = fixoDiario * fatorJornada;
  return { diasMes, fixoDiario, jornadaDispMin, jornadaUsadaMin, fatorJornada, fixoViagem };
}
function calcularFreteViagemAtual(v, viagem, kmMapa) {
  const contrato = freteCarregarContratos().find(c => (c.placa || '').toUpperCase() === (v.placa || '').toUpperCase());
  const spots = freteCarregarSpot();
  const volumeM3 = (viagem.paradas || []).reduce((s, p) => s + (p.volumeTotal || 0), 0);
  const km = Number(kmMapa) || 0;
  if (!contrato) return { contrato: null, km, volumeM3, custo: 0, tipo: 'Sem contrato', detalhe: 'Cadastre contrato para esta placa na aba Frete.' };
  const fixoMensal = _freteNum(contrato.fixo);
  const fixa = _freteParcelaFixaMensal(v, viagem, fixoMensal, km);
  let custo = 0, detalhe = '';
  if (contrato.tipo === 'fixo_km') {
    const taxaKm = _freteNum(contrato.km);
    const variavel = taxaKm * km;
    custo = fixa.fixoViagem + variavel;
    detalhe = `Fixo mensal ${_freteMoeda(fixoMensal)} / ${fixa.diasMes} dias x ${_freteHoras(fixa.jornadaUsadaMin)}/${_freteHoras(fixa.jornadaDispMin)} = ${_freteMoeda(fixa.fixoViagem)} + ${_freteMoeda(taxaKm)}/km`;
  } else if (contrato.tipo === 'fixo_m3') {
    const taxaM3 = _freteNum(contrato.m3);
    const variavel = taxaM3 * volumeM3;
    custo = fixa.fixoViagem + variavel;
    detalhe = `Fixo mensal ${_freteMoeda(fixoMensal)} / ${fixa.diasMes} dias x ${_freteHoras(fixa.jornadaUsadaMin)}/${_freteHoras(fixa.jornadaDispMin)} = ${_freteMoeda(fixa.fixoViagem)} + ${_freteMoeda(taxaM3)}/m3`;
  } else if (contrato.tipo === 'diaria') {
    custo = _freteNum(contrato.diaria);
    detalhe = `${_freteMoeda(custo)} diaria`;
  } else if (contrato.tipo === 'spot') {
    const origem = (viagem.terminalOrigem || v.terminal || '').toLowerCase();
    const destinos = (viagem.paradas || []).map(p => `${p.pedido?.cidade || ''} ${p.pedido?.cliente || ''}`).join(' ').toLowerCase();
    const sp = spots.find(s => origem.includes((s.origem || '').toLowerCase()) && destinos.includes((s.destino || '').toLowerCase()) && (!s.transportadora || s.transportadora === contrato.transportadora));
    const taxaSpot = sp ? _freteNum(sp.valor) : 0;
    custo = taxaSpot * volumeM3;
    detalhe = sp ? `${_freteMoeda(taxaSpot)}/m3 spot` : 'Spot sem rota cadastrada';
  }
  return { contrato, km, volumeM3, custo, tipo: FRETE_TIPOS[contrato.tipo] || contrato.tipo || 'Contrato', detalhe, ...fixa, fixoMensal };
}
function renderCustoMapaViagem() {
  const box = document.getElementById('mv-custo');
  if (!box || window._mvVeiculoId == null || window._mvIdxViagem == null || !ultimoResultado) return;
  const v = veiculos.find(x => x.id === window._mvVeiculoId);
  const viagem = ultimoResultado?.[window._mvVeiculoId]?.[window._mvIdxViagem];
  if (!v || !viagem) return;
  const kmOriginal = (viagem.paradas || []).reduce((s, p) => s + (p.distanciaKm || 0), 0);
  const calc = calcularFreteViagemAtual(v, viagem, viagem._kmAjustado != null ? viagem._kmAjustado : kmOriginal);
  const totalLitros = calc.volumeM3 * 1000;
  const custoLitroMedio = totalLitros > 0 ? calc.custo / totalLitros : 0;
  const paradasFrete = viagem.paradas || [];
  const somaKmParadas = paradasFrete.reduce((s, p) => s + (p.distanciaKm > 0 ? p.distanciaKm : 0), 0);
  let somaPesoRateio = 0;
  const pesosRateio = paradasFrete.map((p, i) => {
    const vol = p.volumeTotal || 0;
    const kmRef = somaKmParadas > 0 && p.distanciaKm > 0
      ? calc.km * (p.distanciaKm / somaKmParadas)
      : (calc.km > 0 ? calc.km / Math.max(paradasFrete.length, 1) : 1);
    const peso = vol * kmRef;
    somaPesoRateio += peso;
    return { vol, kmRef, peso };
  });
  if (!(somaPesoRateio > 0)) {
    somaPesoRateio = paradasFrete.reduce((s, p) => s + (p.volumeTotal || 0), 0);
    pesosRateio.forEach((r, i) => { r.peso = paradasFrete[i]?.volumeTotal || 0; });
  }
  const custoTotalCent = Math.round((calc.custo || 0) * 100);
  let custoRateadoAcumCent = 0;
  const clienteHtml = paradasFrete.map((p, i) => {
    const vol = p.volumeTotal || 0;
    const isLast = i === paradasFrete.length - 1;
    const custoClienteCent = isLast
      ? Math.max(0, custoTotalCent - custoRateadoAcumCent)
      : Math.round(somaPesoRateio > 0 ? custoTotalCent * (pesosRateio[i].peso / somaPesoRateio) : 0);
    custoRateadoAcumCent += custoClienteCent;
    const custoCliente = custoClienteCent / 100;
    const custoLitro = vol > 0 ? custoCliente / (vol * 1000) : 0;
    const nome = String(p.pedido?.cliente || '-').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return `<div style="display:flex;align-items:center;gap:6px;min-width:0;">
      <span style="width:18px;height:18px;border-radius:50%;background:#4F46E5;color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0;">${i+1}</span>
      <span style="font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;">${nome}</span>
      <span style="color:var(--text-3);white-space:nowrap;">${vol.toFixed(1)} m3</span>
      <span style="color:var(--text-3);white-space:nowrap;">${pesosRateio[i].kmRef.toFixed(1)} km</span>
      <span style="color:#15803d;font-weight:700;white-space:nowrap;">${_freteMoeda(custoCliente)}</span>
      <span style="color:#3730A3;background:rgba(79,70,229,0.09);border:1px solid rgba(79,70,229,0.2);border-radius:999px;padding:1px 7px;white-space:nowrap;">R$ ${custoLitro.toLocaleString('pt-BR',{minimumFractionDigits:3,maximumFractionDigits:3})}/L</span>
    </div>`;
  }).join('');
  box.innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;margin-bottom:7px;">
      <div><div style="font-size:9px;color:var(--text-3);font-weight:700;letter-spacing:.06em;text-transform:uppercase;">Contrato</div><div style="font-weight:700;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${calc.tipo}</div></div>
      <div><div style="font-size:9px;color:var(--text-3);font-weight:700;letter-spacing:.06em;text-transform:uppercase;">KM mapa</div><div style="font-weight:700;color:#3730A3;">${calc.km.toFixed(1)} km</div></div>
      <div><div style="font-size:9px;color:var(--text-3);font-weight:700;letter-spacing:.06em;text-transform:uppercase;">Custo viagem</div><div style="font-weight:800;color:#15803d;">${_freteMoeda(calc.custo)}</div></div>
      <div><div style="font-size:9px;color:var(--text-3);font-weight:700;letter-spacing:.06em;text-transform:uppercase;">Media</div><div style="font-weight:700;color:var(--text);">R$ ${custoLitroMedio.toLocaleString('pt-BR',{minimumFractionDigits:3,maximumFractionDigits:3})}/L</div></div>
    </div>
    <div style="font-size:10px;color:var(--text-3);margin-bottom:5px;">${calc.detalhe || '&nbsp;'} · Rateio por volume x distancia da parada; ultimo cliente ajusta centavos para fechar o total.</div>
    <div style="display:grid;gap:4px;font-size:11px;line-height:1.35;">${clienteHtml || '<span style="color:var(--text-3);">Sem clientes para rateio.</span>'}</div>`;
}
window.renderCustoMapaViagem = renderCustoMapaViagem;
const CORES_VIAGEM = ['#7CB82B','#2255CC','#F97316','#7C3AED','#E11D48','#0EA5E9'];
function abrirMapaVeiculo(veiculoId) {
  if (!ultimoResultado || !ultimoResultado[veiculoId]) return;
  const v = veiculos.find(x => x.id === veiculoId);
  if (!v) return;
  const viagens = (ultimoResultado[veiculoId] || []).filter(vi => vi.paradas?.length);
  if (!viagens.length) { alert('Nenhuma viagem com paradas para este veículo.'); return; }
  const totalVol  = viagens.reduce((s,vi) => s + vi.paradas.reduce((ss,p)=>ss+(p.volumeTotal||0),0), 0);
  const totalDest = viagens.reduce((s,vi) => s + vi.paradas.length, 0);
  document.getElementById('mv-titulo').textContent =
    `${v.placa} — Todas as viagens${v.transportadora ? ' · '+v.transportadora : ''}`;
  document.getElementById('mv-resumo').innerHTML =
    `<strong>${viagens.length}</strong> viagem(ns) · <strong>${totalDest}</strong> destinos · Volume total: <strong>${totalVol.toFixed(1)} m³</strong>`;
  document.getElementById('mv-lista').innerHTML = viagens.map((vi, ti) => {
    const cor = CORES_VIAGEM[ti % CORES_VIAGEM.length];
    const volVi = vi.paradas.reduce((s,p)=>s+(p.volumeTotal||0),0);
    return `
      <div style="margin-bottom:14px;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
          <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:${cor};flex-shrink:0;"></span>
          <span style="font-family:var(--font-cond);font-weight:700;font-size:13px;letter-spacing:0.05em;">Viagem ${ti+1}${vi.petId ? ` · ${vi.petId}` : ''}</span>
          <span class="tag tag-lime" style="font-size:9px;">${volVi.toFixed(1)} m³ · ${vi.paradas.length} parada(s)</span>
        </div>
        ${vi.paradas.map((p,i) => `
          <div class="card" style="margin-bottom:6px;padding:8px 12px;">
            <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
              <span style="background:${cor};color:#fff;font-size:10px;font-weight:700;padding:1px 8px;border-radius:20px;">${ti+1}.${i+1}</span>
              <span style="font-weight:700;font-size:12px;">${p.pedido.cliente}</span>
              <span style="font-size:11px;color:#4A6535;">${p.pedido.cidade||'-'}</span>
              <span class="tag tag-lime" style="font-size:9px;">${(p.volumeTotal||0).toFixed(1)} m³</span>
            </div>
          </div>`).join('')}
      </div>`;
  }).join('');
  document.getElementById('modal-viagem').classList.add('show');
  garantirMapaViagem();
  const todosLatlngs = [];
  viagens.forEach((vi, ti) => {
    const cor = CORES_VIAGEM[ti % CORES_VIAGEM.length];
    const termOrig = terminaisCad.find(t => t.nome === (vi.terminalOrigem || v.terminal || ''));
    const pontos = [];
    if (termOrig && !isNaN(termOrig.lat) && !isNaN(termOrig.lon))
      pontos.push({ lat: +termOrig.lat, lon: +termOrig.lon, label: `${ti+1}.O`, nome: termOrig.nome, tipo: 'origem' });
    vi.paradas.forEach((p, i) => {
      const _coord = latLonEfetivo(p.pedido);
      if (!isNaN(_coord.lat) && !isNaN(_coord.lon))
        pontos.push({ lat: _coord.lat, lon: _coord.lon, isCentroide: _coord.isCentroide, label: `${ti+1}.${i+1}`, nome: p.pedido.cliente, tipo: 'destino' });
    });
    if (pontos.length >= 2)
      osrmRoute(pontos, camadaViagem, cor, 4, 0.9).catch(()=>{});
    pontos.forEach((p, i) => {
      todosLatlngs.push([p.lat, p.lon]);
      const isOrigem = p.tipo === 'origem';
      const isRetorno = p.tipo === 'retorno';
      const seq = isOrigem ? '⬤' : isRetorno ? '↩' : p.label || String(i);
      const bg = isOrigem ? '#111827' : isRetorno ? '#6B7280' : cor;
      const icon = L.divIcon({
        className: '',
        html: `<div style="width:24px;height:24px;border-radius:50%;background:${bg};color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;border:2px solid #fff;box-shadow:0 1px 4px rgba(0,0,0,0.3);font-family:Inter,sans-serif;">${seq}</div>`,
        iconSize: [24, 24], iconAnchor: [12, 12],
      });
      const popupHtml = `
        <div style="font-family:Inter,sans-serif;min-width:180px;">
          <div style="font-weight:700;font-size:12px;margin-bottom:4px;">${p.nome}</div>
          ${p.volume ? `<div style="font-size:11px;">Volume: <strong>${p.volume.toFixed(1)} m³</strong></div>` : ''}
          ${p.cidade ? `<div style="font-size:11px;color:#6B7280;">📍 ${p.cidade}</div>` : ''}
          ${p.restricao ? `<div style="font-size:10px;margin-top:4px;color:#D97706;">⏱ ${p.restricao}</div>` : ''}
        </div>`;
      L.marker([p.lat, p.lon], { icon }).addTo(camadaViagem)
        .bindPopup(popupHtml, { maxWidth: 240 })
        .bindTooltip(seq, { permanent: true, direction: 'top', className: 'leaflet-label' });
    });
  });
  if (todosLatlngs.length) mapaViagem.fitBounds(todosLatlngs, { padding: [30, 30] });
}
// Reconstrói o contador de petId a partir de TODOS os arquivos do histórico,
// descartando qualquer valor inflado que possa estar no localStorage.
async function _reconstruirPetSeqDoHistorico({ obrigatorio = false } = {}) {
  if (!dirHandleHistorico) {
    if (obrigatorio) throw new Error('PASTA_NAO_SELECIONADA');
    return;
  }
  let permOk = false;
  try { permOk = (await dirHandleHistorico.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  if (!permOk) {
    if (obrigatorio) throw new Error('PERMISSAO_NEGADA');
    return;
  }
  const stored = {};
  for await (const [name, handle] of dirHandleHistorico.entries()) {
    if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
    try {
      const data = JSON.parse(await (await handle.getFile()).text());
      const resultado = data.resultado;
      if (!resultado) continue;
      Object.values(resultado).forEach(viagens => {
        (viagens || []).forEach(vi => {
          // Formato novo: P{MM}{YY}{NNN} → chave = 4 dígitos
          let m = (vi.petId || '').match(/^P(\d{4})(\d{3,})$/);
          if (!m) {
            // Formato legado: P{MM}{YY}{DD}{NNN} → migra chave para MM+YY
            const ml = (vi.petId || '').match(/^P(\d{2})(\d{2})\d{2}(\d{3,})$/);
            if (ml) m = [null, ml[1] + ml[2], ml[3]];
          }
          if (!m) return;
          const [, chave, seq] = m;
          const n = parseInt(seq);
          if (!stored[chave] || n > stored[chave]) stored[chave] = n;
        });
      });
    } catch(e) {}
  }
  try { localStorage.setItem('petSeq', JSON.stringify(stored)); } catch(e) {}
}
// Retorna o motorista correto — override manual por petId tem prioridade sobre a regra diurno/noturno
function motoristaDaViagem(v, cargaAbsMin, vi) {
  if (vi?.petId && _motoristasOverride[vi.petId] !== undefined) return _motoristasOverride[vi.petId];
  const hod = ((Math.round(cargaAbsMin) % 1440) + 1440) % 1440; // hora do dia em min
  // Usa jornadaInicio do veículo como limiar — ex: jornada 05:00 → carga às 05h é diurna
  const _jornadaIniHod = parseHoraMin(v?.jornadaInicio || '06:00') || 360;
  const ehNoturno = hod >= 18 * 60 || hod < _jornadaIniHod;
  if (ehNoturno && v.motoristaNt)      return v.motoristaNt;
  if (!ehNoturno && v.motoristaDiurno) return v.motoristaDiurno;
  return v.motoristaDiurno || v.motoristaNt || '';
}
// Edição inline de motorista por viagem — qualquer valor digitado sobrepõe a regra
function _editMotoristaViagem(petId, val) {
  if (val.trim()) { _motoristasOverride[petId] = val; } else { delete _motoristasOverride[petId]; }
  renderTemplateOperacao();
}
// Controla re-render: quando o campo de carga está em foco, bloqueia re-render
// para não destruir o input enquanto o usuário edita.
let _cargaInputAtivo = false;
let _cargaPendente = null;

function editarHorarioCargaDebounced(vid, ti, hhmm) {
  if (!hhmm) { editarHorarioCarga(vid, ti, ''); return; }
  // Salva o valor imediatamente no objeto em memória SEM re-renderizar
  const v = veiculos.find(x => x.id === vid);
  if (!v || !ultimoResultado) return;
  const viagens = (ultimoResultado[v.id] || []).filter(vi => !vi._vazio && vi.paradas?.length);
  const viagem = viagens[ti];
  if (!viagem) return;
  const min = parseHoraMin(hhmm);
  if (isNaN(min)) return;
  viagem.horarioCargaManualMin = min;
  recalcularTimingViagem(viagem, v);
  recalcularControleTempo();
  renderTemplateOperacao();
  // Agenda re-render apenas depois que o campo perder o foco
  _cargaPendente = { vid, ti, hhmm };
}

function finalizarEdicaoCarga(vid, ti) {
  if (!_cargaPendente || _cargaPendente.vid !== vid || _cargaPendente.ti !== ti) return;
  const { hhmm } = _cargaPendente;
  _cargaPendente = null;
  editarHorarioCarga(vid, ti, hhmm);
}

// Override manual do horário de carregamento de uma viagem
function editarHorarioCarga(vid, ti, hhmm) {
  const v = veiculos.find(x => x.id === vid);
  if (!v || !ultimoResultado) return;
  // ti é o índice no array FILTRADO (sem _vazio) — precisamos encontrar a viagem certa
  const viagens = (ultimoResultado[v.id] || []).filter(vi => !vi._vazio && vi.paradas?.length);
  const viagem  = viagens[ti]; // usa o array filtrado, igual ao render
  if (!viagem || viagem._vazio || !(viagem.paradas || []).length) return;

  if (!hhmm) {
    delete viagem.horarioCargaManualMin;
    delete viagem._alertaCargaManual;
  } else {
    const min = parseHoraMin(hhmm);
    if (isNaN(min)) return;
    viagem.horarioCargaManualMin = min;

    // Verifica sobreposição com viagem anterior — gera alerta mas NÃO bloqueia
    const idxNaLista = viagens.indexOf(viagem);
    if (idxNaLista > 0) {
      const viagemAnterior = viagens[idxNaLista - 1];
      // fimCicloMin da viagem anterior = tempoConsumidoMin relativo ao início
      const _jIniRaw = parseHoraMin(v.jornadaInicio || '06:00');
      const jIniMin  = isNaN(_jIniRaw) ? 360 : _jIniRaw;
      const iniAnterior = inicioViagemAbsMin(viagens, idxNaLista - 1, jIniMin, v.tempoPerdidoMin || 0, doisTurnos(v) ? 2 : 1);
      const fimAnteriorEstimado = iniAnterior + (viagemAnterior.tempoConsumidoMin || 0);
      const baseDay  = Math.floor(fimAnteriorEstimado / 1440) * 1440;
      let   alvoAbs  = baseDay + min;
      if (alvoAbs < fimAnteriorEstimado - 0.001) alvoAbs += 1440;
      const sobrepoe = alvoAbs < fimAnteriorEstimado - 0.001;
      viagem._alertaCargaManual = sobrepoe
        ? `⚠ Horário manual (${hhmm}) é anterior ao retorno estimado da viagem anterior. Verifique se há tempo suficiente.`
        : null;
    } else {
      viagem._alertaCargaManual = null;
    }
  }

  recalcularTimingViagem(viagem, v);
  recalcularControleTempo();
  renderResultado(ultimoResultado, ultimoControleTempo || {});
  renderTemplateOperacao();
}
function renderResultado(resultado, controleTempo={}) {
  // Garante que toda viagem real tenha petId (cobre ajustes manuais pós-otimização)
  // Recria o contador a partir do histórico ANTES de atribuir IDs — evita repetições
  _reconstruirPetSeqDoHistorico().then(() => {
    atribuirPetIds(resultado);
    _renderResultadoInterno(resultado, controleTempo);
  }).catch(() => {
    atribuirPetIds(resultado);
    _renderResultadoInterno(resultado, controleTempo);
  });
}
function _renderResultadoInterno(resultado, controleTempo={}) {
  let entregueVol=0, totalQuebras=0, veicsUsados=0;
  const filtroResCliente = valId('f-res-cliente');
  const filtroResCidade = valId('f-res-cidade');
  const filtroResTerminal = valId('f-res-terminal');
  const filtroResPlaca = valId('f-res-placa');
  const filtroResTransp = valId('f-res-transp');
  const rotasRaw = veiculos.map(v => ({v, viagens:resultado[v.id]}));
  const rotas = rotasRaw
    .filter(r => {
      const passaTerminal = !txt(filtroResTerminal) ||
        containsFiltro(r.v.terminal, filtroResTerminal) ||
        containsFiltro(baseVeiculoLabel(r.v), filtroResTerminal) ||
        (r.viagens || []).some(vi => (vi.paradas || []).some(p => containsFiltro(p.pedido?.terminal, filtroResTerminal)));
      return (
        containsFiltro(r.v.placa, filtroResPlaca) &&
        containsFiltro(r.v.transportadora, filtroResTransp) &&
        passaTerminal
      );
    })
    .map(r => ({
      ...r,
      viagens: (r.viagens || []).map(vi => ({
        ...vi,
        paradas: (vi.paradas || []).filter(p =>
          containsFiltro(p.pedido?.cliente, filtroResCliente) &&
          containsFiltro(p.pedido?.cidade, filtroResCidade) &&
          containsFiltro(p.pedido?.terminal, filtroResTerminal)
        )
      }))
    }));
  rotas.forEach(r => {
    if (r.viagens.some(vi=>vi.paradas.length>0)) veicsUsados++;
    r.viagens.forEach(vi => { vi.paradas.forEach(p=>entregueVol+=p.volumeTotal); totalQuebras+=vi.quebras.length; });
  });
  const totalCapDia = rotas.reduce((s,r) => s + (r.viagens.filter(vi => vi.paradas.length > 0).length * r.v.capacidade), 0);
  const pctUtil = totalCapDia>0 ? Math.round(entregueVol/totalCapDia*100) : 0;
  const totalTempoUsadoMin = veiculos.reduce((s,v) => s + (controleTempo[v.id]?.usadoMin || 0), 0);
  const totalTempoDispMin = veiculos.reduce((s,v) => s + (controleTempo[v.id]?.limiteMin || v.jornadaMin || 0), 0);
  const jornadaMediaH = veiculos.length ? (totalTempoDispMin / veiculos.length / 60) : 0;
  let html = `
    <div class="metrics-row">
      <div class="metric"><div class="metric-val">${veicsUsados}</div><div class="metric-lbl">Veículos</div></div>
      <div class="metric"><div class="metric-val">${jornadaMediaH.toFixed(1)} h</div><div class="metric-lbl">Jornada média</div></div>
      <div class="metric"><div class="metric-val">${entregueVol.toFixed(1)} m³</div><div class="metric-lbl">Vol. alocado</div></div>
      <div class="metric"><div class="metric-val">${pctUtil}%</div><div class="metric-lbl">Utiliz. frota</div></div>
    </div>`;
  html += `<div class="alert alert-info">Tempo consumido na operação: ${(totalTempoUsadoMin/60).toFixed(1)} h de ${(totalTempoDispMin/60).toFixed(1)} h de jornada disponível.</div>`;
  // Alerta de espera longa em campo (>4h por restrição de janela)
  // Indica que o motorista chegou muito antes da janela abrir — sinal de quebra ou reordenação
  {
    const LIMIAR_ESPERA_MIN = 240; // 4 horas
    const esperasLongas = [];
    rotas.forEach(r => {
      (r.viagens || []).forEach(vi => {
        (vi.paradas || []).forEach(pa => {
          const esp = pa.tempoEsperaRestricaoMin || 0;
          if (esp >= LIMIAR_ESPERA_MIN) {
            esperasLongas.push({
              placa: r.v.placa,
              cliente: pa.pedido?.cliente || '—',
              restricao: pa.pedido?.restricao || '—',
              esperaH: (esp / 60).toFixed(1),
            });
          }
        });
      });
    });
    if (esperasLongas.length) {
      const itens = esperasLongas.map(e =>
        `<li><strong>${e.cliente}</strong> — ${e.restricao} · placa ${e.placa} · <strong>${e.esperaH}h parado</strong> esperando janela</li>`
      ).join('');
      html += `<div class="alert alert-warn" style="margin-top:8px;">
        <strong>⏳ Espera longa em campo (${esperasLongas.length} parada(s)):</strong>
        O motorista chega muito antes da janela do cliente — considere quebrar o pedido
        ou reordenar as paradas para reduzir o tempo parado.
        <ul style="margin:6px 0 0 16px;font-size:11.5px;">${itens}</ul>
      </div>`;
    }
  }
  const excessoJornada = veiculos
    .map(v => {
      const usado = controleTempo[v.id]?.usadoMin || 0;
      const limite = controleTempo[v.id]?.limiteMin || v.jornadaMin || 0;
      return { v, excesso: Math.max(0, usado - limite) };
    })
    .filter(x => x.excesso > 0.001);
  if (excessoJornada.length) {
    const linhas = excessoJornada
      .sort((a,b) => b.excesso - a.excesso)
      .map(x => `${x.v.placa}: +${(x.excesso/60).toFixed(1)} h`);
    html += `<div class="alert alert-warn">⚠ Jornada excedida: ${linhas.join(', ')}. O pedido foi alocado mas a jornada do veículo foi ultrapassada.</div>`;
  }
  if (totalQuebras > 0) html += `<div class="alert alert-warn">⚠ ${totalQuebras} item(ns) de pedido divididos por limitação de capacidade.</div>`;
  else html += `<div class="alert alert-success">✓ Todos os itens alocados sem quebras de pedido.</div>`;
  // Avisos devem usar visão global (sem filtros da tela), para evitar falso
  // "sem alocação" ao filtrar cidade/cliente/terminal.
  const allocMapGlobal = new Map();
  rotasRaw.forEach(r => (r.viagens || []).forEach(vi => (vi.paradas || []).forEach(pa =>
    allocMapGlobal.set(pa.pedido.id, (allocMapGlobal.get(pa.pedido.id)||0) + pa.volumeTotal)
  )));
  const naoProgr = pedidos.filter(p => (allocMapGlobal.get(p.id)||0) < totalVolPedido(p) - 0.01);
  if (naoProgr.length) {
    // Cat 1 — dados de cadastro faltando (terminal ausente bloqueia; lat/lon=0 só bloqueia se não há centróide)
    const semCadastro     = naoProgr.filter(p => !p.terminal || (!p.terminal && latLonEfetivo(p).isCentroide === false && isNaN(latLonEfetivo(p).lat)));
    const semCadastroIds  = new Set(semCadastro.map(p => p.id));
    // Cat 2 — terminal especificado mas nenhum veículo da frota o atende
    const semVeicTerm     = naoProgr.filter(p =>
      !semCadastroIds.has(p.id) &&
      !veiculos.some(v => veiculoAtendeTerminal(v, p.terminal))
    );
    const semVeicTermIds  = new Set(semVeicTerm.map(p => p.id));
    // Cat 3 — veículo existe para o terminal mas tipo/Petronas bloqueia
    const semFrota        = naoProgr.filter(p => {
      if (semCadastroIds.has(p.id) || semVeicTermIds.has(p.id)) return false;
      const pool = veiculos
        .filter(v => !p.tiposCaminhao?.length || p.tiposCaminhao.includes(v.tipo))
        .filter(v => !p.identidadePetronas || !!v.identidadePetronas);
      return !pool.some(v => veiculoAtendeTerminal(v, p.terminal));
    });
    const semFrotaIds     = new Set(semFrota.map(p => p.id));
    // Cat 4 — frota compatível existe mas capacidade/jornada esgotou
    const semCap          = naoProgr.filter(p =>
      !semCadastroIds.has(p.id) && !semVeicTermIds.has(p.id) && !semFrotaIds.has(p.id)
    );
    // Agrupa cada categoria de alerta por terminal
    const alertaPorTerminal = (lista, labelFn) => {
      if (!lista.length) return '';
      const grupos = {};
      lista.forEach(p => {
        const t = p.terminal || '(sem terminal)';
        if (!grupos[t]) grupos[t] = [];
        grupos[t].push(labelFn(p));
      });
      return Object.entries(grupos)
        .map(([t, nomes]) => `<div style="margin-top:4px;"><b style="font-size:11px;">${t}</b>: ${nomes.join(', ')}</div>`)
        .join('');
    };
    if (semCadastro.length)
      html += `<div class="alert alert-warn">⚠ <b>Cadastro incompleto</b> — terminal ausente (pedido não pode ser roteado):
        ${alertaPorTerminal(semCadastro, p => `${p.cliente}${!p.terminal ? ' <em style="opacity:.7">(sem terminal)</em>' : ''}`)}</div>`;
    if (semVeicTerm.length)
      html += `<div class="alert alert-warn">⚠ <b>Terminal sem veículos</b> — nenhum veículo da frota atende o terminal do pedido:
        ${alertaPorTerminal(semVeicTerm, p => p.cliente)}</div>`;
    if (semFrota.length)
      html += `<div class="alert alert-warn">⚠ <b>Restrição de frota</b> — tipo de veículo ou identidade Petronas não disponível:
        ${alertaPorTerminal(semFrota, p => p.cliente)}</div>`;
    if (semCap.length)
      html += `<div class="alert alert-warn">⚠ <b>Sem capacidade</b> — frota compatível existe mas jornada/espaço esgotou:
        ${alertaPorTerminal(semCap, p => p.cliente)}</div>`;
  }
  // Aviso informativo: pedidos roteados com centróide da cidade (lat/lon ausentes no pedido)
  const comCentroide = pedidos.filter(p => {
    const c = latLonEfetivo(p);
    return c.isCentroide && (allocMapGlobal.get(p.id)||0) > 0;
  });
  if (comCentroide.length) {
    const nomes = comCentroide.map(p => `${p.cliente} (${p.cidade})`).join(', ');
    html += `<div class="alert alert-info">📍 <b>Coordenadas estimadas</b> — ${comCentroide.length} pedido(s) sem lat/lon roteados pelo centróide da cidade (distâncias aproximadas). Cadastre as coordenadas para maior precisão: ${nomes}.</div>`;
  }
  const pendComp = (ultimoItensOtimizacao || []).filter(it => (it.restante || 0) > 0.0001 && it.motivoNaoAlocado === 'compartimento');
  if (pendComp.length) {
    const chaves = [...new Set(pendComp.map(it => `${it.pedido?.cliente || '-'} (${it.produto})`))];
    html += `<div class="alert alert-warn">⚠ Falta de compartimento específico (carga parcial por compartimento não permitida). Acionar comercial para: ${chaves.join(', ')}.</div>`;
  }
  // ── Separa alertas de métricas do grid de rotas ──────────────────────────
  const htmlAlertas = html;
  html = '';
  // Agrupa por terminal do veículo (origem) — todos os veículos disponíveis aparecem
  const porTerminal = {};
  rotas.forEach(r => {
    if ((r.v.disponibilidade || 'Disponível') === 'Indisponível') return;
    const t = baseVeiculoLabel(r.v);
    if (!porTerminal[t]) porTerminal[t] = [];
    porTerminal[t].push(r);
  });
  const diaHoje = new Date().getDay();
  Object.entries(porTerminal).forEach(([terminal, rs]) => {
    // Alocados primeiro, sem viagem por último
    rs.sort((a, b) => {
      const aTemParadas = (a.viagens || []).some(vi => !vi._vazio && (vi.paradas || []).length > 0) ? 1 : 0;
      const bTemParadas = (b.viagens || []).some(vi => !vi._vazio && (vi.paradas || []).length > 0) ? 1 : 0;
      return bTemParadas - aTemParadas;
    });
    const tc = terminaisCad.find(t => t.nome === terminal);
    let horarioHoje = '';
    if (tc) {
      if (tc.diasAtivos.includes(diaHoje)) {
        const h  = tc.horarios?.[diaHoje];
        const ab = h?.abertura   || tc.aberturaPadrao;
        const fe = h?.fechamento || tc.fechamentoPadrao;
        if (ab && fe) horarioHoje = ` &nbsp;·&nbsp; <span style="font-size:10px;color:#4A6535;">Hoje: ${ab}–${fe}</span>`;
      } else {
        horarioHoje = ` &nbsp;·&nbsp; <span style="color:#B91C1C;font-size:10px;">Fechado hoje</span>`;
      }
    }
    const isLocked = lockedTerminals.has(terminal);
    html += `<div style="margin:14px 0 8px;">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        ${termTag(terminal)}
        <span style="font-size:11px;color:#4A6535;">${rs.length} veículo${rs.length>1?'s':''}</span>
        ${horarioHoje}
        <button onclick="toggleLockTerminal('${terminal.replace(/'/g,"\\'")}',this)"
          title="${isLocked ? 'Ajustes travados — clique para desbloquear e re-otimizar' : 'Clique para travar ajustes deste terminal (próxima otimização não altera)'}"
          style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;border:1.5px solid ${isLocked ? '#2563EB' : 'var(--border-dk)'};background:${isLocked ? '#EEF3FF' : 'var(--surface)'};color:${isLocked ? '#1E40AF' : 'var(--text-3)'};font-size:10px;font-weight:700;cursor:pointer;font-family:var(--font-cond);letter-spacing:.05em;transition:all .15s;">
          ${isLocked ? '🔒 TRAVADO' : '🔓 Travar'}
        </button>
        ${isLocked ? `<span style="font-size:10px;color:#1E40AF;font-style:italic;">próxima otimização preserva estes ajustes</span>` : ''}
      </div>
    </div>`;
    html += `<div class="rotas-grid">`;
    rs.forEach(({v, viagens}) => {
      // Card sem viagens: marcador _vazio (manual) ou array vazio (não alocado pelo otimizador)
      const temParadas = (viagens || []).some(vi => !vi._vazio && (vi.paradas || []).length > 0);
      if (!temParadas) {
        const contratoTag = (v.contrato || 'Dedicado') === 'Spot'
          ? `<span class="tag" style="font-size:9px;background:#FEF3C7;color:#92400E;border:1px solid #FCD34D;">Spot</span>`
          : `<span class="tag" style="font-size:9px;background:#EEF3FF;color:#1E40AF;border:1px solid #93C5FD;">Dedicado</span>`;
        html += `<div class="route-card"
          style="border:2px dashed var(--border-dk);opacity:0.75;"
          ondragenter="dragEnterCard(event,${v.id})"
          ondragleave="dragLeaveCard(event)"
          ondragover="dragOverCard(event,${v.id})"
          ondrop="dropNaCard(event,${v.id})">
          <div class="route-header" style="border-left:4px solid var(--border-dk);background:var(--bg);">
            <div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;">
              <span class="placa-badge" style="background:#6B7280;cursor:default;">${v.placa}</span>
              ${v.implemento ? `<span class="tag tag-gray" style="font-size:9px;">⊞ ${v.implemento}</span>` : ''}
              ${v.transportadora ? `<span class="tag tag-lime" style="font-size:9px;">${v.transportadora}</span>` : ''}
              <span class="tag tag-gray">${v.tipo} · ${(v.capacidade||0).toFixed(0)} m³</span>
              ${contratoTag}
              ${v.identidadePetronas ? `<span class="tag tag-yellow" style="font-size:9px;">⬡ ID Petronas</span>` : ''}
              <span class="tag tag-blue" style="font-size:9px;">Jornada: ${v.jornadaInicio||'06:00'}–${v.jornadaFim||'18:00'}</span>
            </div>
            <span class="tag tag-gray">Sem viagens</span>
          </div>
          <div class="route-body" style="min-height:70px;display:flex;align-items:center;justify-content:center;">
            <span style="font-size:11px;color:var(--text-3);font-family:var(--font-cond);letter-spacing:0.08em;text-transform:uppercase;">
              ↙ Arraste uma viagem ou entrega aqui
            </span>
          </div>
        </div>`;
        return;
      }
      const volDia  = viagens.reduce((s,vi)=>s+vi.paradas.reduce((ss,p)=>ss+p.volumeTotal,0),0);
      const ciclosUsados = viagens.filter(vi => vi.paradas.length > 0).length;
      const utilDia = (v.capacidade > 0 && ciclosUsados > 0) ? Math.round(volDia / (ciclosUsados * v.capacidade) * 100) : 0;
      const tempoDispVeic = controleTempo[v.id]?.limiteMin || v.jornadaMin || 0;
      // Recalcula tempo real usado somando as paradas (reflete ajustes manuais)
      let _cr = parseHoraMin(v.jornadaInicio || '06:00');
      if (isNaN(_cr)) _cr = 360;
      const _ini = _cr;
      viagens.filter(vi => vi.paradas.length).forEach((vi, vi_idx) => {
        if (vi_idx === 1) _cr += (v.tempoPerdidoMin || 0);
        if (doisTurnos(v) && vi_idx === 2) _cr += (v.tempoPerdidoMin || 0);
        _cr += vi.esperaTerminalMin || 0;
        vi.paradas.forEach((p, pi) => {
          const espOrig = p.tempoEsperaRestricaoMin || 0;
          const wal     = p.overnight ? (p.waitAfterLoadingMin || 0) : 0;
          const atraso  = (!p.overnight && pi===0 && espOrig>0 && (p.tempoCarregamentoMin||0)>0) ? espOrig : 0;
          const espVis  = p.overnight ? 0 : (espOrig - atraso);
          _cr += atraso + (p.tempoCarregamentoMin||0) + wal + (p.deslocCarregadoMin||0) + espVis + (p.tempoDescargaMin||0) + (p.deslocVazioMin||0);
        });
      });
      const tempoUsadoVeic = _cr - _ini;
      const excessoJornadaVeicMin = Math.max(0, tempoUsadoVeic - tempoDispVeic);
      let relogioMin = parseHoraMin(v.jornadaInicio || '06:00');
      if (isNaN(relogioMin)) relogioMin = 6 * 60;
      let _viagensRenderizadas = 0; // contador de viagens reais (com paradas) já processadas
      // ── Data base da jornada deste veículo ──────────────────────────────────
      // Usa _baseDataEntrega gravada na otimização (= hoje quando todos os pedidos
      // são futuros), garantindo que absMin=0 corresponda ao dia correto.
      // Fallback: dataEntregaLogistica do primeiro pedido (comportamento legado).
      const _fp = viagens.find(vi => vi.paradas.length)?.paradas[0];
      const _dl = _fp?.pedido?.dataEntregaLogistica;
      let jornadaBaseDate = null;
      if (ultimoResultado._baseDataEntrega) {
        const _bd = new Date(ultimoResultado._baseDataEntrega);
        if (!isNaN(_bd.getTime())) jornadaBaseDate = _bd;
      }
      if (!jornadaBaseDate && _dl) {
        const _pts = _dl.split('/');
        if (_pts.length >= 3) {
          const _dd = new Date(parseInt(_pts[2]), parseInt(_pts[1]) - 1, parseInt(_pts[0]));
          if (_fp.overnight) _dd.setDate(_dd.getDate() - 1);
          jornadaBaseDate = _dd;
        }
      }
      // Formata minuto absoluto (desde meia-noite da data base) como "DD/MM HH:MM"
      const fmtDT = (absMin) => {
        if (!jornadaBaseDate) return fmtHora(absMin);
        const d = new Date(jornadaBaseDate.getTime() + Math.floor(absMin / 1440) * 86400000);
        return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')} ${fmtHora(absMin)}`;
      };
      // ── Verificar violações de restrições ───────────────────────────────────
      // Usa as mesmas referências de objeto (p) como chaves → lookup no stop render
      const paradasComViolacao = new Map();
      const viagensComViolacao = new Set();
      let _rv = parseHoraMin(v.jornadaInicio || '06:00');
      if (isNaN(_rv)) _rv = 360;
      viagens.filter(vi => vi.paradas?.length && !vi._vazio).forEach((vi, vi_idx) => {
        if (vi_idx === 1) _rv += (v.tempoPerdidoMin || 0);
        if (doisTurnos(v) && vi_idx === 2) _rv += (v.tempoPerdidoMin || 0);
        _rv += vi.esperaTerminalMin || 0;
        vi.paradas.forEach((p, pi) => {
          const _espOrig = p.tempoEsperaRestricaoMin || 0;
          const _wal     = p.overnight ? (p.waitAfterLoadingMin || 0) : 0;
          const _atraso  = (!p.overnight && pi === 0 && _espOrig > 0 && (p.tempoCarregamentoMin || 0) > 0) ? _espOrig : 0;
          const _chegada = _rv + _atraso + (p.tempoCarregamentoMin || 0) + _wal + (p.deslocCarregadoMin || 0);
          const _espVis  = p.overnight ? 0 : (_espOrig - _atraso);
          _rv = _chegada + _espVis + (p.tempoDescargaMin || 0) + (p.deslocVazioMin || 0);
          const viols = [];
          // Identidade Petronas
          if (p.pedido?.identidadePetronas && !v.identidadePetronas)
            viols.push({ tipo: 'petronas', label: 'Exige ID Petronas — veículo não possui' });
          // Tipo de veículo
          if (p.pedido?.tiposCaminhao?.length && !p.pedido.tiposCaminhao.includes(v.tipo))
            viols.push({ tipo: 'veiculo', label: `Exige ${p.pedido.tiposCaminhao.join('/')} — veículo é ${v.tipo}` });
          // Horário de recebimento
          if (p.pedido?.restricao) {
            const _pts = (p.pedido.restricao.match(/\d{1,2}:\d{2}/g) || []).map(parseHoraMin);
            if (_pts.length >= 2 && !isNaN(_pts[0]) && !isNaN(_pts[1])) {
              const _chegHora = _chegada % 1440;
              if (_chegHora > _pts[1])
                viols.push({ tipo: 'horario', label: `Chegada ${fmtHora(_chegada)} após fechamento ${fmtHora(_pts[1])}` });
            }
          }
          if (viols.length) { paradasComViolacao.set(p, viols); viagensComViolacao.add(vi); }
        });
      });
      const totalViolacoes = paradasComViolacao.size;
      html += `<div class="route-card"
        ondragenter="dragEnterCard(event,${v.id})"
        ondragleave="dragLeaveCard(event)"
        ondragover="dragOverCard(event,${v.id})"
        ondrop="dropNaCard(event,${v.id})">
        <div class="route-header" style="${utilEstilo(utilDia)}">
          <div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;">
            <span style="display:inline-flex;align-items:stretch;border-radius:5px;overflow:hidden;">
              <span class="placa-badge" style="background:${utilPlacaBg(utilDia)};border-radius:5px 0 0 5px;cursor:pointer;" onclick="abrirMapaVeiculo(${v.id})" title="Ver mapa completo — todas as viagens">${v.placa} ⊙</span>
              <button onclick="abrirTrocaPlaca(${v.id},this,event)" title="Trocar veículo" style="background:${utilPlacaBg(utilDia)};color:rgba(255,255,255,0.85);border:none;border-left:1px solid rgba(255,255,255,0.25);padding:0 8px;cursor:pointer;font-size:12px;border-radius:0 5px 5px 0;transition:background .15s;" onmouseenter="this.style.background='rgba(0,0,0,0.25)'" onmouseleave="this.style.background='${utilPlacaBg(utilDia)}'">▾</button>
            </span>
            ${v.implemento ? `<span class="tag tag-gray" style="font-size:9px;">⊞ ${v.implemento}</span>` : ''}
            ${v.transportadora ? `<span class="tag tag-lime" style="font-size:9px;">${v.transportadora}</span>` : ''}
            <span class="tag tag-gray">${v.tipo} · ${v.capacidade.toFixed(0)} m³</span>
            ${(v.contrato||'Dedicado') === 'Spot'
              ? `<span class="tag" style="font-size:9px;background:#FEF3C7;color:#92400E;border:1px solid #FCD34D;">Spot</span>`
              : `<span class="tag" style="font-size:9px;background:#EEF3FF;color:#1E40AF;border:1px solid #93C5FD;">Dedicado</span>`}
            ${v.identidadePetronas ? `<span class="tag tag-yellow" style="font-size:9px;">⬡ ID Petronas</span>` : ''}
            <span class="tag tag-blue" style="font-size:9px;">Jornada: ${v.jornadaInicio||'06:00'}–${v.jornadaFim||'18:00'}</span>
          </div>
          <div style="display:flex;align-items:center;gap:6px;">
            <span style="font-size:11px;color:#4A6535;">${volDia.toFixed(1)} m³</span>
            <span style="font-size:11px;color:#4A6535;">${(tempoUsadoVeic/60).toFixed(1)}/${(tempoDispVeic/60).toFixed(1)} h</span>
            ${excessoJornadaVeicMin > 0.001 ? `<span class="tag tag-red">+${(excessoJornadaVeicMin/60).toFixed(1)}h</span>` : ''}
            <span class="tag ${utilTag(utilDia)}" style="${utilTagStyle(utilDia)}">${utilDia}%</span>
            ${totalViolacoes > 0 ? `<span class="tag tag-red" style="font-size:9px;letter-spacing:0.04em;" title="Existem restrições violadas — veja os alertas ⚠ nas entregas">⚠ ${totalViolacoes} restrição${totalViolacoes > 1 ? 'ões' : ''}</span>` : ''}
          </div>
        </div>
        <div class="route-body">`;
      viagens.forEach((viagem, ti) => {
        if (!viagem.paradas.length) return;
        // Aplica pausa de refeição antes da 2ª viagem em diante
        if (_viagensRenderizadas === 1) relogioMin += (v.tempoPerdidoMin || 0);
        if (doisTurnos(v) && _viagensRenderizadas === 2) relogioMin += (v.tempoPerdidoMin || 0);
        _viagensRenderizadas++;
        const volV  = viagem.paradas.reduce((s,p)=>s+p.volumeTotal,0);
        const utilV = Math.round(volV/v.capacidade*100);
        const inicioCicloMin = relogioMin;
        const esperaTermRender = viagem.esperaTerminalMin || 0;
        if (esperaTermRender > 0) relogioMin += esperaTermRender; // avança até abertura do terminal
        const _temOverride = viagem.horarioCargaManualMin !== undefined;
        // Se o usuário definiu horário manual, ajusta relogioMin para refletir esse horário
        // (preserva o horário salvo no re-render, evitando que o campo volte ao calculado)
        if (_temOverride && !isNaN(viagem.horarioCargaManualMin)) {
          // Usa o horário manual diretamente sem baseDay para não avançar dia
          relogioMin = viagem.horarioCargaManualMin;
        }
        const _relogioMinCarga = relogioMin; // horário exato de chegada ao terminal para carga
        const paradasComHorario = viagem.paradas.map((p, idxParada) => {
          const esperaOriginalMin = p.tempoEsperaRestricaoMin || 0;
          const waitAfterLoad = p.overnight ? (p.waitAfterLoadingMin || 0) : 0;
          // Com override manual o caminhão carrega no horário definido pelo usuário;
          // a espera por janela do cliente aparece como esperaVisível (no cliente), não atrasa a carga.
          const atrasoCargaMin = (!p.overnight && idxParada === 0 && esperaOriginalMin > 0 && (p.tempoCarregamentoMin || 0) > 0 && !_temOverride)
            ? esperaOriginalMin
            : 0;
          const esperaVisivelMin = p.overnight ? 0 : (esperaOriginalMin - atrasoCargaMin);
          const inicioCargaMin = relogioMin + atrasoCargaMin;
          // Carregamento apenas na 1ª parada (stops posteriores não recarregam no terminal)
          const tempoCarregaEfetivo = idxParada === 0 ? (p.tempoCarregamentoMin || 0) : 0;
          const fimCargaMin = inicioCargaMin + tempoCarregaEfetivo;
          // overnight: aguarda na base após carga antes de partir
          const chegadaEntregaMin = fimCargaMin + waitAfterLoad + (p.deslocCarregadoMin || 0);
          const inicioDescargaMin = chegadaEntregaMin + esperaVisivelMin;
          const fimDescargaMin = inicioDescargaMin + (p.tempoDescargaMin || 0);
          const retornoTerminalMin = fimDescargaMin + (p.deslocVazioMin || 0);
          relogioMin = retornoTerminalMin;
          return { p, inicioCargaMin, fimCargaMin, chegadaEntregaMin, inicioDescargaMin, fimDescargaMin, retornoTerminalMin, esperaVisivelMin, waitAfterLoad };
        });
        const fimCicloMin = relogioMin;
        viagem._inicioCargaMin = _relogioMinCarga;
        const label = `Viagem ${ti+1}${viagem.petId ? ` · ${viagem.petId}` : ''}`;
        const esperaTagHtml = esperaTermRender > 0
          ? `<span style="font-size:11px;color:#92400E;background:#FFFBEB;border:1px solid #FCD34D;border-radius:4px;padding:2px 7px;">⏳ Espera terminal ${fmtDT(inicioCicloMin)}→${fmtDT(inicioCicloMin+esperaTermRender)} (${Math.round(esperaTermRender)} min)</span>`
          : '';
        // Indica que o usuário ajustou manualmente o trajeto desta viagem no mapa
        const kmAjustadoTagHtml = viagem._kmAjustado != null
          ? `<span style="font-size:11px;color:#3730A3;background:rgba(79,70,229,0.1);border:1px solid rgba(79,70,229,0.3);border-radius:4px;padding:2px 7px;" title="Rota ajustada manualmente no mapa">📏 ${viagem._kmAjustado.toFixed(1)} km (ajustada)</span>`
          : '';
        const _temViol = viagensComViolacao.has(viagem);
        html += `<div class="viagem-header viagem-draggable"
          draggable="true"
          ondragstart="iniciarDragViagem(${v.id},${ti},this)"
          ondragend="fimDragViagem(this)"
          ondragenter="dragEnterViagem(event,${v.id},${ti})"
          ondragleave="dragLeaveViagem(event)"
          ondragover="dragOverViagem(event,${v.id},${ti})"
          ondrop="dropNaViagem(event,${v.id},${ti})"
          style="${_temViol ? 'background:#FEE2E2;border-left:3px solid #DC2626;' : utilViagBg(utilV)}"
          title="Arraste para mover esta viagem para outro veículo · Solte uma entrega aqui para adicioná-la a esta viagem">
          <span class="drag-handle">⠿</span>
          <span class="tag ${_temViol ? 'tag-red' : 'tag-blue'}">${label}</span>
          <span style="font-size:12px;color:#4A6535;">${volV.toFixed(1)} / ${v.capacidade} m³</span>
          <span style="font-size:12px;color:#4A6535;">${fmtDT(inicioCicloMin+esperaTermRender)} → ${fmtDT(fimCicloMin)}</span>
          ${esperaTagHtml}
          <span style="font-size:12px;color:#4A6535;">${(viagem.tempoConsumidoMin||0).toFixed(0)} min</span>
          ${kmAjustadoTagHtml}
          <button class="btn btn-sm" style="padding:2px 8px;" onclick="abrirModalViagem(${v.id},${ti})">Mapa</button>
          <span class="tag ${utilTag(utilV)}" style="margin-left:auto;${utilTagStyle(utilV)}">${utilV}%</span>
        </div>`;
        // ── Motorista da viagem (calculado pelo horário de carregamento) ─────────
        (() => {
          const cargaMin = paradasComHorario[0]?.inicioCargaMin ?? inicioCicloMin;
          const hod = ((Math.round(cargaMin) % 1440) + 1440) % 1440;
          // Usa jornadaInicio do veículo como limiar diurno/noturno — não fixo 06:00.
          // Ex: veículo com jornada 05:00-21:00: carga às 05:00 é diurna, não noturna.
          const _jornadaIniHod = parseHoraMin(v.jornadaInicio || '06:00') || 360;
          const ehNoturno = hod >= 18 * 60 || hod < _jornadaIniHod;
          const _petIdViagem = viagem.petId || '';
          const nomeMotor = (_petIdViagem && _motoristasOverride[_petIdViagem] !== undefined)
            ? _motoristasOverride[_petIdViagem]
            : (ehNoturno ? (v.motoristaNt||'') : (v.motoristaDiurno||''));
          const onInput = `_editMotoristaViagem('${_petIdViagem}',this.value)`;
            const cargaHHMM = (() => {
            const m = Math.round(cargaMin); const md = ((m%1440)+1440)%1440;
            return `${String(Math.floor(md/60)).padStart(2,'0')}:${String(md%60).padStart(2,'0')}`;
          })();
          const temOverride = viagem.horarioCargaManualMin !== undefined;
          html += `<div style="padding:3px 10px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;${utilViagBg(utilV)}border-bottom:1px solid var(--border);filter:brightness(0.95);">
            <span style="font-size:10px;color:var(--text-3);white-space:nowrap;">Motorista</span>
            <input value="${nomeMotor.replace(/"/g,'&quot;')}" placeholder="—"
              oninput="${onInput}"
              style="font-size:11px;padding:1px 6px;border:1px solid var(--border);border-radius:4px;width:190px;background:var(--surface);color:var(--text);font-weight:500;" />
            <span style="font-size:10px;color:var(--text-3);white-space:nowrap;margin-left:6px;">${temOverride ? '✏️' : '🕐'} Carga</span>
            <input type="time" value="${cargaHHMM}"
              title="${temOverride ? 'Horário personalizado — clique em Restaurar para voltar ao calculado' : 'Horário calculado — edite para personalizar'}"
              oninput="editarHorarioCargaDebounced(${v.id},${ti},this.value)"
              onchange="editarHorarioCargaDebounced(${v.id},${ti},this.value)"
              onblur="finalizarEdicaoCarga(${v.id},${ti})"
              style="font-size:12px;padding:2px 6px;border:2px solid ${temOverride ? '#16a34a' : 'var(--border)'};border-radius:5px;background:${temOverride ? 'rgba(22,163,74,0.08)' : 'var(--surface)'};color:${temOverride ? '#15803d' : 'var(--text)'};font-weight:${temOverride ? '700' : '500'};width:96px;cursor:pointer;" />
            ${temOverride
              ? `<button onclick="editarHorarioCarga(${v.id},${ti},'')" title="Restaurar horário calculado"
                  style="font-size:10px;padding:2px 8px;border:1px solid #16a34a;border-radius:5px;background:transparent;color:#16a34a;cursor:pointer;font-weight:600;">↺ Restaurar</button>`
              : `<span style="font-size:9px;color:var(--text-3);font-style:italic;">editável</span>`}
          </div>
          ${viagem._alertaCargaManual ? `<div style="padding:4px 10px 5px;display:flex;align-items:center;gap:6px;font-size:11px;color:#92400E;background:#FFFBEB;border-bottom:1px solid #FCD34D;">⚠️ ${viagem._alertaCargaManual.replace('⚠ ','')}</div>` : ''}`;
        })();
        paradasComHorario.forEach(({p, inicioCargaMin, fimCargaMin, chegadaEntregaMin, inicioDescargaMin, fimDescargaMin, retornoTerminalMin, esperaVisivelMin, waitAfterLoad},i) => {
          const temQuebra = p.itens.some(it => !it.completo);
          const tiposStop = p.pedido.tiposCaminhao && p.pedido.tiposCaminhao.length
            ? p.pedido.tiposCaminhao.map(t=>`<span class="tag tag-purple" style="font-size:9px;">${t}</span>`).join(' ')
            : '';
          const retornoTexto = (p.deslocVazioMin || 0) > 0 ? fmtDT(retornoTerminalMin) : 'segue rota';
          const origemDesloc = p.origemDeslocamento || 'Terminal';
          const cronogramaPartes = [];
          if (fimCargaMin > inicioCargaMin) {
            cronogramaPartes.push(`Carga ${fmtDT(inicioCargaMin)}-${fmtHora(fimCargaMin)}`);
          }
          if (p.overnight && waitAfterLoad > 0) {
            const saidaTerminalMin = fimCargaMin + waitAfterLoad;
            cronogramaPartes.push(`Aguarda partida ${fmtDT(fimCargaMin)} → ${fmtDT(saidaTerminalMin)}`);
          }
          cronogramaPartes.push(`Chegada ${fmtDT(chegadaEntregaMin)}`);
          if (!p.overnight && (esperaVisivelMin || 0) > 0) {
            cronogramaPartes.push(`Espera ${(esperaVisivelMin||0).toFixed(0)} min`);
          }
          cronogramaPartes.push(`Fim descarga ${fmtDT(fimDescargaMin)}`);
          if ((p.deslocVazioMin || 0) > 0) {
            cronogramaPartes.push(`Retorno ${retornoTexto}`);
          }
          const _stopViol = paradasComViolacao.has(p);
          html += `<div class="stop stop-draggable"
            draggable="true"
            ondragstart="iniciarDragParada(${v.id},${ti},${i},this)"
            ondragend="fimDragParada(this)"
            ondragenter="dragEnterStop(event,${v.id},${ti},${i})"
            ondragleave="dragLeaveStop(event)"
            ondragover="dragOverStop(event,${v.id},${ti},${i})"
            ondrop="dropNaStop(event,${v.id},${ti},${i})"
            title="Arraste para reordenar dentro da viagem ou mover para outra"
            style="${_stopViol ? 'background:#FFF1F2;' : utilStopBg(utilV)}">
            <span class="stop-drag-handle">⠿</span>
            <div class="stop-num">${i+1}</div>
            <div style="flex:1;">
              <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:4px;">
                ${p.pedido.codigoSAP ? `<span class="sap-code">SAP ${p.pedido.codigoSAP}</span>` : ''}
                <span style="font-family:var(--font-cond);font-weight:700;font-size:13px;">${p.pedido.cliente}</span>
                <span style="font-size:12px;color:#4A6535;">${p.pedido.cidade}</span>
                <span style="font-size:11px;font-weight:700;color:#1D4ED8;background:#DBEAFE;border:1px solid #93C5FD;border-radius:4px;padding:2px 8px;white-space:nowrap;">⏱ ${fmtDT(chegadaEntregaMin)}</span>
                ${latLonEfetivo(p.pedido).isCentroide ? `<span class="tag" style="font-size:9px;background:#FEF9C3;color:#854D0E;border:1px solid #FDE68A;" title="Coordenadas não cadastradas — usando centróide da cidade para distância">📍 ~centróide</span>` : ''}
                ${p.pedido.terminal ? termTag(p.pedido.terminal) : ''}
                ${temQuebra ? '<span class="tag tag-red">quebra</span>' : ''}
                ${p.pedido.restricao ? `<span class="tag tag-yellow">${p.pedido.restricao}</span>` : ''}
                ${p.pedido.identidadePetronas ? `<span class="tag tag-yellow" style="font-size:9px;">⬡ Exige ID Petronas</span>` : ''}
                ${tiposStop}
                <span class="pill" style="color:var(--pet-green);border-color:#C4E87A;margin-left:2px;">Vol. alocado: ${p.volumeTotal.toFixed(1)} m³</span>
                ${(paradasComViolacao.get(p)||[]).map(viol => `<span class="tag tag-red" style="font-size:9px;cursor:help;" title="${viol.label}">⚠ ${viol.tipo==='petronas'?'ID Petronas':viol.tipo==='veiculo'?'Tipo veíc.':'Horário'}</span>`).join('')}
              </div>
              <div class="pills">
                <span class="pill">${cronogramaPartes.join(' · ')}</span>
              </div>
            </div>
          </div>`;
        });
      });
      html += '</div>'; // fecha route-body
      html += '</div>'; // fecha route-card
    });
    html += `</div>`; // fecha rotas-grid
  });
  if (!rotas.some(r=>r.viagens.some(vi=>vi.paradas.length>0)))
    html = '<div class="empty">Nenhuma rota gerada. Verifique os alertas acima ou use alocação manual.</div>';
  document.getElementById('resultado-content').innerHTML = htmlAlertas + html;
}
// ═══════════════════════════════════════════════════════════════════════════
// SHAREPOINT
// ═══════════════════════════════════════════════════════════════════════════
function spToggleConfig() {
  const panel = document.getElementById('sp-config-panel');
  const btn   = document.getElementById('sp-toggle-btn');
  const aberto = !panel.classList.contains('hidden');
  panel.classList.toggle('hidden', aberto);
  btn.textContent = aberto ? 'Configurar ▾' : 'Fechar ▴';
}
function spSetStatus(tipo, msg) {
  const el = document.getElementById('sp-status');
  if (!el) return;
  el.className = 'sp-status sp-status-' + tipo;
  el.textContent = msg;
}
function spParseJson(str, fallback) {
  try { return JSON.parse(str || ''); } catch { return fallback; }
}
// Mapeamentos SP → local ──────────────────────────────────────────────────────
function spMapTerminal(item) {
  return {
    id: item.ID, spId: item.ID,
    nome:             item.Title              || '',
    cidade:           item.Cidade             || '',
    distribuidora:    item.Distribuidora      || item.CIA || '',
    lat:              parseFloat(item.Latitude)  || 0,
    lon:              parseFloat(item.Longitude) || 0,
    fuso:             item.Fuso               || '-3',
    aberturaPadrao:   item.AberturaPadrao     || '06:00',
    fechamentoPadrao: item.FechamentoPadrao   || '18:00',
    diasAtivos:       spParseJson(item.DiasAtivos, [1,2,3,4,5]),
    horarios:         spParseJson(item.Horarios, {}),
  };
}
function spMapCliente(item) {
  return {
    id: item.ID, spId: item.ID,
    codigoSAP:        item.CodigoSAP          || '',
    nome:             item.Title              || '',
    cidade:           item.Cidade             || '',
    lat:              parseFloat(item.Latitude)  || 0,
    lon:              parseFloat(item.Longitude) || 0,
    restricaoHorario: item.RestricaoHorario   || '',
    tiposCaminhao:    (item.TiposCaminhao || '').split(';').filter(Boolean),
    observacoes:      item.Observacoes        || '',
  };
}
function spMapPedido(item) {
  return {
    id: item.ID, spId: item.ID,
    codigoSAP:    item.CodigoSAP          || '',
    cliente:      item.Title              || '',
    cidade:       item.Cidade             || '',
    lat:          parseFloat(item.Latitude)  || 0,
    lon:          parseFloat(item.Longitude) || 0,
    terminal:     item.Terminal           || '',
    restricao:    item.RestricaoHorario   || null,
    tiposCaminhao:(item.TiposCaminhao || '').split(';').filter(Boolean),
    produtos:     spParseJson(item.Produtos, []),
  };
}
// Mapeamentos local → SP ──────────────────────────────────────────────────────
function spLocalTerminal(t) {
  return {
    Title:            t.nome,
    Cidade:           t.cidade,
    Distribuidora:    t.distribuidora || '',
    Latitude:         t.lat,
    Longitude:        t.lon,
    Fuso:             t.fuso,
    AberturaPadrao:   t.aberturaPadrao,
    FechamentoPadrao: t.fechamentoPadrao,
    DiasAtivos:       JSON.stringify(t.diasAtivos),
    Horarios:         JSON.stringify(t.horarios),
  };
}
function spLocalCliente(c) {
  return {
    Title:            c.nome,
    CodigoSAP:        c.codigoSAP,
    Cidade:           c.cidade,
    Latitude:         c.lat,
    Longitude:        c.lon,
    RestricaoHorario: c.restricaoHorario || '',
    TiposCaminhao:    (c.tiposCaminhao || []).join(';'),
    Observacoes:      c.observacoes || '',
  };
}
function spLocalPedido(p) {
  return {
    Title:            p.cliente,
    CodigoSAP:        p.codigoSAP || '',
    Cidade:           p.cidade,
    Latitude:         p.lat,
    Longitude:        p.lon,
    Terminal:         p.terminal || '',
    RestricaoHorario: p.restricao || '',
    TiposCaminhao:    (p.tiposCaminhao || []).join(';'),
    Produtos:         JSON.stringify(p.produtos),
  };
}
// Operações de leitura ────────────────────────────────────────────────────────
async function spCarregar(entidade) {
  const siteUrl = document.getElementById('sp-url').value.trim().replace(/\/$/, '');
  const lista   = document.getElementById('sp-lista-' + entidade).value.trim();
  if (!siteUrl || !lista) { alert('Informe a URL do site e o nome da lista.'); return; }
  spSetStatus('sync', `Carregando ${lista}…`);
  try {
    const url  = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lista)}')/items?$top=5000`;
    const resp = await fetch(url, {
      headers: { Accept: 'application/json;odata=verbose' },
      credentials: 'same-origin',
    });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err?.error?.message?.value || `HTTP ${resp.status}`);
    }
    const items = (await resp.json()).d.results;
    if (entidade === 'terminais') {
      terminaisCad = items.map(spMapTerminal);
      renderTerminais(); atualizarDropdownsTerminais();
    } else if (entidade === 'clientes') {
      clientes = items.map(spMapCliente);
      renderClientes(); atualizarDropdownsClientes();
    } else if (entidade === 'pedidos') {
      pedidos = items.map(spMapPedido);
      renderPedidos();
    }
    spSetStatus('ok', `✓ ${items.length} ${entidade} carregados`);
  } catch (e) {
    spSetStatus('erro', `✕ ${entidade}: ${e.message}`);
  }
}
async function spCarregarTudo() {
  for (const ent of ['terminais', 'clientes', 'pedidos']) {
    await spCarregar(ent);
  }
}
// Operações de escrita ────────────────────────────────────────────────────────
async function spObterDigest(siteUrl) {
  const resp = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: { Accept: 'application/json;odata=verbose' },
    credentials: 'same-origin',
  });
  if (!resp.ok) throw new Error(`Falha ao obter digest: HTTP ${resp.status}`);
  return (await resp.json()).d.GetContextWebInformation.FormDigestValue;
}
async function spObterTipoLista(siteUrl, lista) {
  const resp = await fetch(
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lista)}')?$select=ListItemEntityTypeFullName`,
    { headers: { Accept: 'application/json;odata=verbose' }, credentials: 'same-origin' }
  );
  if (!resp.ok) throw new Error(`Lista "${lista}" não encontrada: HTTP ${resp.status}`);
  return (await resp.json()).d.ListItemEntityTypeFullName;
}
async function spSalvar(entidade, itens, mapFn) {
  const siteUrl = document.getElementById('sp-url').value.trim().replace(/\/$/, '');
  const lista   = document.getElementById('sp-lista-' + entidade).value.trim();
  if (!siteUrl || !lista) { alert('Informe a URL do site e o nome da lista.'); return; }
  if (!itens.length) { spSetStatus('ok', `Nenhum item para salvar em ${lista}`); return; }
  spSetStatus('sync', `Salvando ${lista}…`);
  try {
    const digest = await spObterDigest(siteUrl);
    const tipo   = await spObterTipoLista(siteUrl, lista);
    const baseUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lista)}')/items`;
    const hdrsBase = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest,
    };
    let criados = 0, atualizados = 0;
    for (const item of itens) {
      const body = JSON.stringify({ __metadata: { type: tipo }, ...mapFn(item) });
      if (item.spId) {
        await fetch(`${baseUrl}(${item.spId})`, {
          method: 'POST', credentials: 'same-origin',
          headers: { ...hdrsBase, 'X-HTTP-Method': 'MERGE', 'If-Match': '*' },
          body,
        });
        atualizados++;
      } else {
        const r  = await fetch(baseUrl, { method: 'POST', credentials: 'same-origin', headers: hdrsBase, body });
        const rd = await r.json();
        item.spId = rd.d?.ID;
        criados++;
      }
    }
    spSetStatus('ok', `✓ ${lista}: ${criados} criados, ${atualizados} atualizados`);
  } catch (e) {
    spSetStatus('erro', `✕ ${entidade}: ${e.message}`);
  }
}
// ═══════════════════════════════════════════════════════════════════════════
// HISTÓRICO DE ROTEIRIZAÇÕES (File System Access API)
// ═══════════════════════════════════════════════════════════════════════════
const HIST_IDB_NAME  = 'nexta-historico';
const HIST_IDB_STORE = 'handles';
const HIST_IDB_KEY   = 'dir-handle';
function _histAbrirIDB() {
  return new Promise((res, rej) => {
    const req = indexedDB.open(HIST_IDB_NAME, 1);
    req.onupgradeneeded = e => e.target.result.createObjectStore(HIST_IDB_STORE);
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}
async function _histSalvarHandle(handle) {
  const db = await _histAbrirIDB();
  return new Promise((res, rej) => {
    const tx = db.transaction(HIST_IDB_STORE, 'readwrite');
    tx.objectStore(HIST_IDB_STORE).put(handle, HIST_IDB_KEY);
    tx.oncomplete = res; tx.onerror = e => rej(e.target.error);
  });
}
async function _histLerHandle() {
  const db = await _histAbrirIDB();
  return new Promise((res, rej) => {
    const tx  = db.transaction(HIST_IDB_STORE, 'readonly');
    const req = tx.objectStore(HIST_IDB_STORE).get(HIST_IDB_KEY);
    req.onsuccess = e => res(e.target.result || null);
    req.onerror   = e => rej(e.target.error);
  });
}
async function _histGarantirPermissao() {
  if (!dirHandleHistorico) return false;
  let perm = await dirHandleHistorico.queryPermission({ mode: 'readwrite' });
  if (perm !== 'granted') perm = await dirHandleHistorico.requestPermission({ mode: 'readwrite' });
  return perm === 'granted';
}
async function recuperarHandleHistorico() {
  try {
    const handle = await _histLerHandle();
    if (!handle) return;
    const perm = await handle.queryPermission({ mode: 'readwrite' });
    dirHandleHistorico = handle;
    const el = document.getElementById('hist-pasta-path');
    if (perm === 'granted') {
      el.textContent = handle.name;
      popularDropdownRoteirizacoes();
      popularSeletorResumoDia().catch(()=>{});
    } else {
      el.textContent = handle.name + '  (clique em "Atualizar lista" para reautorizar)';
    }
  } catch(e) { /* IDB indisponível ou handle inválido */ }
}
async function selecionarPastaHistorico() {
  if (!window.showDirectoryPicker) {
    alert('Seu navegador não suporta File System Access API.\nUse Chrome ou Edge via HTTPS.');
    return;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    dirHandleHistorico = handle;
    _histMetaCache = {}; // limpa cache ao trocar de pasta
    await _histSalvarHandle(handle);
    document.getElementById('hist-pasta-path').textContent = handle.name;
    await carregarListaHistorico();
    await popularDropdownRoteirizacoes();
    popularSeletorResumoDia().catch(()=>{});
  } catch(e) {
    if (e.name !== 'AbortError') alert('Erro ao selecionar pasta: ' + e.message);
  }
}
async function salvarNoHistorico(silencioso = false) {
  if (!ultimoResultado) { if (!silencioso) alert('Execute a otimização antes de salvar.'); return; }
  if (!dirHandleHistorico) {
    if (silencioso) return; // não abre diálogo no modo automático
    showTab('historico');
    await selecionarPastaHistorico();
    if (!dirHandleHistorico) return;
  }
  if (!await _histGarantirPermissao()) { if (!silencioso) alert('Permissão de escrita negada.'); return; }
  let totalViagens = 0, totalVolume = 0, totalPedidos = 0, totalRotas = 0;
  const terminaisUsados = new Set();
  veiculos.forEach(v => {
    const viagens = (ultimoResultado[v.id] || []).filter(vi => vi.paradas && vi.paradas.length > 0);
    if (!viagens.length) return;
    totalRotas++;
    totalViagens += viagens.length;
    viagens.forEach(vi => {
      vi.paradas.forEach(p => {
        totalVolume  += p.volumeTotal || 0;
        totalPedidos++;
      });
      if (v.terminal) terminaisUsados.add(v.terminal);
    });
  });
  const datasEntrega = [...new Set(pedidos.map(p => p.dataEntregaLogistica).filter(Boolean))];
  // Usuário logado vem sempre do app principal (Firebase Auth) — é a única
  // forma de acessar o Roteirizador (ver setTab() / controle de acesso).
  const usuarioLogado = (S && S.user && window.USERS_DB && window.USERS_DB[S.user])
    ? (window.USERS_DB[S.user].name || window.USERS_DB[S.user].login || S.user)
    : (S?.user || 'Desconhecido');
  const snapshot = {
    versao: 1,
    savedAt: new Date().toISOString(),
    salvoPor: usuarioLogado,
    datasEntrega,
    resumo: {
      totalRotas,
      totalViagens,
      totalPedidos,
      totalVolume_m3: Math.round(totalVolume * 10) / 10,
      terminaisUsados: [...terminaisUsados],
    },
    pedidos:       JSON.parse(JSON.stringify(pedidos)),
    terminais:     JSON.parse(JSON.stringify(terminaisCad)),
    veiculos:      JSON.parse(JSON.stringify(veiculos)),
    resultado:     JSON.parse(JSON.stringify(ultimoResultado)),
    controleTempo: JSON.parse(JSON.stringify(ultimoControleTempo || {})),
  };
  const now = new Date();
  const p2  = n => String(n).padStart(2, '0');
  const filename = `${now.getFullYear()}${p2(now.getMonth()+1)}${p2(now.getDate())}_${p2(now.getHours())}${p2(now.getMinutes())}_rotas.json`;
  try {
    const fh = await dirHandleHistorico.getFileHandle(filename, { create: true });
    const ws = await fh.createWritable();
    await ws.write(JSON.stringify(snapshot, null, 2));
    await ws.close();
    _atualizarPetSeq(ultimoResultado); // persiste contador apenas ao salvar
    delete _histMetaCache[filename]; // invalida cache para releitura na próxima listagem
    if (!silencioso) alert(`Roteirização salva: ${filename}`);
    await popularDropdownRoteirizacoes();
    popularSeletorResumoDia().catch(()=>{});
  } catch(e) {
    if (!silencioso) alert('Erro ao salvar: ' + e.message);
    else throw e; // propaga erro para o chamador tratar
  }
}
async function carregarListaHistorico() {
  const el = document.getElementById('historico-list');
  if (!dirHandleHistorico) {
    el.innerHTML = '<div class="empty">Selecione uma pasta para visualizar o histórico.</div>';
    return;
  }
  if (!await _histGarantirPermissao()) {
    el.innerHTML = '<div class="empty">Permissão negada. Clique em "Selecionar pasta" novamente.</div>';
    return;
  }
  document.getElementById('hist-pasta-path').textContent = dirHandleHistorico.name;
  el.innerHTML = '<div class="empty">Carregando...</div>';

  // ── Coleta handles de todos os arquivos JSON em paralelo ───────────────────
  const fileHandles = [];
  for await (const [name, handle] of dirHandleHistorico.entries()) {
    if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
    fileHandles.push({ name, handle });
  }

  // ── Leitura paralela com cache de metadados ────────────────────────────────
  // Para cada arquivo: se já está em cache E o lastModified não mudou, usa o cache.
  // Caso contrário, lê apenas o início do arquivo (os primeiros 4KB contêm savedAt,
  // versao, resumo e datasEntrega) evitando deserializar o JSON completo (que pode
  // ter centenas de KB de pedidos/veículos/resultado).
  const BATCH = 8; // processa N arquivos por vez para não saturar I/O
  const entries = [];

  for (let i = 0; i < fileHandles.length; i += BATCH) {
    const batch = fileHandles.slice(i, i + BATCH);
    const results = await Promise.all(batch.map(async ({ name, handle }) => {
      try {
        const file = await handle.getFile();
        const cached = _histMetaCache[name];
        // Cache hit: mesmo tamanho e lastModified → não relê
        if (cached && cached._lastModified === file.lastModified && cached._size === file.size) {
          return { name, data: cached };
        }
        // Cache miss: lê arquivo completo (necessário para garantir integridade do JSON)
        // mas extrai apenas os metadados e descarta o resto da memória
        const text = await file.text();
        const data = JSON.parse(text);
        if (!data.versao || !data.savedAt || !data.resumo) return null;
        // Salva apenas metadados no cache (não o resultado/pedidos/veículos)
        const meta = {
          versao: data.versao,
          savedAt: data.savedAt,
          resumo: data.resumo,
          datasEntrega: data.datasEntrega || [],
          salvoPor: data.salvoPor || '',
          _lastModified: file.lastModified,
          _size: file.size,
        };
        _histMetaCache[name] = meta;
        return { name, data: meta };
      } catch(e) { return null; }
    }));
    results.forEach(r => { if (r) entries.push(r); });
    // Atualiza UI a cada batch para dar feedback de progresso
    if (fileHandles.length > BATCH) {
      el.innerHTML = `<div class="empty">Carregando… ${Math.min(i + BATCH, fileHandles.length)} / ${fileHandles.length}</div>`;
      await new Promise(r => setTimeout(r, 0)); // yield para o browser renderizar
    }
  }

  entries.sort((a, b) => b.data.savedAt.localeCompare(a.data.savedAt));
  // Reconstrói o contador de IDs a partir do histórico real (evita sequências infladas)
  _reconstruirPetSeqDoHistorico();
  if (!entries.length) {
    el.innerHTML = '<div class="empty">Nenhuma roteirização salva nesta pasta.</div>';
    return;
  }
  el.innerHTML = entries.map(({ name, data }) => {
    const dt    = new Date(data.savedAt);
    const p2    = n => String(n).padStart(2, '0');
    const dtStr = `${p2(dt.getDate())}/${p2(dt.getMonth()+1)}/${dt.getFullYear()} às ${p2(dt.getHours())}:${p2(dt.getMinutes())}`;
    const { totalRotas = 0, totalViagens = 0, totalPedidos = 0, totalVolume_m3 = 0, terminaisUsados = [] } = data.resumo;
    const entrega     = data.datasEntrega?.length ? data.datasEntrega.join(', ') : '—';
    const terminaisStr = terminaisUsados.length ? terminaisUsados.join(' · ') : '';
    const safeNome    = name.replace(/'/g, "\\'");
    return `
      <div class="hist-entry">
        <div class="hist-entry-info">
          <div class="hist-entry-date">${dtStr}</div>
          <div class="hist-entry-meta">Entrega: <strong>${entrega}</strong>${data.salvoPor ? ` &nbsp;·&nbsp; Salvo por: <strong>${data.salvoPor === 'Desconhecido' ? ((S && S.user && window.USERS_DB && window.USERS_DB[S.user]) ? (window.USERS_DB[S.user].name || S.user) : data.salvoPor) : data.salvoPor}</strong>` : ''}</div>
          <div class="hist-entry-chips">
            <span class="tag tag-green">${totalRotas} rota${totalRotas !== 1 ? 's' : ''}</span>
            <span class="tag tag-blue">${totalViagens} ${totalViagens !== 1 ? 'viagens' : 'viagem'}</span>
            <span class="tag tag-gray">${totalPedidos} pedido${totalPedidos !== 1 ? 's' : ''}</span>
            <span class="tag tag-yellow">${String(totalVolume_m3).replace('.', ',')} m³</span>
            ${terminaisStr ? `<span class="pill">${terminaisStr}</span>` : ''}
          </div>
        </div>
        <div class="hist-entry-actions">
          <button class="btn btn-green btn-sm" onclick="abrirEntradaHistorico('${safeNome}')">Abrir</button>
          <button class="btn btn-danger btn-sm" onclick="excluirEntradaHistorico('${safeNome}', this)">Excluir</button>
        </div>
      </div>`;
  }).join('');
}
async function abrirEntradaHistorico(filename) {
  if (!await _histGarantirPermissao()) return;
  try {
    const fh   = await dirHandleHistorico.getFileHandle(filename);
    const file = await fh.getFile();
    const data = JSON.parse(await file.text());
    pedidos             = data.pedidos      || [];
    terminaisCad        = data.terminais    || [];
    veiculos            = data.veiculos     || [];
    ultimoResultado     = data.resultado    || null;
    ultimoControleTempo = data.controleTempo || {};
    resultadoOriginal   = ultimoResultado ? JSON.parse(JSON.stringify(ultimoResultado)) : null;
    historicoManual     = [];
    renderTerminais();
    renderClientes();
    renderPedidos();
    renderVeiculos();
    if (ultimoResultado) {
      renderResultado(ultimoResultado, ultimoControleTempo);
      renderMapaGeral();
    }
    showTab('resultado');
  } catch(e) {
    alert('Erro ao abrir: ' + e.message);
  }
}
async function popularDropdownRoteirizacoes() {
  const selects = document.querySelectorAll('.sel-roteirizacao');
  if (!selects.length) return;
  const opAtual = '<option value="">— Atual —</option>';
  if (!dirHandleHistorico) {
    selects.forEach(s => { const v = s.value; s.innerHTML = opAtual; s.value = v; });
    return;
  }
  let permOk = false;
  try { permOk = (await dirHandleHistorico.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  if (!permOk) {
    selects.forEach(s => { const v = s.value; s.innerHTML = opAtual; s.value = v; });
    return;
  }
  const entries = [];
  for await (const [name, handle] of dirHandleHistorico.entries()) {
    if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
    try {
      const file = await handle.getFile();
      const data = JSON.parse(await file.text());
      if (data.versao && data.savedAt && data.resumo) entries.push({ name, data });
    } catch(e) {}
  }
  entries.sort((a, b) => b.data.savedAt.localeCompare(a.data.savedAt));
  const histOpts = entries.map(({ name, data }) => {
    const dt    = new Date(data.savedAt);
    const p2    = n => String(n).padStart(2, '0');
    const dtStr = `${p2(dt.getDate())}/${p2(dt.getMonth()+1)}/${dt.getFullYear()} ${p2(dt.getHours())}:${p2(dt.getMinutes())}`;
    const entrega = data.datasEntrega?.length ? ` — Entrega: ${data.datasEntrega.join(', ')}` : '';
    const safeNome = name.replace(/"/g, '&quot;');
    return `<option value="${safeNome}">${dtStr}${entrega}</option>`;
  }).join('');
  selects.forEach(s => {
    const current = s.value;
    s.innerHTML = opAtual + histOpts;
    if (current) s.value = current;
  });
}
function _restaurarEstado(estado) {
  pedidos             = estado.pedidos       || [];
  terminaisCad        = estado.terminais     || [];
  veiculos            = estado.veiculos      || [];
  ultimoResultado     = estado.resultado     || null;
  ultimoControleTempo = estado.controleTempo || {};
  resultadoOriginal   = ultimoResultado ? JSON.parse(JSON.stringify(ultimoResultado)) : null;
  historicoManual     = [];
  _atualizarPetSeq(ultimoResultado); // atualiza contador para nunca reusar IDs já no histórico
}
async function onMudarRoteirizacao(val) {
  document.querySelectorAll('.sel-roteirizacao').forEach(s => s.value = val);
  if (!val) {
    if (estadoAtual) _restaurarEstado(estadoAtual);
  } else {
    if (!await _histGarantirPermissao()) return;
    try {
      const fh   = await dirHandleHistorico.getFileHandle(val);
      const file = await fh.getFile();
      _restaurarEstado(JSON.parse(await file.text()));
    } catch(e) {
      alert('Erro ao carregar roteirização: ' + e.message);
      return;
    }
  }
  if (ultimoResultado) {
    lockedTerminals.clear();
    filtroMapaPlaca = '';
    filtroMapaTerminais.clear();
    renderResultado(ultimoResultado, ultimoControleTempo);
    renderMapaGeral();
    renderTemplateOperacao();
  }
}
async function excluirEntradaHistorico(filename, btn) {
  // Invalida cache do arquivo excluído
  delete _histMetaCache[filename];
  if (!confirm(`Excluir "${filename}" do histórico?`)) return;
  if (!await _histGarantirPermissao()) return;
  try {
    await dirHandleHistorico.removeEntry(filename);
    btn.closest('.hist-entry').remove();
    const el = document.getElementById('historico-list');
    if (!el.querySelector('.hist-entry')) {
      el.innerHTML = '<div class="empty">Nenhuma roteirização salva nesta pasta.</div>';
    }
  } catch(e) {
    alert('Erro ao excluir: ' + e.message);
  }
}
// ─── Init ─────────────────────────────────────────────────────────────────────
// Auto-detecção da URL do SharePoint quando o arquivo é aberto de dentro do SP
(function spAutoDetect() {
  const m = window.location.href.match(/^(https:\/\/[^/]+(?:\/sites\/[^/?#]+)?)/);
  if (m && m[1].toLowerCase().includes('sharepoint')) {
    const spUrlEl = document.getElementById('sp-url');
    if (spUrlEl) {
      spUrlEl.value = m[1];
      spSetStatus('ok', 'Site detectado automaticamente');
    }
  }
})();
if (document.getElementById('horario-grid')) gerarHorarioGrid([0,1,2,3,4,5], null);
if (document.getElementById('cl-tipos-btns')) renderTipoBtns('cl-tipos-btns', []);
if (document.getElementById('p-tipos-btns')) renderTipoBtns('p-tipos-btns', []);
// Exposta em window para que o script do Dashboard (carregado depois) possa
// aguardar a restauração do handle salvo no IndexedDB antes de tentar ler o
// histórico em disco — sem isso, dashPopularMeses() no boot do Dashboard corria
// em paralelo com esta função e quase sempre "ganhava" a corrida, lendo
// window.dirHandleHistorico ainda como null mesmo quando uma pasta já tinha
// sido selecionada em uma sessão anterior (o handle só não tinha chegado a
// tempo do IndexedDB).
window.recuperarHandleHistoricoPromise = recuperarHandleHistorico();
window.initRoteirizadorIntegrado = async function(){
  const root = document.getElementById('roteirizador-shell');
  if(!root) return;
  if(!window.__roteirizadorDadosCarregados){
    window.__roteirizadorDadosCarregados = true;
    try { await carregarDadosFixos(); } catch(e) { console.warn('Roteirizador: falha ao carregar dados fixos', e); }
    try { await carregarPedidosLiberados(); } catch(e) { console.warn('Roteirizador: falha ao carregar pedidos', e); }
  }
  try { popularDropdownRoteirizacoes(); popularSeletorResumoDia().catch(()=>{}); } catch(e) {}
  initDataOperacao();
  if(typeof invalidateSizeMapasRoteirizador === 'function') invalidateSizeMapasRoteirizador();
};
window.invalidateSizeMapasRoteirizador = function(){
  [120, 350, 700, 1200].forEach(delay => {
    setTimeout(()=>{
      try { if(mapaGeral  && typeof mapaGeral.invalidateSize  === 'function') mapaGeral.invalidateSize();  } catch(e) {}
      try { if(mapaViagem && typeof mapaViagem.invalidateSize === 'function') mapaViagem.invalidateSize(); } catch(e) {}
    }, delay);
  });
};
/* ══════════════════════════════════════════════════════════════
   RELATÓRIO EXCEL — Roteirização Completa
   Exporta 3 abas:
   1. Resumo por Veículo
   2. Detalhamento por Entrega (uma linha por produto por parada)
   3. Resumo por Cliente
══════════════════════════════════════════════════════════════ */
// Helper local — equivalente ao dateStr() do sistema pai
function _isoHoje(offsetDias) {
  const d = new Date(); d.setDate(d.getDate() + (offsetDias||0));
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function onDataOperacaoChange(valor) {
  if (!valor) return;
  sincronizarDisponibilidadeVeiculos(valor).then(() => {
    if (document.getElementById('tab-veiculos')?.classList.contains('active')) renderVeiculos();
    showToast(`Disponibilidade atualizada para ${new Date(valor + 'T12:00:00').toLocaleDateString('pt-BR')} ✅`, true);
  }).catch(() => {});
}
function initDataOperacao() {
  const el = document.getElementById('rot-data-operacao');
  if (!el || el._inicializado) return;
  el._inicializado = true;
  const amanha = new Date(); amanha.setDate(amanha.getDate() + 1);
  const iso = `${amanha.getFullYear()}-${String(amanha.getMonth()+1).padStart(2,'0')}-${String(amanha.getDate()).padStart(2,'0')}`;
  el.value = iso;
  sincronizarDisponibilidadeVeiculos(iso).catch(() => {});
}
window.exportarRelatorioRoteirizacao = function exportarRelatorioRoteirizacao() {
  try {
    if (!ultimoResultado) { showToast('Roteirize primeiro para gerar o relatório.', false); return; }
    const todosVeic = veiculos.filter(v => ultimoResultado[v.id] && ultimoResultado[v.id].some(vi => !vi._vazio && (vi.paradas||[]).length));
    if (!todosVeic.length) { showToast('Nenhuma rota gerada para exportar.', false); return; }
    const dataEntrega = document.getElementById('rot-data-operacao')?.value || document.getElementById('rot-data-entrega')?.value || _isoHoje(0);
    const nomeArq = 'Roteirizacao_' + dataEntrega.replace(/-/g,'') + '_' + new Date().toISOString().slice(11,16).replace(':','');
    function hav(a,b,c,d){const R=6371,dL=(c-a)*Math.PI/180,dl=(d-b)*Math.PI/180,x=Math.sin(dL/2)**2+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(dl/2)**2;return R*2*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));}
    function mH(m){if(m==null||isNaN(m))return '-';const hm=((Math.round(m)%1440)+1440)%1440;return String(Math.floor(hm/60)).padStart(2,'0')+':'+String(hm%60).padStart(2,'0');}
    function mkWs(data){const ws=XLSX.utils.aoa_to_sheet(data);const r=XLSX.utils.decode_range(ws['!ref']||'A1');ws['!cols']=Array.from({length:r.e.c+1},()=>({wch:22}));return ws;}
    // ── Aba 1: Resumo por Veículo ─────────────────────────────────────────────
    const a1=[['RESUMO POR VEÍCULO — '+dataEntrega],[],
      ['PLACA','TRANSPORTADORA','CONTRATO','TIPO','TERMINAL','MOTORISTA DIURNO','MOTORISTA NOTURNO','Nº VIAGENS','Nº ENTREGAS','VOLUME TOTAL (m³)','CAPACIDADE (m³)','OCUPAÇÃO (%)','JORNADA (h)']];
    todosVeic.forEach(v=>{
      const vis=(ultimoResultado[v.id]||[]).filter(vi=>!vi._vazio&&(vi.paradas||[]).length);
      const tv=vis.reduce((s,vi)=>s+(vi.paradas||[]).reduce((sv,p)=>sv+(p.volumeTotal||0),0),0);
      const td=vis.reduce((s,vi)=>s+(vi.paradas||[]).length,0);
      const cap=v.capacidade||v.capacidadeTotal||0;
      a1.push([v.placa,v.transportadora||'-',v.contrato||'-',v.tipo||'-',v.terminal||'-',v.motoristaDiurno||'-',v.motoristaNt||'-',vis.length,td,parseFloat(tv.toFixed(3)),cap,cap>0?((tv/(cap*vis.length))*100).toFixed(1):'-',parseFloat((vis.reduce((s,vi)=>s+(vi.jornadaProdMin||0),0)/60).toFixed(2))]);
    });
    // ── Aba 2: Resumo por Viagem (UMA LINHA POR VIAGEM) ──────────────────────
    const a2=[['RESUMO POR VIAGEM — '+dataEntrega],[],
      ['ID VIAGEM','PLACA','TRANSPORTADORA','CONTRATO','TIPO','TERMINAL','MOTORISTA','Nº ENTREGAS','VOLUME (m³)','OCUPAÇÃO (%)','HORA SAÍDA','HORA CHEGADA ÚLTIMO','JORNADA (h)']];
    todosVeic.forEach(v=>{
      const vis=(ultimoResultado[v.id]||[]).filter(vi=>!vi._vazio&&(vi.paradas||[]).length);
      const jI=(()=>{const h=parseHoraMin(v.jornadaInicio||'06:00');return isNaN(h)?360:h;})();
      const cap=v.capacidade||v.capacidadeTotal||0;
      vis.forEach((vi,iV)=>{
        let ck=inicioViagemAbsMin(vis,iV,jI,v.tempoPerdidoMin||0,doisTurnos(v)?2:1)+(vi.esperaTerminalMin||0);
        const hSaida=mH(ck+(vi.paradas[0]?.tempoCarregamentoMin||0));
        // Hora chegada no último cliente: percorrer todas as paradas sequencialmente
        let ckFim=ck+(vi.paradas[0]?.tempoCarregamentoMin||0);
        vi.paradas.forEach((par,iP)=>{
          ckFim+=Math.round(par.deslocCarregadoMin||0)+(par.tempoEsperaRestricaoMin||0);
          if(iP<vi.paradas.length-1) ckFim+=Math.round(par.tempoDescargaMin||0)+Math.round(par.deslocVazioMin||0);
        });
        const hChegada=mH(ckFim);
        const volV=vi.paradas.reduce((s,p)=>s+(p.volumeTotal||0),0);
        const ocup=cap>0?((volV/cap)*100).toFixed(1):'-';
        const jornH=parseFloat(((vi.jornadaProdMin||0)/60).toFixed(2));
        const petId=vi.petId||`VIA${v.placa}${iV+1}`;
        a2.push([petId,v.placa,v.transportadora||'-',v.contrato||'-',v.tipo||'-',v.terminal||'-',v.motoristaDiurno||v.motoristaNt||'-',vi.paradas.length,parseFloat(volV.toFixed(3)),ocup,hSaida,hChegada,jornH]);
      });
    });
    // ── Aba 3: Detalhe por Entrega ────────────────────────────────────────────
    const a3=[['DETALHAMENTO POR ENTREGA — '+dataEntrega],[],
      ['ID VIAGEM','PLACA','TRANSPORTADORA','CONTRATO','TIPO','TERMINAL','MOTORISTA','SEQ.','CLIENTE','CIDADE','UF','CÓD. PRODUTO','PRODUTO','VOLUME (m³)','VOLUME (L)','HORA SAÍDA','HORA CHEGADA','JANELA','DIST. TERMINAL (km)','DESCARGA (min)','ORDEM SAP']];
    todosVeic.forEach(v=>{
      const vis=(ultimoResultado[v.id]||[]).filter(vi=>!vi._vazio&&(vi.paradas||[]).length);
      const jI=(()=>{const h=parseHoraMin(v.jornadaInicio||'06:00');return isNaN(h)?360:h;})();
      vis.forEach((vi,iV)=>{
        const petId=vi.petId||`VIA${v.placa}${iV+1}`;
        let ck=inicioViagemAbsMin(vis,iV,jI,v.tempoPerdidoMin||0,doisTurnos(v)?2:1)+(vi.esperaTerminalMin||0);
        const hs=mH(ck+(vi.paradas[0]?.tempoCarregamentoMin||0));
        // Busca coordenadas do terminal via terminaisCad
        const termNome = vi.terminalOrigem || v.terminal || vi.paradas[0]?.pedido?.terminal || '';
        const termObj  = (typeof terminaisCad !== 'undefined' ? terminaisCad : []).find(t => t.nome === termNome);
        const tLa = termObj ? parseFloat(termObj.lat) : null;
        const tLo = termObj ? parseFloat(termObj.lon) : null;
        vi.paradas.forEach((par,iP)=>{
          if(iP===0) ck+=Math.round(par.tempoCarregamentoMin||0);
          ck+=Math.round(par.deslocCarregadoMin||0)+(par.tempoEsperaRestricaoMin||0);
          const hc=mH(ck);
          const dk=(tLa&&par.lat&&par.lon)?hav(tLa,tLo,par.lat,par.lon).toFixed(1):'-';
          const ped=par.pedido||{};
          const prods=(ped.produtos||[]).length?ped.produtos:[{produto:'-',volume:par.volumeTotal||0,ordemSAP:ped.ordemSAP||'-'}];
          prods.forEach(pr=>{
            const[cod,...desc]=(pr.produto||'').split(' - ');
            a3.push([petId,v.placa,v.transportadora||'-',v.contrato||'-',v.tipo||'-',v.terminal||'-',v.motoristaDiurno||v.motoristaNt||'-',iP+1,ped.cliente||ped.nomeCliente||par.nome||'-',ped.cidade||'-',ped.uf||'-',cod?.trim()||'-',desc.join(' - ')||pr.produto||'-',pr.volume||0,parseFloat(((pr.volume||0)*1000).toFixed(0)),hs,hc,ped.restricao||'-',dk,Math.round(par.tempoDescargaMin)||'-',pr.ordemSAP||ped.ordemSAP||'-']);
          });
          ck+=Math.round(par.tempoDescargaMin||0)+Math.round(par.deslocVazioMin||0);
        });
      });
    });
    // ── Aba 4: Resumo por Cliente ─────────────────────────────────────────────
    const pC=new Map();
    todosVeic.forEach(v=>{(ultimoResultado[v.id]||[]).filter(vi=>!vi._vazio&&(vi.paradas||[]).length).forEach(vi=>{vi.paradas.forEach(par=>{const ped=par.pedido||{};const key=ped.cliente||ped.nomeCliente||par.nome||'?';if(!pC.has(key))pC.set(key,{c:key,ci:ped.cidade||'-',uf:ped.uf||'-',v:0,e:0,pl:new Set(),pr:new Map()});const x=pC.get(key);x.v+=(par.volumeTotal||0);x.e++;x.pl.add(v.placa);(ped.produtos||[]).forEach(pr=>{const nm=(pr.produto||'').replace(/^\d+\s*-\s*/,'');x.pr.set(nm,(x.pr.get(nm)||0)+(pr.volume||0));});});});});
    const a4=[['RESUMO POR CLIENTE — '+dataEntrega],[],['CLIENTE','CIDADE','UF','ENTREGAS','VOLUME TOTAL (m³)','PRODUTOS','PLACAS']];
    [...pC.values()].sort((a,b)=>a.c.localeCompare(b.c)).forEach(x=>{a4.push([x.c,x.ci,x.uf,x.e,parseFloat(x.v.toFixed(3)),[...x.pr.entries()].map(([n,v])=>n+': '+v.toFixed(1)+'m³').join(' | ')||'-',[...x.pl].join(', ')]);});
    // ── Workbook ──────────────────────────────────────────────────────────────
    const wb=XLSX.utils.book_new();
    [{d:a1,n:'1. Resumo Veículos'},{d:a2,n:'2. Resumo por Viagem'},{d:a3,n:'3. Detalhe Entregas'},{d:a4,n:'4. Resumo Clientes'}].forEach(({d,n})=>{XLSX.utils.book_append_sheet(wb,mkWs(d),n);});
    XLSX.writeFile(wb, nomeArq+'.xlsx');
    const _st = typeof showToast === 'function' ? showToast : (m) => alert(m);
    _st('Relatório Excel exportado! ✅', true);
  } catch(e) {
    console.error('Erro ao exportar relatório:', e);
    const _st = typeof showToast === 'function' ? showToast : (m) => alert(m);
    _st('Erro ao gerar relatório: '+e.message, false);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// RESUMO DO DIA — agrega TODAS as roteirizações do dia selecionado
// e gera texto formatado para WhatsApp
// ══════════════════════════════════════════════════════════════════════════════

// Popula o seletor de datas no Envio Transportadora com os dias únicos do histórico
async function popularSeletorResumoDia() {
  const sel = document.getElementById('resumo-dia-sel');
  if (!sel) return;
  sel.innerHTML = '<option value="">— Selecionar data —</option>';
  if (!dirHandleHistorico) return;
  let permOk = false;
  try { permOk = (await dirHandleHistorico.queryPermission({ mode: 'readwrite' })) === 'granted'; } catch(e) {}
  if (!permOk) return;

  // Coleta todas as entradas e agrupa por dia (YYYYMMDD = primeiros 8 chars do nome)
  const diasMap = new Map(); // YYYYMMDD → { label, datasEntrega[] }
  for await (const [name, handle] of dirHandleHistorico.entries()) {
    if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
    const diaKey = name.slice(0, 8); // ex: "20260624"
    if (!/^\d{8}$/.test(diaKey)) continue;
    try {
      const file = await handle.getFile();
      const data = JSON.parse(await file.text());
      if (!data.versao || !data.savedAt || !data.resumo) continue;
      if (!diasMap.has(diaKey)) {
        const y = diaKey.slice(0,4), m = diaKey.slice(4,6), d = diaKey.slice(6,8);
        diasMap.set(diaKey, { label: `${d}/${m}/${y}`, datasEntrega: new Set() });
      }
      (data.datasEntrega || []).forEach(de => diasMap.get(diaKey).datasEntrega.add(de));
    } catch(e) {}
  }

  // Ordena decrescente (mais recente primeiro)
  const dias = [...diasMap.entries()].sort((a, b) => b[0].localeCompare(a[0]));
  dias.forEach(([key, info]) => {
    const opt = document.createElement('option');
    opt.value = key;
    const entrega = info.datasEntrega.size ? ` — Entrega: ${[...info.datasEntrega].join(', ')}` : '';
    opt.textContent = `${info.label}${entrega}`;
    sel.appendChild(opt);
  });
}

async function gerarResumoDia() {
  const sel    = document.getElementById('resumo-dia-sel');
  const diaKey = sel?.value; // "YYYYMMDD" ou ""

  let snaps = [];

  if (diaKey) {
    // Carrega TODOS os snapshots do dia selecionado
    if (!await _histGarantirPermissao()) {
      showToast('Permissão negada. Selecione a pasta do histórico novamente.', false);
      return;
    }
    for await (const [name, handle] of dirHandleHistorico.entries()) {
      if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
      if (!name.startsWith(diaKey)) continue;
      try {
        const file = await handle.getFile();
        const data = JSON.parse(await file.text());
        if (data.versao && data.savedAt && data.resumo) snaps.push(data);
      } catch(e) {}
    }
    if (!snaps.length) {
      showToast('Nenhuma roteirização encontrada para este dia.', false);
      return;
    }
    // Ordena por hora (mais recente último — a última roteirização do dia tem prioridade
    // para a lista de veículos disponíveis)
    snaps.sort((a, b) => a.savedAt.localeCompare(b.savedAt));
  } else if (ultimoResultado) {
    // Usa a roteirização atual em memória
    snaps = [{
      savedAt: new Date().toISOString(),
      datasEntrega: [],
      resumo: {},
      veiculos,
      resultado: ultimoResultado,
      controleTempo: ultimoControleTempo || {},
      pedidos,
    }];
  } else {
    showToast('Selecione uma data ou execute a otimização primeiro.', false);
    return;
  }

  const texto = _montarTextoResumoDia(snaps);
  try {
    await navigator.clipboard.writeText(texto);
  } catch(e) {
    const ta = document.createElement('textarea');
    ta.value = texto;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
  }
  showToast('✅ Resumo do dia copiado! Cole no WhatsApp.', true);
}

// Agrega N snapshots do mesmo dia e monta o texto
function _montarTextoResumoDia(snaps) {
  const p2 = n => String(n).padStart(2, '0');

  // Data de entrega: une todas as datas de entrega únicas de todos os snaps
  const datasEntregaSet = new Set();
  snaps.forEach(s => (s.datasEntrega || []).forEach(d => datasEntregaSet.add(d)));
  const dataEntregaStr = datasEntregaSet.size
    ? [...datasEntregaSet].sort().join(', ')
    : (() => {
        const d = new Date(snaps[snaps.length - 1].savedAt);
        return `${p2(d.getDate())}/${p2(d.getMonth()+1)}/${d.getFullYear()}`;
      })();

  // ── Agrega métricas de todos os snapshots ────────────────────────────────
  // Para placas: usa o snapshot mais recente como fonte de verdade
  // Para volumes/viagens/pedidos/jornada: soma tudo
  let totalVolume      = 0;
  let totalViagens     = 0;
  let totalPedidos     = 0;
  let horasConsumidas  = 0;
  let horasDisponiveis = 0;
  let temEstouroJornada   = false;
  const pedidosComProblema   = new Set();
  const placasComProgramacao = new Set();
  const placasDisponiveisMap = new Map(); // placa → { transportadora }

  snaps.forEach(snap => {
    const res   = snap.resultado || {};
    const veics = snap.veiculos  || [];
    const ct    = snap.controleTempo || {};

    veics.forEach(v => {
      // Registra placa disponível (o snap mais recente sobrescreve transportadora)
      // Só registra veículos marcados como Disponível no painel do dia
      if (v.disponibilidade === 'Disponível') {
        placasDisponiveisMap.set(v.placa, {
          transportadora: v.transportadora || '',
          capacidade: v.capacidade || 0,
          cidadeBase: v.cidadeBase || v.cidade || '',
        });
      }

      const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas || []).length);
      const ctV = ct[String(v.id)] || {};

      viagens.forEach(vi => {
        totalViagens++;
        vi.paradas.forEach(p => {
          totalVolume += p.volumeTotal || 0;
          totalPedidos++;
          if (p.tempoEsperaRestricaoMin > 0) {
            pedidosComProblema.add(p.pedido?.cliente || p.pedido?.nomeCliente || '?');
          }
        });
        placasComProgramacao.add(v.placa);
      });

      const limMin   = ctV.limiteMin  || 0;
      const usadoMin = ctV.usadoMin   || 0;
      horasDisponiveis += limMin;
      horasConsumidas  += usadoMin;
      if (usadoMin > limMin + 1) temEstouroJornada = true;
    });
  });

  const placasSemProg = [...placasDisponiveisMap.entries()]
    .filter(([placa]) => !placasComProgramacao.has(placa))
    .sort(([a], [b]) => a.localeCompare(b));

  // Ocupação da frota
  let totalCapacidade = 0;
  let totalVolumeAlocado = 0;
  let veiculosOcupacaoTotal = 0;
  let veiculosOcupacaoParcial = 0;
  snaps.forEach(snap => {
    const res  = snap.resultado || {};
    const veics = snap.veiculos || [];
    veics.forEach(v => {
      const viagens = (res[v.id] || []).filter(vi => !vi._vazio && (vi.paradas || []).length);
      if (!viagens.length) return;
      const volVeiculo = viagens.reduce((s, vi) => s + vi.paradas.reduce((ss, p) => ss + (p.volumeTotal || 0), 0), 0);
      const cap = v.capacidade || 0;
      if (cap > 0) {
        totalCapacidade    += cap;
        totalVolumeAlocado += volVeiculo;
        const pct = volVeiculo / cap;
        if (pct >= 0.99) veiculosOcupacaoTotal++;
        else veiculosOcupacaoParcial++;
      }
    });
  });
  const pctFrota = totalCapacidade > 0 ? Math.round((totalVolumeAlocado / totalCapacidade) * 100) : 0;

  const horasConsH = (horasConsumidas  / 60).toFixed(1).replace('.', ',');
  const horasDispH = (horasDisponiveis / 60).toFixed(1).replace('.', ',');
  const volumeStr  = totalVolume.toFixed(1).replace('.', ',');
  const semProblemas = pedidosComProblema.size === 0 && !temEstouroJornada;

  // ── Monta texto ───────────────────────────────────────────────────────────
  const L = [];
  L.push(`*Resumo do Outbound - ${dataEntregaStr}*`);
  L.push('');

  if (semProblemas) {
    L.push(
      `A programação para o dia ${dataEntregaStr} foi concluída com bom nível de aproveitamento ` +
      `da frota e atendimento integral da demanda prevista. A operação foi estruturada sem quebras ` +
      `de pedido e as rotas foram importadas em lote para a Herrlog.`
    );
  } else {
    const prob = [];
    if (pedidosComProblema.size > 0) prob.push(`${pedidosComProblema.size} pedido(s) com restrição de horário`);
    if (temEstouroJornada) prob.push('possível estouro de jornada em alguma programação');
    L.push(
      `A programação para o dia ${dataEntregaStr} foi concluída. ` +
      `Foram identificados: ${prob.join(' e ')}.`
    );
  }
  L.push('');

  L.push(`• *Volume planejado:* ${volumeStr} m³`);
  L.push(`• *Pedidos atendidos:* ${totalPedidos}`);
  L.push(`• *Viagens geradas:* ${totalViagens} — importadas em lote para a Herrlog`);
  L.push(`• *Possível estouro de jornada:* ${temEstouroJornada ? '⚠️ Sim' : '✅ Não'}`);
  L.push(`• *Ocupação da frota:* ${pctFrota}% · ${veiculosOcupacaoTotal} veículo(s) com ocupação total · ${veiculosOcupacaoParcial} parcial`);

  L.push(`• *Jornada consumida:* ${horasConsH} h de ${horasDispH} h disponíveis`);
  L.push('');

  L.push(`*Veículos Disponíveis*`);
  L.push('');
  if (placasSemProg.length > 0) {
    placasSemProg.forEach(([placa, info]) => {
      const transp = info.transportadora ? ` · ${info.transportadora}` : '';
      const cap    = info.capacidade ? ` · ${info.capacidade} m³` : '';
      // Terminal: pega só a cidade/base (ex: "Paulínia TORRÃO Nexta" → "Paulínia")
      const termLabel = info.cidadeBase ? ` · ${info.cidadeBase}` : '';
      L.push(`• ${placa}${cap}${termLabel}${transp}`);
    });
  } else {
    L.push('• Todos os veículos disponíveis foram programados.');
  }

  return L.join('\n');
}

window.gerarResumoDia        = gerarResumoDia;
window.popularSeletorResumoDia = popularSeletorResumoDia;

// ═══════════════════════════════════════════════════════════════════════════
// EXPORTAÇÃO PDF / PNG — Programações do Resumo Transportadora
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Clona um op-bloco para um iframe oculto com estilos inline,
 * garantindo renderização fiel sem depender das classes CSS da página.
 */
function _clonarBlocoParaExport(blocoEl) {
  // Coleta as folhas de estilo relevantes como texto inline
  const cssTexto = Array.from(document.styleSheets)
    .flatMap(ss => {
      try { return Array.from(ss.cssRules).map(r => r.cssText); }
      catch(e) { return []; }
    }).join('\n');

  // Cria iframe temporário fora da viewport
  const iframe = document.createElement('iframe');
  iframe.style.cssText = 'position:fixed;top:-9999px;left:-9999px;width:1050px;height:auto;border:none;visibility:hidden;';
  document.body.appendChild(iframe);

  const doc = iframe.contentDocument;
  doc.open();
  doc.write(`<!DOCTYPE html><html><head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&family=Barlow+Condensed:wght@600;700;800&display=swap" rel="stylesheet">
    <style>
      ${cssTexto}
      /* Overrides para export limpo */
      body { margin: 0; padding: 16px; background: #fff; font-family: 'DM Sans', sans-serif; }
      :root {
        --bg: #F8FAF4; --surface: #fff; --border: #D4E4C0; --border-dk: #BDD6A3;
        --text: #1A2A0A; --text-3: #5A7A42; --radius: 8px; --radius-lg: 12px;
        --shadow: 0 1px 4px rgba(0,0,0,.06);
        --green-bg: #EDF7E0; --green-text: #2D6A1B; --green-border: #9FD07A;
        --amber-bg: #FEF3C7; --amber-text: #92400E; --amber-border: #FCD34D;
        --red-bg: #FEE2E2; --red-text: #991B1B; --red-border: #FCA5A5;
        --pet-green: #4A7C30; --pet-green-bg: rgba(74,124,48,.08);
        --font: 'DM Sans', sans-serif; --font-cond: 'Barlow Condensed', sans-serif; --font-mono: 'DM Mono', monospace;
      }
      #roteirizador-shell { all: unset; }
      .op-bloco { border: 0.5px solid var(--border); border-radius: var(--radius-lg); background: var(--surface); overflow: hidden; box-shadow: var(--shadow); }
      .op-head { padding: 11px 16px; background: var(--bg); border-bottom: 0.5px solid var(--border); }
      .op-head-lbl { color: var(--text-3); font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; font-size: 10px; }
      .op-table { width: 100%; border-collapse: collapse; font-size: 12px; }
      .op-table th, .op-table td { border: 0.5px solid var(--border); padding: 7px 9px; text-align: left; }
      .op-table th { background: var(--bg); font-family: var(--font); letter-spacing: 0.04em; text-transform: uppercase; font-size: 11px; color: var(--text-3); font-weight: 600; }
      .op-ocup { display: inline-block; padding: 2px 8px; border-radius: 20px; font-weight: 500; font-size: 11px; border: 0.5px solid; }
      .op-ocup-verde   { background: var(--green-bg); color: var(--green-text); border-color: var(--green-border); }
      .op-ocup-amarelo { background: var(--amber-bg); color: var(--amber-text); border-color: var(--amber-border); }
      .op-ocup-vermelho{ background: var(--red-bg);   color: var(--red-text);   border-color: var(--red-border); }
      /* Oculta botões de exportação no clone */
      button { display: none !important; }
    </style>
  </head><body><div id="roteirizador-shell">${blocoEl.outerHTML}</div></body></html>`);
  doc.close();
  return iframe;
}

/**
 * Usa html2canvas para capturar um elemento como canvas.
 */
async function _capturarCanvas(el, escala = 2) {
  return html2canvas(el, {
    scale: escala,
    useCORS: true,
    backgroundColor: '#ffffff',
    logging: false,
    allowTaint: false,
  });
}

/**
 * Exporta um único bloco de programação como PDF.
 * @param {string} blocoId  - valor de data-bloco-id do card
 * @param {string} nomeArq  - nome base do arquivo (sem extensão)
 * @param {Event}  ev       - evento click (para stopPropagation)
 */
async function exportarBlocoPDF(blocoId, nomeArq, ev) {
  if (ev) ev.stopPropagation();
  const blocoEl = document.querySelector(`[data-bloco-id="${blocoId}"]`);
  if (!blocoEl) { showToast('Card não encontrado.', false); return; }
  showToast('Gerando PDF…', true);
  try {
    const iframe = _clonarBlocoParaExport(blocoEl);
    // Aguarda fontes carregarem
    await new Promise(r => setTimeout(r, 600));
    const clone = iframe.contentDocument.querySelector('.op-bloco');
    const canvas = await _capturarCanvas(clone);
    document.body.removeChild(iframe);

    const { jsPDF } = window.jspdf;
    // A4 paisagem para cards horizontais do Resumo Transportadora
    const pdfW  = 841.89;
    const pdfHMin = 595.28;
    const margin = 20;
    const maxW  = pdfW - margin * 2;
    const ratio = maxW / canvas.width;
    const imgH  = canvas.height * ratio;
    const pageH = Math.max(imgH + margin * 2, pdfHMin);
    const pdf = new jsPDF({ orientation: 'l', unit: 'pt', format: [pdfW, pageH] });
    pdf.addImage(canvas.toDataURL('image/png'), 'PNG', margin, margin, maxW, imgH);
    pdf.save(`${nomeArq}.pdf`);
    showToast('PDF exportado ✅', true);
  } catch(e) {
    console.error('Erro ao exportar PDF:', e);
    showToast('Erro ao gerar PDF: ' + e.message, false);
  }
}

/**
 * Exporta um único bloco de programação como PNG.
 */
async function exportarBlocoPNG(blocoId, nomeArq, ev) {
  if (ev) ev.stopPropagation();
  const blocoEl = document.querySelector(`[data-bloco-id="${blocoId}"]`);
  if (!blocoEl) { showToast('Card não encontrado.', false); return; }
  showToast('Gerando PNG…', true);
  try {
    const iframe = _clonarBlocoParaExport(blocoEl);
    await new Promise(r => setTimeout(r, 600));
    const clone = iframe.contentDocument.querySelector('.op-bloco');
    const canvas = await _capturarCanvas(clone, 3);
    document.body.removeChild(iframe);

    const a = document.createElement('a');
    a.href = canvas.toDataURL('image/png');
    a.download = `${nomeArq}.png`;
    a.click();
    showToast('PNG exportado ✅', true);
  } catch(e) {
    console.error('Erro ao exportar PNG:', e);
    showToast('Erro ao gerar PNG: ' + e.message, false);
  }
}

/**
 * Exporta TODOS os blocos visíveis como um único PDF multi-página.
 */
async function exportarTodasProgramacoesPDF() {
  const blocos = document.querySelectorAll('#operacao-content [data-bloco-id]');
  if (!blocos.length) { showToast('Nenhuma programação para exportar.', false); return; }
  showToast(`Gerando ${blocos.length} PDF(s)…`, true);
  const agora = new Date();
  const p2 = n => String(n).padStart(2, '0');
  const prefixo = `Prog_${agora.getFullYear()}${p2(agora.getMonth()+1)}${p2(agora.getDate())}_${p2(agora.getHours())}${p2(agora.getMinutes())}`;
  let i = 0;
  for (const blocoEl of blocos) {
    i++;
    try {
      const blocoId = blocoEl.getAttribute('data-bloco-id') || `bloco${i}`;
      const iframe  = _clonarBlocoParaExport(blocoEl);
      await new Promise(r => setTimeout(r, 600));
      const clone  = iframe.contentDocument.querySelector('.op-bloco');
      const canvas = await _capturarCanvas(clone);
      document.body.removeChild(iframe);

      const { jsPDF } = window.jspdf;
      // Layout do card é horizontal (1050px de largura) — usa A4 paisagem para maximizar
      // a área útil e evitar que a imagem fique espremida em retrato.
      const pdfW  = 841.89; // A4 landscape largura (pt)
      const pdfHMin = 595.28; // A4 landscape altura mínima (pt)
      const margin = 20;
      const maxW  = pdfW - margin * 2;
      const ratio = maxW / canvas.width;
      const imgH  = canvas.height * ratio;
      const pageH = Math.max(imgH + margin * 2, pdfHMin);
      const pdf = new jsPDF({ orientation: 'l', unit: 'pt', format: [pdfW, pageH] });
      pdf.addImage(canvas.toDataURL('image/png'), 'PNG', margin, margin, maxW, imgH);
      pdf.save(`${prefixo}_${String(i).padStart(2,'0')}_${blocoId.replace(/[^a-zA-Z0-9_-]/g,'')}.pdf`);
      await new Promise(r => setTimeout(r, 400));
    } catch(e) {
      console.warn(`Erro ao exportar bloco ${i}:`, e);
    }
  }
  showToast(`${i} PDF(s) exportados ✅`, true);
}

/**
 * Exporta TODOS os blocos visíveis, cada um como um PNG separado (zip não disponível,
 * então faz download sequencial com intervalo para não bloquear o browser).
 */
async function exportarTodasProgramacoesPNG() {
  const blocos = document.querySelectorAll('#operacao-content [data-bloco-id]');
  if (!blocos.length) { showToast('Nenhuma programação para exportar.', false); return; }
  showToast(`Gerando ${blocos.length} PNG(s)…`, true);
  const agora = new Date();
  const p2 = n => String(n).padStart(2, '0');
  const prefixo = `Prog_${agora.getFullYear()}${p2(agora.getMonth()+1)}${p2(agora.getDate())}_${p2(agora.getHours())}${p2(agora.getMinutes())}`;
  let i = 0;
  for (const blocoEl of blocos) {
    i++;
    try {
      const blocoId  = blocoEl.getAttribute('data-bloco-id') || `bloco${i}`;
      const iframe   = _clonarBlocoParaExport(blocoEl);
      await new Promise(r => setTimeout(r, 600));
      const clone  = iframe.contentDocument.querySelector('.op-bloco');
      const canvas = await _capturarCanvas(clone, 3);
      document.body.removeChild(iframe);
      const a = document.createElement('a');
      a.href = canvas.toDataURL('image/png');
      a.download = `${prefixo}_${String(i).padStart(2,'0')}_${blocoId.replace(/[^a-zA-Z0-9_-]/g,'')}.png`;
      a.click();
      // Pausa entre downloads para o browser não bloquear
      await new Promise(r => setTimeout(r, 400));
    } catch(e) {
      console.warn(`Erro ao exportar bloco ${i}:`, e);
    }
  }
  showToast(`${i} PNG(s) exportados ✅`, true);
}

window.exportarBlocoPDF              = exportarBlocoPDF;
window.exportarBlocoPNG              = exportarBlocoPNG;
window.exportarTodasProgramacoesPDF  = exportarTodasProgramacoesPDF;
window.exportarTodasProgramacoesPNG  = exportarTodasProgramacoesPNG;
// ══════════════════════════════════════════════════════════════════════════════
// CALCULADORA DE FRETE
// Contratos: fixo+km | fixo+m3 | diaria | spot por rota
// ══════════════════════════════════════════════════════════════════════════════

const FRETE_LS_KEY   = 'nexta_frete_contratos_v1';
const FRETE_SPOT_KEY = 'nexta_frete_spot_v1';
const FRETE_TIPOS    = { fixo_km:'Fixo + R$/km', fixo_m3:'Fixo + R$/m³', diaria:'Diária', spot:'Spot' };

function freteCarregarContratos() {
  try { return JSON.parse(localStorage.getItem(FRETE_LS_KEY) || '[]'); } catch(e) { return []; }
}
function freteSalvarContratos(arr) {
  localStorage.setItem(FRETE_LS_KEY, JSON.stringify(arr));
}
function freteCarregarSpot() {
  try { return JSON.parse(localStorage.getItem(FRETE_SPOT_KEY) || '[]'); } catch(e) { return []; }
}
function freteSalvarSpot(arr) {
  localStorage.setItem(FRETE_SPOT_KEY, JSON.stringify(arr));
}

function _freteInputStyle(extra) {
  return 'font-size:11px;border:1px solid var(--border-dk);border-radius:5px;padding:3px 5px;background:var(--surface);color:var(--text);box-sizing:border-box;' + (extra||'');
}

function _freteListaTransportadoras() {
  var base = (window.CARRIERS || []).slice();
  (veiculos || []).forEach(function(v) { if (v.transportadora && base.indexOf(v.transportadora) === -1) base.push(v.transportadora); });
  return base.filter(Boolean).sort(function(a,b) { return a.localeCompare(b, 'pt-BR'); });
}

function _freteSelectTransportadora(valorAtual, onChangeFn) {
  var lista = _freteListaTransportadoras();
  var sel = document.createElement('select');
  sel.style.cssText = _freteInputStyle('width:100%;');
  var optVazia = document.createElement('option');
  optVazia.value = '';
  optVazia.textContent = lista.length ? 'Selecione...' : 'Nenhuma cadastrada';
  sel.appendChild(optVazia);
  lista.forEach(function(nome) {
    var opt = document.createElement('option');
    opt.value = nome; opt.textContent = nome;
    if (nome === valorAtual) opt.selected = true;
    sel.appendChild(opt);
  });
  if (valorAtual && lista.indexOf(valorAtual) === -1) {
    var optExtra = document.createElement('option');
    optExtra.value = valorAtual;
    optExtra.textContent = valorAtual + ' (não cadastrada)';
    optExtra.selected = true;
    sel.appendChild(optExtra);
  }
  sel.onchange = onChangeFn;
  return sel;
}

function freteRenderContratos() {
  const tbody = document.getElementById('frete-contratos-body');
  if (!tbody) return;
  const contratos = freteCarregarContratos();
  if (!contratos.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="padding:20px;text-align:center;color:var(--text-3);font-size:12px;">Nenhum contrato cadastrado. Clique em "+ Adicionar" para começar.</td></tr>';
    return;
  }
  tbody.innerHTML = '';
  contratos.forEach(function(c, i) {
    var tr = document.createElement('tr');
    tr.style.borderTop = '1px solid var(--border-dk)';

    var tdPlaca = document.createElement('td');
    tdPlaca.style.cssText = 'padding:7px 8px;font-weight:600;color:var(--text);';
    var inpPlaca = document.createElement('input');
    inpPlaca.value = c.placa || ''; inpPlaca.placeholder = 'Placa';
    inpPlaca.style.cssText = _freteInputStyle('width:100%;font-weight:600;');
    inpPlaca.onchange = function() { freteEditarContrato(i, 'placa', this.value); };
    tdPlaca.appendChild(inpPlaca); tr.appendChild(tdPlaca);

    var tdTransp = document.createElement('td');
    tdTransp.style.cssText = 'padding:7px 8px;';
    var selTransp = _freteSelectTransportadora(c.transportadora || '', function() { freteEditarContrato(i, 'transportadora', this.value); });
    tdTransp.appendChild(selTransp); tr.appendChild(tdTransp);

    var tdTipo = document.createElement('td');
    tdTipo.style.cssText = 'padding:7px 8px;';
    var sel = document.createElement('select');
    sel.style.cssText = _freteInputStyle('width:100%;');
    Object.entries(FRETE_TIPOS).forEach(function(entry) {
      var opt = document.createElement('option');
      opt.value = entry[0]; opt.textContent = entry[1];
      if (c.tipo === entry[0]) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.onchange = function() { freteEditarContrato(i, 'tipo', this.value); };
    tdTipo.appendChild(sel); tr.appendChild(tdTipo);

    var tdKmModo = document.createElement('td');
    tdKmModo.style.cssText = 'padding:7px 8px;';
    var selKmModo = document.createElement('select');
    selKmModo.title = 'Define se o KM considera ida e volta ou somente a ida';
    selKmModo.style.cssText = _freteInputStyle('width:100%;');
    [['ida_volta','Ida e volta'], ['ida','Somente ida']].forEach(function(opt) {
      var o = document.createElement('option');
      o.value = opt[0]; o.textContent = opt[1];
      if ((c.kmModo || 'ida_volta') === opt[0]) o.selected = true;
      selKmModo.appendChild(o);
    });
    selKmModo.onchange = function() { freteEditarContrato(i, 'kmModo', this.value); };
    tdKmModo.appendChild(selKmModo); tr.appendChild(tdKmModo);

    var tdFixo = document.createElement('td');
    tdFixo.style.cssText = 'padding:7px 8px;text-align:right;';
    var inpFixo = document.createElement('input');
    inpFixo.type = 'number'; inpFixo.value = c.fixo || ''; inpFixo.placeholder = '0,00';
    inpFixo.min = '0'; inpFixo.step = '0.01';
    inpFixo.style.cssText = _freteInputStyle('width:100%;text-align:right;');
    inpFixo.disabled = (c.tipo === 'diaria' || c.tipo === 'spot');
    if (inpFixo.disabled) inpFixo.style.opacity = '0.4';
    inpFixo.onchange = function() { freteEditarContrato(i, 'fixo', this.value); };
    tdFixo.appendChild(inpFixo); tr.appendChild(tdFixo);

    var tdKm = document.createElement('td');
    tdKm.style.cssText = 'padding:7px 8px;text-align:right;';
    var inpKm = document.createElement('input');
    inpKm.type = 'number'; inpKm.value = c.km || ''; inpKm.placeholder = '0,00';
    inpKm.min = '0'; inpKm.step = '0.01';
    inpKm.style.cssText = _freteInputStyle('width:100%;text-align:right;');
    inpKm.disabled = (c.tipo !== 'fixo_km');
    if (inpKm.disabled) inpKm.style.opacity = '0.4';
    inpKm.onchange = function() { freteEditarContrato(i, 'km', this.value); };
    tdKm.appendChild(inpKm); tr.appendChild(tdKm);

    var tdM3 = document.createElement('td');
    tdM3.style.cssText = 'padding:7px 8px;text-align:right;';
    var inpM3 = document.createElement('input');
    inpM3.type = 'number'; inpM3.value = c.m3 || ''; inpM3.placeholder = '0,00';
    inpM3.min = '0'; inpM3.step = '0.01';
    inpM3.style.cssText = _freteInputStyle('width:100%;text-align:right;');
    inpM3.disabled = (c.tipo !== 'fixo_m3');
    if (inpM3.disabled) inpM3.style.opacity = '0.4';
    inpM3.onchange = function() { freteEditarContrato(i, 'm3', this.value); };
    tdM3.appendChild(inpM3); tr.appendChild(tdM3);

    var tdDia = document.createElement('td');
    tdDia.style.cssText = 'padding:7px 8px;text-align:right;';
    var inpDia = document.createElement('input');
    inpDia.type = 'number'; inpDia.value = c.diaria || ''; inpDia.placeholder = '0,00';
    inpDia.min = '0'; inpDia.step = '0.01';
    inpDia.style.cssText = _freteInputStyle('width:100%;text-align:right;');
    inpDia.disabled = (c.tipo !== 'diaria');
    if (inpDia.disabled) inpDia.style.opacity = '0.4';
    inpDia.onchange = function() { freteEditarContrato(i, 'diaria', this.value); };
    tdDia.appendChild(inpDia); tr.appendChild(tdDia);

    var tdAcao = document.createElement('td');
    tdAcao.style.cssText = 'padding:7px 6px;text-align:center;';
    var btnDel = document.createElement('button');
    btnDel.textContent = '✕';
    btnDel.style.cssText = 'font-size:11px;padding:3px 9px;border:1px solid #ef4444;border-radius:5px;background:transparent;color:#ef4444;cursor:pointer;';
    btnDel.onclick = (function(idx) { return function() { freteRemoverContrato(idx); }; })(i);
    tdAcao.appendChild(btnDel); tr.appendChild(tdAcao);

    tbody.appendChild(tr);
  });
}

function freteRenderSpot() {
  const tbody = document.getElementById('frete-spot-body');
  if (!tbody) return;
  const spots = freteCarregarSpot();
  if (!spots.length) {
    tbody.innerHTML = '<tr><td colspan="5" style="padding:14px;text-align:center;color:var(--text-3);font-size:12px;">Nenhuma rota spot cadastrada.</td></tr>';
    return;
  }
  tbody.innerHTML = '';
  spots.forEach(function(s, i) {
    var tr = document.createElement('tr');
    tr.style.borderTop = '1px solid var(--border-dk)';

    var tdOrig = document.createElement('td'); tdOrig.style.padding = '8px 12px';
    var inpOrig = document.createElement('input');
    inpOrig.value = s.origem || ''; inpOrig.placeholder = 'Terminal de origem';
    inpOrig.style.cssText = _freteInputStyle('width:100%;');
    inpOrig.onchange = function() { freteEditarSpot(i, 'origem', this.value); };
    tdOrig.appendChild(inpOrig); tr.appendChild(tdOrig);

    var tdDest = document.createElement('td'); tdDest.style.padding = '8px 12px';
    var inpDest = document.createElement('input');
    inpDest.value = s.destino || ''; inpDest.placeholder = 'Cidade ou cliente destino';
    inpDest.style.cssText = _freteInputStyle('width:100%;');
    inpDest.onchange = function() { freteEditarSpot(i, 'destino', this.value); };
    tdDest.appendChild(inpDest); tr.appendChild(tdDest);

    var tdTranspSpot = document.createElement('td'); tdTranspSpot.style.padding = '8px 12px';
    var selTranspSpot = _freteSelectTransportadora(s.transportadora || '', function() { freteEditarSpot(i, 'transportadora', this.value); });
    tdTranspSpot.appendChild(selTranspSpot); tr.appendChild(tdTranspSpot);

    var tdVal = document.createElement('td');
    tdVal.style.cssText = 'padding:8px 12px;text-align:right;';
    var inpVal = document.createElement('input');
    inpVal.type = 'number'; inpVal.value = s.valor || ''; inpVal.placeholder = '0,00';
    inpVal.min = '0'; inpVal.step = '0.01';
    inpVal.style.cssText = _freteInputStyle('width:100%;text-align:right;');
    inpVal.onchange = function() { freteEditarSpot(i, 'valor', this.value); };
    tdVal.appendChild(inpVal); tr.appendChild(tdVal);

    var tdAcao = document.createElement('td');
    tdAcao.style.cssText = 'padding:8px 12px;text-align:center;';
    var btnDel = document.createElement('button');
    btnDel.textContent = '✕';
    btnDel.style.cssText = 'font-size:11px;padding:3px 9px;border:1px solid #ef4444;border-radius:5px;background:transparent;color:#ef4444;cursor:pointer;';
    btnDel.onclick = (function(idx) { return function() { freteRemoverSpot(idx); }; })(i);
    tdAcao.appendChild(btnDel); tr.appendChild(tdAcao);

    tbody.appendChild(tr);
  });
}

// ── CRUD ─────────────────────────────────────────────────────────────────────
function freteAdicionarContrato() {
  var arr = freteCarregarContratos();
  var placasCadastradas = new Set(arr.map(function(c) { return c.placa; }));
  var proxV = veiculos.find(function(v) { return !placasCadastradas.has(v.placa); });
  arr.push({ placa: proxV ? proxV.placa : '', transportadora: proxV ? (proxV.transportadora||'') : '', tipo:'fixo_km', kmModo:'ida_volta', fixo:'', km:'', m3:'', diaria:'' });
  freteSalvarContratos(arr); freteRenderContratos();
}

function freteEditarContrato(i, campo, valor) {
  var arr = freteCarregarContratos();
  if (!arr[i]) return;
  arr[i][campo] = valor;
  if (campo === 'placa') {
    var v = veiculos.find(function(v) { return v.placa === valor; });
    if (v) arr[i].transportadora = v.transportadora || arr[i].transportadora;
  }
  freteSalvarContratos(arr);
  if (campo === 'tipo') freteRenderContratos();
}

function freteRemoverContrato(i) {
  var arr = freteCarregarContratos();
  arr.splice(i, 1);
  freteSalvarContratos(arr); freteRenderContratos();
}

function freteAdicionarSpot() {
  var arr = freteCarregarSpot();
  arr.push({ origem:'', destino:'', transportadora:'', valor:'' });
  freteSalvarSpot(arr); freteRenderSpot();
}

function freteEditarSpot(i, campo, valor) {
  var arr = freteCarregarSpot();
  if (!arr[i]) return;
  arr[i][campo] = valor;
  freteSalvarSpot(arr);
}

function freteRemoverSpot(i) {
  var arr = freteCarregarSpot();
  arr.splice(i, 1);
  freteSalvarSpot(arr); freteRenderSpot();
}

// ── Vista e Ordem ─────────────────────────────────────────────────────────────
var _freteVista = 'viagem';
var _freteOrdem = 'custo';

function freteSetVista(v) {
  _freteVista = v;
  ['viagem','dia','semana','mes','periodo'].forEach(function(k) {
    var btn = document.getElementById('frete-btn-' + k);
    if (!btn) return;
    if (k === v) { btn.style.background='var(--pet-green)'; btn.style.color='#000'; btn.style.fontWeight='700'; btn.style.border='none'; }
    else { btn.style.background=''; btn.style.color=''; btn.style.fontWeight=''; btn.style.border=''; }
  });
  freteCalcular();
}

function freteSetOrdem(o) {
  _freteOrdem = o;
  ['custo','transportadora','placa'].forEach(function(k) {
    var btn = document.getElementById('frete-ordem-' + k);
    if (!btn) return;
    if (k === o) { btn.style.background='var(--pet-green)'; btn.style.color='#000'; btn.style.fontWeight='700'; btn.style.border='none'; }
    else { btn.style.background=''; btn.style.color=''; btn.style.fontWeight=''; btn.style.border=''; }
  });
  freteCalcular();
}

function freteToggleGrupo(groupId) {
  var icon = document.getElementById(groupId + '-icon');
  var rows = document.querySelectorAll('.' + groupId + '-row');
  if (!rows.length) return;
  var aberto = rows[0].style.display !== 'none';
  rows.forEach(function(r) { r.style.display = aberto ? 'none' : 'table-row'; });
  if (icon) icon.textContent = aberto ? '▸' : '▾';
}

async function freteCalcular() {
  var el = document.getElementById('frete-resultado');
  var resumoEl = document.getElementById('frete-resumo-total');
  if (!el) return;
  var de  = (document.getElementById('frete-de')  || {}).value || '';
  var ate = (document.getElementById('frete-ate') || {}).value || '';
  var contratos = freteCarregarContratos();
  var spots     = freteCarregarSpot();

  if (!dirHandleHistorico) {
    el.innerHTML = '<div style="padding:32px;text-align:center;color:var(--text-3);font-size:13px;">Selecione a pasta do Histórico na aba Histórico para carregar os dados.</div>';
    return;
  }
  el.innerHTML = '<div style="padding:32px;text-align:center;color:var(--text-3);font-size:13px;">Carregando dados do histórico...</div>';

  var snaps = [];
  try {
    for await (var [name, handle] of dirHandleHistorico.entries()) {
      if (handle.kind !== 'file' || !name.endsWith('.json')) continue;
      var diaKey = name.slice(0, 8);
      if (!/^\d{8}$/.test(diaKey)) continue;
      if (de  && diaKey < de.replace(/-/g,''))  continue;
      if (ate && diaKey > ate.replace(/-/g,'')) continue;
      try {
        var file = await handle.getFile();
        var data = JSON.parse(await file.text());
        if (!(data.resultado || data.pedidos || data.versao)) continue;
        if (!data.savedAt) data.savedAt = new Date(file.lastModified).toISOString();
        snaps.push(data);
      } catch(e) {}
    }
  } catch(e) {
    el.innerHTML = '<div style="padding:32px;text-align:center;color:#ef4444;font-size:13px;">Erro ao ler histórico: ' + e.message + '</div>';
    return;
  }

  if (!snaps.length) {
    el.innerHTML = '<div style="padding:32px;text-align:center;color:var(--text-3);font-size:13px;">Nenhuma roteirização encontrada no período selecionado.</div>';
    return;
  }

  var viagensMap = {}, mesMapa = {};

  snaps.forEach(function(snap) {
    var res = snap.resultado || {}, veics = snap.veiculos || [];
    var dataSnap = (snap.savedAt || '').slice(0,10);
    var mesKey = dataSnap.slice(0,7);
    var diaKey2 = dataSnap.replace(/-/g,'').slice(0,8);
    veics.forEach(function(v) {
      var placa = v.placa;
      if (!placa) return;
      var viagens = (res[v.id] || []).filter(function(vi) { return !vi._vazio && (vi.paradas||[]).length; });
      if (!viagens.length) return;
      if (!viagensMap[placa]) viagensMap[placa] = [];
      if (!mesMapa[placa]) mesMapa[placa] = {};
      if (!mesMapa[placa][mesKey]) mesMapa[placa][mesKey] = { viagens: [] };
      viagens.forEach(function(vi) {
        var kmIda = vi.paradas.reduce(function(s,p) { return s + (p.distanciaKm||0); }, 0);
        var m3Total = vi.paradas.reduce(function(s,p) { return s + (p.volumeTotal||0); }, 0);
        var termOrigem = vi.terminalOrigem || '';
        var destinos = Array.from(new Set(vi.paradas.map(function(p) { return p.pedido ? (p.pedido.cidade||p.pedido.cliente||'') : ''; }))).join(', ');
        var entry = { data: dataSnap, diaKey: diaKey2, mesKey: mesKey, kmIda: kmIda, m3Total: m3Total, termOrigem: termOrigem, destinos: destinos, placa: placa };
        viagensMap[placa].push(entry);
        mesMapa[placa][mesKey].viagens.push(entry);
      });
    });
  });

  function kmEfetivo(entry, contrato) {
    var modo = (contrato && contrato.kmModo) || 'ida_volta';
    return modo === 'ida' ? entry.kmIda : entry.kmIda * 2;
  }

  function custoViagem(entry, contrato) {
    var mes = mesMapa[entry.placa] && mesMapa[entry.placa][entry.mesKey];
    var nViagMes = mes ? mes.viagens.length : 1;
    var fixo = parseFloat(contrato.fixo) || 0;
    var fixoRateado = fixo / Math.max(nViagMes, 1);
    var km = kmEfetivo(entry, contrato);
    if (contrato.tipo === 'fixo_km') return fixoRateado + (parseFloat(contrato.km)||0) * km;
    if (contrato.tipo === 'fixo_m3') return fixoRateado + (parseFloat(contrato.m3)||0) * entry.m3Total;
    if (contrato.tipo === 'diaria')  return parseFloat(contrato.diaria) || 0;
    if (contrato.tipo === 'spot') {
      var sp = spots.find(function(s) {
        var bateRota = entry.termOrigem.toLowerCase().includes((s.origem||'').toLowerCase()) &&
                       entry.destinos.toLowerCase().includes((s.destino||'').toLowerCase());
        var bateTransp = !s.transportadora || s.transportadora === contrato.transportadora;
        return bateRota && bateTransp;
      });
      return sp ? (parseFloat(sp.valor)||0) * entry.m3Total : 0;
    }
    return 0;
  }

  function chaveVista(entry) {
    if (_freteVista === 'viagem') return entry.placa + '__' + entry.data + '__' + entry.termOrigem + '__' + entry.destinos;
    if (_freteVista === 'dia')    return entry.placa + '__' + entry.diaKey;
    if (_freteVista === 'semana') {
      var d = new Date(entry.data), jan1 = new Date(d.getFullYear(),0,1);
      var sem = Math.ceil((((d - jan1)/86400000) + jan1.getDay() + 1) / 7);
      return entry.placa + '__' + d.getFullYear() + '-S' + String(sem).padStart(2,'0');
    }
    if (_freteVista === 'mes') return entry.placa + '__' + entry.mesKey;
    return entry.placa;
  }

  function labelLegivel(g) {
    var parts = g.label.split('__');
    if (_freteVista === 'viagem') return (parts[1]||'') + ' · ' + (parts[2]||'').split(' ')[0] + ' → ' + (parts[3]||'').slice(0,30);
    if (_freteVista === 'dia')    return (parts[1]||'').replace(/(\d{4})(\d{2})(\d{2})/, '$3/$2/$1');
    if (_freteVista === 'semana') return parts[1]||'';
    if (_freteVista === 'mes')    { var ym = (parts[1]||'').split('-'); return (ym[1]||'') + '/' + (ym[0]||''); }
    return 'Total do período';
  }

  var grupos = {}, totalGeral = 0;

  Object.keys(viagensMap).forEach(function(placa) {
    var entradas = viagensMap[placa];
    var contrato = contratos.find(function(c) { return c.placa === placa; });
    entradas.forEach(function(entry) {
      var custo = contrato ? custoViagem(entry, contrato) : 0;
      var kmDisplay = kmEfetivo(entry, contrato);
      var chave = chaveVista(entry);
      if (!grupos[chave]) {
        var transp = contrato ? contrato.transportadora : (veiculos.find(function(v) { return v.placa === placa; }) || {}).transportadora || '—';
        grupos[chave] = { placa: placa, transportadora: transp, custo: 0, km: 0, m3: 0, nViagens: 0, label: chave };
      }
      grupos[chave].custo    += custo;
      grupos[chave].km       += kmDisplay;
      grupos[chave].m3       += entry.m3Total;
      grupos[chave].nViagens += 1;
      totalGeral             += custo;
    });
  });

  var linhas = Object.values(grupos);
  var fmt = function(v) { return 'R$ ' + v.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}); };

  if (resumoEl) {
    var totalKm = linhas.reduce(function(s,g){return s+g.km;},0);
    var totalM3 = linhas.reduce(function(s,g){return s+g.m3;},0);
    var totalVi = linhas.reduce(function(s,g){return s+g.nViagens;},0);
    resumoEl.innerHTML = [
      ['💰 Custo Total', fmt(totalGeral)],
      ['🚛 Viagens', totalVi],
      ['📏 KM Total (est.)', totalKm.toFixed(0) + ' km'],
      ['📦 Volume Total', totalM3.toFixed(1) + ' m³'],
      ['📊 Custo Médio/Viagem', totalVi ? fmt(totalGeral/totalVi) : '—'],
    ].map(function(c) {
      return '<div style="background:rgba(0,0,0,0.03);border:1px solid var(--border-dk);border-radius:8px;padding:10px 18px;min-width:140px;">' +
        '<div style="font-size:10px;color:var(--text-3);letter-spacing:.06em;font-weight:600;margin-bottom:4px;">' + c[0] + '</div>' +
        '<div style="font-size:16px;font-weight:700;color:var(--text);">' + c[1] + '</div></div>';
    }).join('');
  }

  var vistas = { viagem:'Viagem', dia:'Dia', semana:'Semana', mes:'Mês', periodo:'Período' };
  var thLabel = (vistas[_freteVista] || _freteVista).toUpperCase();
  var ordemBtnStyle = function(ativo) {
    return 'font-size:10px;padding:3px 10px;border-radius:6px;cursor:pointer;' +
      (ativo ? 'background:var(--pet-green);color:#000;font-weight:700;border:none;' : 'background:transparent;color:var(--text-3);border:1px solid var(--border-dk);');
  };

  var html =
    '<div style="display:flex;align-items:center;gap:6px;padding:8px 14px;border-bottom:1px solid var(--border-dk);">' +
    '<span style="font-size:10px;color:var(--text-3);font-weight:600;letter-spacing:.04em;">ORDENAR POR</span>' +
    '<button id="frete-ordem-custo" onclick="freteSetOrdem(\'custo\')" style="' + ordemBtnStyle(_freteOrdem==='custo') + '">Custo</button>' +
    '<button id="frete-ordem-transportadora" onclick="freteSetOrdem(\'transportadora\')" style="' + ordemBtnStyle(_freteOrdem==='transportadora') + '">Transportadora</button>' +
    '<button id="frete-ordem-placa" onclick="freteSetOrdem(\'placa\')" style="' + ordemBtnStyle(_freteOrdem==='placa') + '">Placa</button>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:12px;">' +
    '<thead><tr style="background:rgba(0,0,0,0.03);border-bottom:2px solid var(--border-dk);">' +
    '<th style="padding:10px 14px;text-align:left;font-size:11px;font-weight:600;color:var(--text-3);">PLACA</th>' +
    '<th style="padding:10px 14px;text-align:left;font-size:11px;font-weight:600;color:var(--text-3);">TRANSPORTADORA</th>' +
    '<th style="padding:10px 14px;text-align:left;font-size:11px;font-weight:600;color:var(--text-3);">' + thLabel + '</th>' +
    '<th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--text-3);">VIAGENS</th>' +
    '<th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--text-3);">KM EST.</th>' +
    '<th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--text-3);">VOLUME (m³)</th>' +
    '<th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--text-3);">CUSTO</th>' +
    '</tr></thead><tbody>';

  if (linhas.length) {
    var porPlaca = {}, ordemPlacas = [];
    linhas.forEach(function(g) {
      if (!porPlaca[g.placa]) { porPlaca[g.placa] = []; ordemPlacas.push(g.placa); }
      porPlaca[g.placa].push(g);
    });
    ordemPlacas.sort(function(a,b) {
      if (_freteOrdem === 'transportadora') {
        var ta = (porPlaca[a][0].transportadora||'').toLowerCase();
        var tb = (porPlaca[b][0].transportadora||'').toLowerCase();
        if (ta !== tb) return ta < tb ? -1 : 1;
        return a < b ? -1 : (a > b ? 1 : 0);
      }
      if (_freteOrdem === 'placa') return a < b ? -1 : (a > b ? 1 : 0);
      var sa = porPlaca[a].reduce(function(s,g){return s+g.custo;},0);
      var sb = porPlaca[b].reduce(function(s,g){return s+g.custo;},0);
      return sb - sa;
    });
    ordemPlacas.forEach(function(placa, gi) {
      var itens = porPlaca[placa];
      var custoPlaca = itens.reduce(function(s,g){return s+g.custo;},0);
      var kmPlaca    = itens.reduce(function(s,g){return s+g.km;},0);
      var m3Placa    = itens.reduce(function(s,g){return s+g.m3;},0);
      var viPlaca    = itens.reduce(function(s,g){return s+g.nViagens;},0);
      var transpPlaca = itens[0].transportadora || '';
      var groupId = 'frete-grp-' + gi;
      html += '<tr onclick="freteToggleGrupo(\'' + groupId + '\')" style="border-top:1px solid var(--border-dk);cursor:pointer;background:rgba(0,0,0,0.025);">' +
        '<td style="padding:10px 14px;font-weight:700;color:var(--text);white-space:nowrap;">' +
        '<span id="' + groupId + '-icon" style="display:inline-block;width:16px;text-align:center;margin-right:6px;color:var(--text-3);">▸</span>' + placa + '</td>' +
        '<td style="padding:10px 14px;color:var(--text-2);">' + transpPlaca + '</td>' +
        '<td style="padding:10px 14px;color:var(--text-3);font-style:italic;">' + itens.length + ' registro' + (itens.length>1?'s':'') + '</td>' +
        '<td style="padding:10px 14px;text-align:right;color:var(--text-2);">' + viPlaca + '</td>' +
        '<td style="padding:10px 14px;text-align:right;color:var(--text-2);">' + kmPlaca.toFixed(0) + ' km</td>' +
        '<td style="padding:10px 14px;text-align:right;color:var(--text-2);">' + m3Placa.toFixed(1) + '</td>' +
        '<td style="padding:10px 14px;text-align:right;font-weight:700;color:var(--pet-green,#84cc16);">' + fmt(custoPlaca) + '</td></tr>';
      itens.forEach(function(g, ii) {
        html += '<tr class="' + groupId + '-row" style="display:none;border-top:1px solid var(--border-dk);' + (ii%2===1?'background:rgba(0,0,0,0.015)':'') + '">' +
          '<td style="padding:8px 14px 8px 36px;color:var(--text-3);font-size:11px;">' + (g.placa||'') + '</td>' +
          '<td style="padding:8px 14px;color:var(--text-2);">' + (g.transportadora||'') + '</td>' +
          '<td style="padding:8px 14px;color:var(--text-2);max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + labelLegivel(g) + '</td>' +
          '<td style="padding:8px 14px;text-align:right;color:var(--text-2);">' + g.nViagens + '</td>' +
          '<td style="padding:8px 14px;text-align:right;color:var(--text-2);">' + g.km.toFixed(0) + ' km</td>' +
          '<td style="padding:8px 14px;text-align:right;color:var(--text-2);">' + g.m3.toFixed(1) + '</td>' +
          '<td style="padding:8px 14px;text-align:right;font-weight:700;color:var(--pet-green,#84cc16);">' + fmt(g.custo) + '</td></tr>';
      });
    });
  } else {
    html += '<tr><td colspan="7" style="padding:32px;text-align:center;color:var(--text-3);">Nenhum dado calculado. Verifique os contratos cadastrados.</td></tr>';
  }

  html += '<tr style="border-top:2px solid var(--border-dk);background:rgba(0,0,0,0.03);">' +
    '<td colspan="6" style="padding:10px 14px;font-weight:700;color:var(--text);text-align:right;">TOTAL GERAL</td>' +
    '<td style="padding:10px 14px;text-align:right;font-weight:800;font-size:14px;color:var(--pet-green,#84cc16);">' + fmt(totalGeral) + '</td>' +
    '</tr></tbody></table>';

  el.innerHTML = html;
}

window.freteCalcular          = freteCalcular;
window.freteAdicionarContrato = freteAdicionarContrato;
window.freteEditarContrato    = freteEditarContrato;
window.freteRemoverContrato   = freteRemoverContrato;
window.freteAdicionarSpot     = freteAdicionarSpot;
window.freteEditarSpot        = freteEditarSpot;
window.freteRemoverSpot       = freteRemoverSpot;
window.freteSetVista          = freteSetVista;
window.freteSetOrdem          = freteSetOrdem;
window.freteToggleGrupo       = freteToggleGrupo;
