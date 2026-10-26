// ═══════════════════════════════════════════════════════════════════════════
// OPERAÇÃO SOB DEMANDA — ciclos dos caminhões (substitui a planilha de reporte)
// ═══════════════════════════════════════════════════════════════════════════
// Cada ciclo = uma viagem de uma placa sob demanda: carrega em SP, entrega no
// RJ, volta. Fica no Firestore (coleção sob_ciclos) para qualquer computador
// ver. Origem do ciclo:
//   • roteirizador → criado/atualizado sozinho quando uma programação com
//     placa sob demanda é salva/revisada no Histórico;
//   • previa       → lançado à mão (pedidos vistos no SAP, ainda não subidos);
//                    quando a viagem real é salva no roteirizador (mesma placa,
//                    carregamento até 3 dias de diferença) a prévia é
//                    substituída pelos dados reais, mantendo nº, status e marcos.
// Nº do ciclo: sequencial por placa, sem reiniciar. config/sob_config.base[PLACA]
// guarda o nº do último ciclo finalizado; os ciclos ativos da placa são
// numerados em ordem de carregamento a partir dele.
// Status (só equipe NEXTA): Programado → Carregamento → Deslocamento p/ RJ →
// Descarga → Deslocamento p/ SP → Finalizado. Estado do caminhão sem ciclo
// seguinte: Sem pedido / DSR / Manutenção (config/sob_config.veiculos).
// ═══════════════════════════════════════════════════════════════════════════
(function () {
'use strict';

const STATUS = [
  { id: 'programado',   label: 'PROGRAMADO',         bg: '#F2CEEF', fg: '#5B2157' },
  { id: 'carregamento', label: 'CARREGAMENTO',       bg: '#FFE699', fg: '#5C4300' },
  { id: 'desloc_rj',    label: 'DESLOCAMENTO P/ RJ', bg: '#9DD9F3', fg: '#0B3A55' },
  { id: 'descarga',     label: 'DESCARGA',           bg: '#C6EFCE', fg: '#14532D' },
  { id: 'desloc_sp',    label: 'DESLOCAMENTO P/ SP', bg: '#D8F08A', fg: '#3B4A06' },
  { id: 'finalizado',   label: 'FINALIZADO',         bg: '#D9D9D9', fg: '#374151' },
];
const ESTADOS = [
  { id: 'sem_pedido', label: 'SEM PEDIDO', bg: '#F8CBAD', fg: '#7C2D12' },
  { id: 'dsr',        label: 'DSR',        bg: '#F4CCCC', fg: '#7F1D1D' },
  { id: 'manutencao', label: 'MANUTENÇÃO', bg: '#F4B183', fg: '#6B2A06' },
];
const TURNOS = ['MANHÃ', 'TARDE', 'NOITE'];
const NOTAS_PROX = ['', 'MANUTENÇÃO', 'DSR', 'SEM PREVISÃO'];
const MARCOS = [
  { id: 'oferta',           label: 'Oferta do caminhão',   auto: 'Automático: oferta no Painel' },
  { id: 'dispCaminhao',     label: 'Disp. caminhão',       auto: 'Automático: "disponível a partir de" da oferta' },
  { id: 'liberacaoProduto', label: 'Liberação do produto', auto: '' },
  { id: 'envioPedido',      label: 'Envio do pedido',      auto: 'Automático ao enviar a OR pelo 📲' },
  { id: 'liberacaoPedido',  label: 'Liberação do pedido',  auto: '' },
];
const stById = Object.fromEntries(STATUS.map(s => [s.id, s]));
const estById = Object.fromEntries(ESTADOS.map(s => [s.id, s]));

// ── Firestore ───────────────────────────────────────────────────────────────
const fb = () => ({ db: window.fbDb, doc: window.fbDoc, getDoc: window.fbGetDoc, setDoc: window.fbSetDoc,
  col: window.fbCollection, getDocs: window.fbGetDocs, query: window.fbQuery, where: window.fbWhere });
const refCiclo = id => fb().doc(fb().db, 'sob_ciclos', id);
const refCfg = () => fb().doc(fb().db, 'config', 'sob_config');

// ── Estado da tela ──────────────────────────────────────────────────────────
const st = {
  ciclos: new Map(), cfg: null, placas: [], carregadoEm: 0, periodoCarregado: null,
  periodo: 30, filtro: '', transp: '', marcos: false, finalAberto: false,
};
let _body = null;

// ── Utilidades ──────────────────────────────────────────────────────────────
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const jsq = v => String(v ?? '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
const normPlaca = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const semAcento = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '');
const p2 = n => String(n).padStart(2, '0');
const isoHoje = (d = 0) => { const x = new Date(); x.setDate(x.getDate() + d); return `${x.getFullYear()}-${p2(x.getMonth() + 1)}-${p2(x.getDate())}`; };
const br = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? `${m[3]}/${m[2]}/${m[1]}` : ''; };
const brCurto = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? `${m[3]}/${m[2]}` : ''; };
const brDataHora = v => { if (!v) return ''; const d = new Date(v); if (isNaN(d)) return br(v); return `${p2(d.getDate())}/${p2(d.getMonth() + 1)} ${p2(d.getHours())}:${p2(d.getMinutes())}`; };
const localDT = v => { if (!v) return ''; const d = new Date(v); if (isNaN(d)) return ''; return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}T${p2(d.getHours())}:${p2(d.getMinutes())}`; };
const fmtVol = v => { const n = Math.round((Number(v) || 0) * 10) / 10; return Number.isInteger(n) ? String(n) : n.toFixed(1).replace('.', ','); };
const usuario = () => { const S = window.S || {}; return window.USERS_DB?.[S.user]?.name || S.user || ''; };
const toast = (m, ok = true) => (window.showToast ? window.showToast(m, ok) : console.log(m));
const ehOperacional = () => { const S = window.S || {}; const r = window.USERS_DB?.[S.user]?.role; return r === 'admin' || r === 'operacional'; };

// Sigla do produto (como no reporte: GC, GA, EA, S10, S500, ETANOL…)
function siglaAuto(nome) {
  const n = semAcento(nome).toUpperCase();
  if (!n) return '';
  if (/^(GC|GA|EA|S10|S500|ETANOL)$/.test(n.trim())) return n.trim();
  if (/S[\s-]?500/.test(n)) return 'S500';
  if (/S[\s-]?10\b|S10/.test(n)) return 'S10';
  if (/ETANOL|ALCOOL|\bEHC\b|\bEHR\b/.test(n)) return /ADIT|PREMIUM|GRID|\bADT\b/.test(n) ? 'EA' : 'ETANOL';
  if (/GASOLINA|\bGAS\b/.test(n)) return /ADIT|PREMIUM|PODIUM|GRID|POWER|\bADT\b/.test(n) ? 'GA' : 'GC';
  if (/DIESEL/.test(n)) return 'S10';
  return nome.trim().slice(0, 12).toUpperCase();
}
const sigla = nome => (st.cfg?.siglas?.[nome]) || siglaAuto(nome);
const apelido = cli => (st.cfg?.apelidos?.[cli]) || cli;

// Ciclos importados da planilha guardam o volume/produto/texto de clientes como
// estavam lá (o texto da planilha nem sempre tem volume por cliente).
const volCiclo = c => (Number(c.volumeInformado) > 0) ? Number(c.volumeInformado) : (c.entregas || []).reduce((s, e) => s + (e.itens || []).reduce((a, it) => a + (Number(it.volume) || 0), 0), 0);
function produtoCiclo(c) {
  if (c.produtoInformado) return c.produtoInformado;
  const s = new Set();
  (c.entregas || []).forEach(e => (e.itens || []).forEach(it => { if ((Number(it.volume) || 0) > 0) s.add(sigla(it.produto)); }));
  return s.size === 0 ? '' : (s.size === 1 ? [...s][0] : 'MISTO');
}
// "TREMENDÃO (5m³ GC + 20m³ S10) | CUTUBA (10m³ S10)"
function clientesTexto(c) {
  if (c.clientesTextoOriginal) return c.clientesTextoOriginal;
  return (c.entregas || []).map(e => {
    const por = new Map();
    (e.itens || []).forEach(it => { const sg = sigla(it.produto); por.set(sg, (por.get(sg) || 0) + (Number(it.volume) || 0)); });
    const partes = [...por].filter(([, v]) => v > 0).map(([sg, v]) => `${fmtVol(v)}m³ ${sg}`);
    return `${apelido(e.cliente || '—')}${partes.length ? ` (${partes.join(' + ')})` : ''}`;
  }).join(' | ');
}
const carregTexto = c => c.dataCarregamento ? `${brCurto(c.dataCarregamento)} às ${c.horaCarregamento ? c.horaCarregamento + 'h' : '--:--'}` : '-';
const etaTexto = c => c.etaData ? `${brCurto(c.etaData)}${c.etaTurno ? ' | ' + c.etaTurno : ''}` : '-';
const proxTexto = c => c.proxCargaNota || (c.proxCarga ? br(c.proxCarga) : '');
const ordemCiclo = (a, b) => (a.dataCarregamento || '9999').localeCompare(b.dataCarregamento || '9999')
  || (a.horaCarregamento || '99').localeCompare(b.horaCarregamento || '99')
  || String(a.criadoEm || '').localeCompare(String(b.criadoEm || ''));

// ── Carga dos dados ─────────────────────────────────────────────────────────
async function carregarCfg(force) {
  if (st.cfg && !force) return st.cfg;
  const s = await fb().getDoc(refCfg());
  st.cfg = s.exists() ? s.data() : {};
  st.cfg.siglas = st.cfg.siglas || {}; st.cfg.apelidos = st.cfg.apelidos || {};
  st.cfg.base = st.cfg.base || {}; st.cfg.veiculos = st.cfg.veiculos || {};
  return st.cfg;
}
async function salvarCfg(patch) {
  await fb().setDoc(refCfg(), { ...patch, atualizadoEm: new Date().toISOString(), atualizadoPor: usuario() }, { merge: true });
  Object.entries(patch).forEach(([k, v]) => {
    if (v && typeof v === 'object' && !Array.isArray(v)) st.cfg[k] = { ...(st.cfg[k] || {}), ...v };
    else st.cfg[k] = v;
  });
}
async function carregarAtivos() {
  const { col, db, getDocs, query, where } = fb();
  const snap = await getDocs(query(col(db, 'sob_ciclos'), where('finalizado', '==', false)));
  const ativos = [];
  snap.forEach(d => { const c = { id: d.id, ...d.data() }; st.ciclos.set(c.id, c); ativos.push(c); });
  return ativos;
}
async function carregarTudo(force) {
  if (!force && st.carregadoEm && Date.now() - st.carregadoEm < 120000 && st.periodoCarregado === st.periodo) return;
  const { col, db, getDocs, query, where } = fb();
  st.ciclos = new Map();
  const [, snapPer] = await Promise.all([
    carregarAtivos(),
    getDocs(query(col(db, 'sob_ciclos'), where('dataCarregamento', '>=', isoHoje(-st.periodo)))),
    carregarCfg(true),
  ]);
  snapPer.forEach(d => st.ciclos.set(d.id, { id: d.id, ...d.data() }));
  const allP = await window.dbGetPlates();
  st.placas = [];
  Object.keys(allP || {}).forEach(carrier => (allP[carrier] || []).forEach(p => {
    if (p.ativo !== false && window.ehSobDemanda?.(p)) st.placas.push({ carrier, placa: p.placa, capacidade: p.capacidade, bases: (p.basesCarregamento && p.basesCarregamento.length) ? p.basesCarregamento : [p.operacao].filter(Boolean) });
  }));
  st.placas.sort((a, b) => a.carrier.localeCompare(b.carrier, 'pt-BR') || a.placa.localeCompare(b.placa));
  st.carregadoEm = Date.now(); st.periodoCarregado = st.periodo;
}

// Grava alterações num ciclo (merge) + registro de quem fez o quê.
async function gravarCiclo(id, patch, acao) {
  const atual = st.ciclos.get(id) || {};
  const agora = new Date().toISOString();
  const log = acao ? [...(atual.log || []).slice(-49), { em: agora, por: usuario(), acao }] : (atual.log || []);
  const dados = { ...patch, log, atualizadoEm: agora, atualizadoPor: usuario() };
  await fb().setDoc(refCiclo(id), dados, { merge: true });
  const novo = { ...atual, ...dados, id };
  if (patch.marcos) novo.marcos = { ...(atual.marcos || {}), ...patch.marcos };
  st.ciclos.set(id, novo);
  return novo;
}

// Renumera os ciclos ativos (não finalizados/cancelados) de uma placa, em
// ordem de carregamento, a partir do nº do último finalizado.
async function renumerarPlaca(placaN, ativos) {
  if (!placaN) return;
  await carregarCfg();
  const base = Number(st.cfg.base?.[placaN]) || 0;
  const lista = (ativos || [...st.ciclos.values()]).filter(c => normPlaca(c.placa) === placaN && !c.cancelado && !c.finalizado).sort(ordemCiclo);
  for (let i = 0; i < lista.length; i++) {
    const n = base + i + 1;
    if (lista[i].cicloNum !== n) {
      await fb().setDoc(refCiclo(lista[i].id), { cicloNum: n }, { merge: true });
      lista[i].cicloNum = n;
      if (st.ciclos.has(lista[i].id)) st.ciclos.get(lista[i].id).cicloNum = n;
    }
  }
}

// Marcos automáticos vindos da oferta no Painel (availability do dia).
async function marcosDaOferta(carrier, placa, ds) {
  if (!carrier || !placa || !ds) return {};
  try {
    const s = await fb().getDoc(fb().doc(fb().db, 'availability', `${carrier}__${placa}__${ds}`));
    if (!s.exists()) return {};
    const r = s.data();
    const m = {};
    if (r.oferta && r.ofertadoEm) m.oferta = r.ofertadoEm;
    if (r.status === 'disponivel') m.dispCaminhao = `${ds}T${r.time && /^\d{2}:\d{2}$/.test(r.time) ? r.time : '00:00'}`;
    return m;
  } catch (e) { return {}; }
}

// ═══════════════════════════════════════════════════════════════════════════
// GANCHOS DO ROTEIRIZADOR
// ═══════════════════════════════════════════════════════════════════════════
// viagens: [{petId, veiculoId, placa, transportadora, base, dataCarregamento,
//            horaCarregamento, entregas:[{cliente, pedidoId, itens:[{produto,volume}]}]}]
async function sobCiclosSincronizarArquivo(viagens, { arquivo, revisaoDe } = {}) {
  if (!ehOperacional() || !window.fbDb) return;
  const allP = await window.dbGetPlates();
  const sob = (viagens || []).map(t => {
    const m = window._casarPlacaPainel?.(allP, t.transportadora, t.placa);
    return (m && window.ehSobDemanda?.(m.p)) ? { ...t, carrier: m.carrier, placa: m.p.placa } : null;
  }).filter(Boolean);
  if (!sob.length && !revisaoDe) return;
  const ativos = await carregarAtivos();
  await carregarCfg();
  const arquivosLig = new Set([arquivo, revisaoDe].filter(Boolean));
  const petsAgora = new Set(sob.map(t => String(t.petId)));
  const placasMexidas = new Set();
  const msgs = [];
  for (const t of sob) {
    const dados = {
      placa: t.placa, carrier: t.carrier, base: t.base || '',
      dataCarregamento: t.dataCarregamento || '', horaCarregamento: t.horaCarregamento || '',
      entregas: (t.entregas || []).map(e => ({ cliente: e.cliente || '', pedidoId: e.pedidoId ?? '', cidade: e.cidade || '', itens: (e.itens || []).map(it => ({ produto: it.produto || '', volume: Number(it.volume) || 0 })) })),
      viagem: { arquivo: arquivo || '', petId: String(t.petId), veiculoId: t.veiculoId ?? '' },
      previa: false, cancelado: false,
    };
    let existente = ativos.find(c => String(c.viagem?.petId || '') === String(t.petId));
    if (!existente) {
      // Ciclo dessa viagem já finalizado (programação antiga regravada)?
      const { col, db, getDocs, query, where } = fb();
      const q = await getDocs(query(col(db, 'sob_ciclos'), where('viagem.petId', '==', String(t.petId))));
      q.forEach(d => { if (!existente) existente = { id: d.id, ...d.data() }; });
      if (existente) st.ciclos.set(existente.id, existente);
    }
    if (existente) {
      const mudouChave = normPlaca(existente.placa) !== normPlaca(t.placa) || existente.dataCarregamento !== dados.dataCarregamento || existente.horaCarregamento !== dados.horaCarregamento;
      if (existente.finalizado && !existente.cancelado) { await gravarCiclo(existente.id, { viagem: dados.viagem, entregas: dados.entregas }, null); continue; }
      if (mudouChave) { placasMexidas.add(normPlaca(existente.placa)); placasMexidas.add(normPlaca(t.placa)); }
      if (existente.cancelado) { dados.finalizado = false; dados.status = existente.status === 'finalizado' ? 'programado' : (existente.status || 'programado'); placasMexidas.add(normPlaca(t.placa)); }
      await gravarCiclo(existente.id, dados, existente.cancelado ? 'Reativado pelo roteirizador' : null);
      continue;
    }
    // Prévia da mesma placa com carregamento próximo (±3 dias) → vira esta viagem
    const dia = d => d ? new Date(d + 'T12:00:00').getTime() : NaN;
    const previa = ativos.filter(c => c.previa && !c.cancelado && normPlaca(c.placa) === normPlaca(t.placa))
      .map(c => ({ c, dif: Math.abs((dia(c.dataCarregamento) - dia(dados.dataCarregamento)) / 86400000) }))
      .filter(x => !(x.dif > 3)).sort((a, b) => (a.dif || 0) - (b.dif || 0))[0]?.c;
    const marcos = await marcosDaOferta(t.carrier, t.placa, dados.dataCarregamento);
    placasMexidas.add(normPlaca(t.placa));
    if (previa) {
      const m = { ...marcos, ...(previa.marcos || {}) };
      const c = await gravarCiclo(previa.id, { ...dados, origem: 'roteirizador', marcos: m }, `Prévia substituída pela viagem ${t.petId} do roteirizador`);
      Object.assign(previa, c);
      msgs.push(`prévia de ${t.placa} vinculada à viagem ${t.petId}`);
    } else {
      const id = `rot__${String(t.petId).replace(/[\/\s]/g, '_')}`;
      const c = await gravarCiclo(id, { ...dados, origem: 'roteirizador', status: 'programado', finalizado: false, marcos, criadoEm: new Date().toISOString(), criadoPor: usuario() }, `Criado pela viagem ${t.petId} do roteirizador`);
      ativos.push(c);
      msgs.push(`${t.placa} (${t.petId})`);
    }
  }
  // Viagens que saíram da programação (excluídas, ou trocadas para placa diária)
  for (const c of ativos) {
    if (c.cancelado || c.finalizado || !c.viagem?.petId || !arquivosLig.has(c.viagem.arquivo)) continue;
    if (petsAgora.has(String(c.viagem.petId))) continue;
    await gravarCiclo(c.id, { cancelado: true, finalizado: true }, `Cancelado: viagem ${c.viagem.petId} saiu da programação`);
    placasMexidas.add(normPlaca(c.placa));
    msgs.push(`${c.placa} cancelado (viagem saiu da programação)`);
  }
  for (const pn of placasMexidas) await renumerarPlaca(pn, ativos.filter(c => !c.cancelado));
  st.carregadoEm = 0; // tela relê na próxima abertura
  if (msgs.length) toast(`🚚 Operação sob demanda: ${msgs.join(', ')}`, true);
}
async function sobCiclosArquivoExcluido(arquivo) {
  if (!arquivo || !ehOperacional() || !window.fbDb) return;
  const ativos = await carregarAtivos();
  const placas = new Set();
  for (const c of ativos) {
    if (c.cancelado || c.viagem?.arquivo !== arquivo) continue;
    await gravarCiclo(c.id, { cancelado: true, finalizado: true }, 'Cancelado: programação excluída do Histórico');
    placas.add(normPlaca(c.placa));
  }
  for (const pn of placas) await renumerarPlaca(pn, ativos.filter(c => !c.cancelado));
  st.carregadoEm = 0;
}
async function sobCicloMarcarEnvioPedido(petId) {
  if (!petId || !window.fbDb) return;
  try {
    const { col, db, getDocs, query, where } = fb();
    const q = await getDocs(query(col(db, 'sob_ciclos'), where('viagem.petId', '==', String(petId))));
    for (const d of q.docs) {
      const c = { id: d.id, ...d.data() }; st.ciclos.set(c.id, c);
      if (!c.marcos?.envioPedido) await gravarCiclo(c.id, { marcos: { envioPedido: new Date().toISOString() } }, 'Envio do pedido (OR pelo WhatsApp)');
    }
  } catch (e) { console.warn('[sob ciclos] envio pedido', e); }
}
window.sobCiclosSincronizarArquivo = sobCiclosSincronizarArquivo;
window.sobCiclosArquivoExcluido = sobCiclosArquivoExcluido;
window.sobCicloMarcarEnvioPedido = sobCicloMarcarEnvioPedido;

// ═══════════════════════════════════════════════════════════════════════════
// MONTAGEM DAS SEÇÕES (tela e reporte usam a mesma regra)
// ═══════════════════════════════════════════════════════════════════════════
// Em curso   = 1º ciclo ativo de cada placa;
// Previstas  = demais ciclos ativos + 1 linha "próximo" para a placa sem
//              ciclo seguinte (Sem pedido / DSR / Manutenção);
// Pedidos para alocar = prévias sem placa;
// Finalizadas = finalizados no período (mais recentes primeiro).
function montarSecoes() {
  const filtro = semAcento(st.filtro).toUpperCase().trim();
  const passa = (carrier, placa, c) => {
    if (st.transp && carrier !== st.transp) return false;
    if (!filtro) return true;
    const alvo = semAcento(`${carrier} ${placa} ${c ? clientesTexto(c) : ''}`).toUpperCase();
    return alvo.includes(filtro);
  };
  const todos = [...st.ciclos.values()].filter(c => !c.cancelado);
  const ativos = todos.filter(c => !c.finalizado);
  const emCurso = [], previstas = [], alocar = [];
  const placasCad = new Map(st.placas.map(p => [normPlaca(p.placa), p]));
  const porPlaca = new Map();
  ativos.forEach(c => {
    if (!c.placa) { alocar.push(c); return; }
    const k = normPlaca(c.placa);
    if (!porPlaca.has(k)) porPlaca.set(k, []);
    porPlaca.get(k).push(c);
  });
  const chaves = new Set([...placasCad.keys(), ...porPlaca.keys()]);
  [...chaves].forEach(k => {
    const lista = (porPlaca.get(k) || []).sort(ordemCiclo);
    const cad = placasCad.get(k);
    const carrier = lista[0]?.carrier || cad?.carrier || '';
    const placa = lista[0]?.placa || cad?.placa || '';
    if (!passa(carrier, placa, null) && !lista.some(c => passa(carrier, placa, c))) return;
    if (lista[0]) emCurso.push({ tipo: 'ciclo', c: lista[0] });
    lista.slice(1).forEach(c => previstas.push({ tipo: 'ciclo', c }));
    if (cad) {
      const ult = lista[lista.length - 1];
      const v = st.cfg?.veiculos?.[k] || {};
      const base = Number(st.cfg?.base?.[k]) || 0;
      const estadoPadrao = ult?.proxCargaNota === 'MANUTENÇÃO' ? 'manutencao' : (ult?.proxCargaNota === 'DSR' ? 'dsr' : 'sem_pedido');
      previstas.push({ tipo: 'veiculo', k, carrier: cad.carrier, placa: cad.placa, capacidade: cad.capacidade, bases: cad.bases,
        num: (ult?.cicloNum || base) + 1,
        estado: v.estado || estadoPadrao,
        disponivelEm: v.disponivelEm || ult?.proxCarga || '', hora: v.hora || ult?.proxCargaHora || '', obs: v.obs || '' });
    }
  });
  const chaveOrd = r => r.tipo === 'ciclo' ? `${r.c.dataCarregamento || '9999'}${r.c.horaCarregamento || '99'}` : `${r.disponivelEm || '9999'}${r.hora || '99'}`;
  emCurso.sort((a, b) => chaveOrd(a).localeCompare(chaveOrd(b)));
  previstas.sort((a, b) => chaveOrd(a).localeCompare(chaveOrd(b)) || (a.tipo === 'ciclo' ? -1 : 1));
  const de = isoHoje(-st.periodo);
  const finalizadas = todos.filter(c => c.finalizado && (c.dataCarregamento || '') >= de && passa(c.carrier, c.placa, c))
    .sort((a, b) => ordemCiclo(b, a)).map(c => ({ tipo: 'ciclo', c }));
  return {
    emCurso, previstas,
    alocar: alocar.filter(c => passa(c.carrier || '', '', c)).sort(ordemCiclo).map(c => ({ tipo: 'ciclo', c })),
    finalizadas,
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// TELA
// ═══════════════════════════════════════════════════════════════════════════
const inpCss = 'background:var(--dark);border:1px solid var(--border2);border-radius:6px;color:var(--text);padding:3px 5px;font-size:11.5px';
const chip = (o, extra = '') => `<span style="display:inline-block;background:${o.bg};color:${o.fg};font-weight:700;font-size:10.5px;border-radius:4px;padding:3px 7px;white-space:nowrap;${extra}">${esc(o.label)}</span>`;

function celMarco(c, m) {
  const v = c.marcos?.[m.id];
  if (v) return `<td style="font-size:11px;white-space:nowrap" title="${esc(m.label)}">${esc(brDataHora(v))}</td>`;
  return `<td style="white-space:nowrap"><button class="btn btn-sm" style="font-size:10px;padding:2px 6px" onclick="sobMarcarAgora('${jsq(c.id)}','${m.id}')" title="Registrar ${esc(m.label)} agora">✓ agora</button></td>`;
}

function linhaCiclo(c, secao) {
  const s = stById[c.status] || stById.programado;
  const prod = produtoCiclo(c);
  const tags = [
    c.previa ? `<span title="Lançada à mão (pedidos do SAP). Vira a viagem real quando for salva no roteirizador." style="font-size:9.5px;font-weight:700;color:#A78BFA;border:1px solid rgba(167,139,250,.5);border-radius:4px;padding:0 4px;margin-left:4px">PRÉVIA</span>` : '',
    c.origem === 'roteirizador' && c.viagem?.petId ? `<span title="Viagem ${esc(c.viagem.petId)} do roteirizador" style="font-size:9.5px;color:var(--muted);margin-left:4px">${esc(c.viagem.petId)}</span>` : '',
  ].join('');
  const statusSel = `<select onchange="sobMudarStatus('${jsq(c.id)}',this.value)" style="background:${s.bg};color:${s.fg};font-weight:700;font-size:10.5px;border:none;border-radius:4px;padding:4px 4px;cursor:pointer;max-width:150px">
      ${STATUS.map(o => `<option value="${o.id}"${o.id === c.status ? ' selected' : ''} style="background:#fff;color:#111">${esc(o.label)}</option>`).join('')}</select>`;
  const eta = `<div style="display:flex;gap:3px"><input type="date" value="${esc(c.etaData || '')}" onchange="sobCampo('${jsq(c.id)}','etaData',this.value)" style="${inpCss};width:118px">
      <select onchange="sobCampo('${jsq(c.id)}','etaTurno',this.value)" style="${inpCss}"><option value="">turno</option>${TURNOS.map(t => `<option${t === c.etaTurno ? ' selected' : ''}>${t}</option>`).join('')}</select></div>`;
  const dsr = `<input type="date" value="${esc(c.dsr || '')}" onchange="sobCampo('${jsq(c.id)}','dsr',this.value)" style="${inpCss};width:100%;box-sizing:border-box">`;
  // Próx. carga: OU a data OU uma nota (Manutenção / DSR / Sem previsão) —
  // nunca as duas lado a lado (não cabe na coluna).
  const optsNota = sel => NOTAS_PROX.map(n => `<option value="${esc(n)}"${n === sel ? ' selected' : ''} style="background:#fff;color:#111">${n || '— usar data —'}</option>`).join('');
  const prox = c.proxCargaNota
    ? `<select onchange="sobCampo('${jsq(c.id)}','proxCargaNota',this.value)" title="Escolha '— usar data —' para voltar a informar a data" style="${inpCss};width:100%;box-sizing:border-box;background:#F4B183;color:#6B2A06;font-weight:700;border:none">${optsNota(c.proxCargaNota)}</select>`
    : `<div style="display:flex;gap:3px;align-items:center;min-width:0"><input type="date" value="${esc(c.proxCarga || '')}" onchange="sobCampo('${jsq(c.id)}','proxCarga',this.value)" style="${inpCss};flex:1;min-width:0;font-weight:700">
      <select onchange="sobCampo('${jsq(c.id)}','proxCargaNota',this.value)" title="Em vez de data: Manutenção, DSR, Sem previsão" style="${inpCss};width:38px;flex:none">${optsNota('')}</select></div>`;
  const ofertarBtn = (c.proxCarga && c.placa) ? `<button class="btn btn-sm" style="font-size:10px;padding:2px 6px" onclick="sobOfertar('${jsq(c.carrier)}','${jsq(c.placa)}','${c.proxCarga}','${jsq(c.proxCargaHora || '')}','Retorno do ciclo ${c.cicloNum || ''}')" title="Ofertar a placa no Painel para ${br(c.proxCarga)}">📅 Ofertar ${brCurto(c.proxCarga)}</button>` : '';
  const somenteLeitura = secao === 'finalizadas';
  return `<tr style="${c.previa ? 'background:rgba(167,139,250,.05)' : ''}">
    <td style="font-weight:700;white-space:nowrap">${c.cicloNum ? c.cicloNum + 'º' : '-'}</td>
    <td style="font-size:12px">${esc(c.carrier || '-')}</td>
    <td style="font-family:'DM Mono',monospace;font-weight:600;white-space:nowrap">${esc(c.placa || '-')}</td>
    <td style="font-size:12px;white-space:normal;word-break:break-word">${esc(clientesTexto(c) || '—')}${tags}</td>
    <td style="white-space:nowrap;font-size:12px">${esc(carregTexto(c))}${c.base ? `<div style="font-size:10px;color:var(--muted)">${esc(c.base)}</div>` : ''}</td>
    <td>${somenteLeitura ? `<span style="font-size:12px">${c.status === 'finalizado' && c.etaData ? 'FINALIZADA ' + br(c.etaData) : esc(etaTexto(c))}</span>` : eta}</td>
    <td>${somenteLeitura ? chip(s) : statusSel}</td>
    <td style="text-align:right;font-weight:600">${fmtVol(volCiclo(c))}</td>
    <td style="text-align:center;font-size:12px">${esc(prod)}</td>
    ${somenteLeitura ? '' : `<td style="overflow:hidden">${dsr}</td><td style="overflow:hidden">${prox}${ofertarBtn ? `<div style="margin-top:3px">${ofertarBtn}</div>` : ''}</td>`}
    ${st.marcos ? MARCOS.map(m => celMarco(c, m)).join('') + `<td style="font-size:11px;white-space:nowrap">${esc(c.dataCarregamento ? br(c.dataCarregamento) : '')}</td>` : ''}
    <td style="white-space:nowrap"><button class="btn btn-sm" style="font-size:11px;padding:3px 7px" onclick="sobAbrirCiclo('${jsq(c.id)}')" title="Detalhes, marcos, histórico de alterações">✏️</button></td>
  </tr>`;
}
function linhaVeiculo(r) {
  const e = estById[r.estado] || estById.sem_pedido;
  const k = jsq(r.k);
  return `<tr style="opacity:.92">
    <td style="font-weight:700;color:var(--muted)">${r.num}º</td>
    <td style="font-size:12px">${esc(r.carrier)}</td>
    <td style="font-family:'DM Mono',monospace;font-weight:600">${esc(r.placa)}</td>
    <td style="font-size:11.5px;color:var(--muted)"><input type="text" value="${esc(r.obs)}" placeholder="obs. (opcional)" onchange="sobVeiculo('${k}','obs',this.value)" style="${inpCss};width:100%;box-sizing:border-box"></td>
    <td style="white-space:nowrap"><div style="display:flex;gap:3px"><input type="date" value="${esc(r.disponivelEm)}" onchange="sobVeiculo('${k}','disponivelEm',this.value)" style="${inpCss};width:118px" title="Disponível para carregar a partir de"><input type="time" value="${esc(r.hora)}" onchange="sobVeiculo('${k}','hora',this.value)" style="${inpCss};width:78px"></div></td>
    <td style="color:var(--muted)">-</td>
    <td><select onchange="sobVeiculo('${k}','estado',this.value)" style="background:${e.bg};color:${e.fg};font-weight:700;font-size:10.5px;border:none;border-radius:4px;padding:4px;cursor:pointer">${ESTADOS.map(o => `<option value="${o.id}"${o.id === r.estado ? ' selected' : ''} style="background:#fff;color:#111">${esc(o.label)}</option>`).join('')}</select></td>
    <td style="text-align:right;color:var(--muted)">${esc(r.capacidade || '')}</td>
    <td></td><td></td>
    <td style="white-space:nowrap">${r.disponivelEm ? `<button class="btn btn-sm" style="font-size:10px;padding:2px 6px" onclick="sobOfertar('${jsq(r.carrier)}','${jsq(r.placa)}','${r.disponivelEm}','${jsq(r.hora)}','${jsq(r.obs || 'Próximo ciclo')}')" title="Ofertar no Painel para ${br(r.disponivelEm)}">📅 Ofertar ${brCurto(r.disponivelEm)}</button>` : ''}</td>
    ${st.marcos ? '<td></td>'.repeat(MARCOS.length + 1) : ''}
    <td><button class="btn btn-sm" style="font-size:11px;padding:3px 7px" onclick="sobNovaPrevia('${jsq(r.carrier)}','${jsq(r.placa)}','${r.disponivelEm}','${jsq(r.hora)}')" title="Lançar a prévia do próximo ciclo desta placa (pedidos do SAP)">➕ Prévia</button></td>
  </tr>`;
}
function tabelaSecao(titulo, linhas, secao, extraHead = '') {
  const fin = secao === 'finalizadas';
  const cab = ['CICLO', 'TRANSPORTADORA', 'PLACA', 'CLIENTE', 'CARREGAMENTO', fin ? 'ETA / FINALIZADA' : 'ETA DESTINO', 'STATUS', 'VOLUME', 'PRODUTO']
    .concat(fin ? [] : ['DSR', 'PRÓX. CARGA'])
    .concat(st.marcos ? MARCOS.map(m => m.label.toUpperCase()).concat(['DATA CARREG.']) : [])
    .concat(['']);
  const corpo = linhas.length
    ? linhas.map(r => r.tipo === 'ciclo' ? linhaCiclo(r.c, secao) : linhaVeiculo(r)).join('')
    : `<tr><td colspan="${cab.length}" style="text-align:center;color:var(--muted);font-size:12px;padding:14px">Nada aqui.</td></tr>`;
  // Larguras por coluna (a tabela do sistema usa layout fixo). A coluna
  // CLIENTE fica sem largura: ocupa todo o espaço que sobrar, então a tabela
  // preenche o quadro inteiro; só rola na horizontal em tela bem estreita.
  const larg = { 'CICLO': 46, 'TRANSPORTADORA': 112, 'PLACA': 84, 'CARREGAMENTO': 108, 'ETA DESTINO': 204, 'ETA / FINALIZADA': 120,
    'STATUS': 160, 'VOLUME': 62, 'PRODUTO': 70, 'DSR': 126, 'PRÓX. CARGA': 186, 'DATA CARREG.': 96, '': 84 };
  const ws = cab.map(h => h === 'CLIENTE' ? null : (larg[h] || 112));
  const minimo = ws.reduce((a, b) => a + (b || 0), 0) + 230;
  return `<div class="carrier-block" style="margin-bottom:1.1rem">
    <div class="carrier-head" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">${titulo}<span class="ch-count">${linhas.length}</span>${extraHead}</div>
    <div class="tscroll"><table class="table" style="table-layout:fixed;width:100%;min-width:${minimo}px">
      <colgroup>${ws.map(w => w ? `<col style="width:${w}px">` : '<col>').join('')}</colgroup>
      <thead><tr>${cab.map(h => `<th style="white-space:nowrap">${h}</th>`).join('')}</tr></thead>
      <tbody>${corpo}</tbody></table></div></div>`;
}

async function renderSobDemanda(body) {
  _body = body || _body || document.getElementById('tab-body');
  if (!_body) return;
  // Esta aba usa a largura toda da tela (o sistema limita o conteúdo a 1440px).
  if (!document.getElementById('sob-estilo-largo')) {
    const stl = document.createElement('style'); stl.id = 'sob-estilo-largo';
    stl.textContent = '.main.sob-larga{max-width:none;padding-left:1.25rem;padding-right:1.25rem}';
    document.head.appendChild(stl);
  }
  document.querySelector('.main')?.classList.add('sob-larga');
  if (!ehOperacional()) { _body.innerHTML = '<div class="empty">Acesso restrito.</div>'; return; }
  if (!st.carregadoEm) _body.innerHTML = '<div class="loading"><span class="spin"></span>Carregando ciclos da operação sob demanda...</div>';
  try { await carregarTudo(false); }
  catch (e) {
    console.error('[sob demanda]', e);
    _body.innerHTML = `<div class="empty" style="flex-direction:column;gap:8px"><span class="ei">⚠</span><p>Não consegui ler os ciclos: ${esc(e.code || e.message || e)}</p><p style="font-size:12px;color:var(--muted)">Se for "permission-denied", publique as regras do Firestore atualizadas (bloco sob_ciclos).</p></div>`;
    return;
  }
  const sec = montarSecoes();
  const transps = [...new Set(st.placas.map(p => p.carrier).concat([...st.ciclos.values()].map(c => c.carrier).filter(Boolean)))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const emDesloc = sec.emCurso.filter(r => ['desloc_rj', 'descarga', 'desloc_sp', 'carregamento'].includes(r.c.status)).length;
  const volPer = sec.finalizadas.reduce((s, r) => s + volCiclo(r.c), 0);
  const kpi = (v, l) => `<div style="background:var(--mid);border:1px solid var(--border);border-radius:var(--radius);padding:8px 14px;min-width:120px"><div style="font-size:20px;font-weight:700;color:var(--lime)">${v}</div><div style="font-size:11px;color:var(--muted)">${l}</div></div>`;
  _body.innerHTML = `
    <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px">
      <p class="sec-title" style="margin:0">🚚 Operação sob demanda — ciclos</p>
      <div style="margin-left:auto;display:flex;gap:6px;flex-wrap:wrap">
        <button class="btn btn-lime btn-sm" onclick="sobNovaPrevia()">➕ Prévia de ciclo</button>
        <button class="btn btn-sm" onclick="sobExportarExcel()">📊 Excel</button>
        <button class="btn btn-sm" onclick="sobExportarPDF()">📄 PDF</button>
        <button class="btn btn-sm" onclick="sobAbrirConfig()">⚙ Siglas, apelidos e contagem</button>
        <button class="btn btn-sm" onclick="sobRecarregar()" title="Reler do banco">🔄</button>
      </div>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px">
      ${kpi(sec.emCurso.length, 'ciclos em curso')}${kpi(emDesloc, 'carregando / em trânsito / descarga')}
      ${kpi(sec.previstas.filter(r => r.tipo === 'ciclo').length, 'próximos ciclos já previstos')}
      ${kpi(sec.previstas.filter(r => r.tipo === 'veiculo' && r.estado === 'sem_pedido').length, 'caminhões sem pedido')}
      ${kpi(`${sec.finalizadas.length} · ${fmtVol(volPer)} m³`, `finalizados nos últimos ${st.periodo} dias`)}
    </div>
    <div class="filters-bar">
      <select onchange="sobFiltro('transp',this.value)"><option value="">Todas as transportadoras</option>${transps.map(t => `<option${t === st.transp ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>
      <input type="text" id="sob-busca" placeholder="🔍 Placa, transportadora ou cliente..." value="${esc(st.filtro)}" oninput="sobFiltroTexto(this.value)" style="background:var(--dark);border:1px solid var(--border2);border-radius:var(--radius);color:var(--text);font-size:13px;padding:8px 12px;outline:none;">
      <select onchange="sobFiltro('periodo',+this.value)" title="Período das finalizadas">${[7, 15, 30, 60, 90].map(d => `<option value="${d}"${d === st.periodo ? ' selected' : ''}>Finalizadas: ${d} dias</option>`).join('')}</select>
      <label style="font-size:12px;color:var(--muted);display:flex;align-items:center;gap:6px;cursor:pointer"><input type="checkbox" ${st.marcos ? 'checked' : ''} onchange="sobFiltro('marcos',this.checked)"> Mostrar marcos (oferta, liberações, envio)</label>
    </div>
    ${tabelaSecao('🛣 Entregas em curso', sec.emCurso, 'emCurso')}
    ${tabelaSecao('🗓 Entregas previstas', sec.previstas, 'previstas', '<span style="font-size:11px;color:var(--muted);font-weight:400">Próximos ciclos já lançados + o próximo de cada caminhão (sem pedido / DSR / manutenção).</span>')}
    ${tabelaSecao('📦 Pedidos para alocar', sec.alocar, 'alocar', '<span style="font-size:11px;color:var(--muted);font-weight:400">Prévias ainda sem caminhão definido.</span>')}
    <div style="margin:4px 0 8px"><button class="btn btn-sm" onclick="sobFiltro('finalAberto',${!st.finalAberto})">${st.finalAberto ? '▼' : '▶'} Entregas finalizadas (${sec.finalizadas.length} nos últimos ${st.periodo} dias)</button></div>
    ${st.finalAberto ? tabelaSecao('✅ Entregas finalizadas', sec.finalizadas, 'finalizadas') : ''}
  `;
}
let _tBusca = null;
window.sobFiltroTexto = v => { st.filtro = v; clearTimeout(_tBusca); _tBusca = setTimeout(async () => { await renderSobDemanda(); const el = document.getElementById('sob-busca'); if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); } }, 350); };
window.sobFiltro = (k, v) => { st[k] = v; renderSobDemanda(); };
window.sobRecarregar = async () => { st.carregadoEm = 0; await renderSobDemanda(); toast('Ciclos atualizados.', true); };

// ── Ações rápidas na tabela ─────────────────────────────────────────────────
window.sobMudarStatus = async (id, status) => {
  const c = st.ciclos.get(id); if (!c) return;
  const antes = stById[c.status]?.label || c.status;
  const patch = { status };
  await carregarCfg();
  if (status === 'finalizado') {
    patch.finalizado = true; patch.finalizadoEm = new Date().toISOString();
    if (!c.etaData) patch.etaData = isoHoje(0);
  } else if (c.finalizado) patch.finalizado = false;
  try {
    await gravarCiclo(id, patch, `Status: ${antes} → ${stById[status]?.label || status}`);
    const pn = normPlaca(c.placa);
    if (status === 'finalizado' && pn && c.cicloNum) {
      await carregarCfg();
      if ((Number(st.cfg.base?.[pn]) || 0) < c.cicloNum) await salvarCfg({ base: { [pn]: c.cicloNum } });
      if (st.cfg.veiculos?.[pn]) await salvarCfg({ veiculos: { [pn]: { estado: 'sem_pedido', disponivelEm: '', hora: '', obs: '' } } });
      toast(`Ciclo ${c.cicloNum}º de ${c.placa} finalizado.${c.proxCarga ? '' : ' Preencha a próxima carga em Previstas.'}`, true);
    } else if (c.finalizado && status !== 'finalizado' && pn) {
      if ((Number(st.cfg.base?.[pn]) || 0) === c.cicloNum) await salvarCfg({ base: { [pn]: c.cicloNum - 1 } });
    }
  } catch (e) { toast('Não consegui salvar: ' + (e.code || e.message), false); }
  renderSobDemanda();
};
window.sobCampo = async (id, campo, valor) => {
  const c = st.ciclos.get(id); if (!c) return;
  const nomes = { etaData: 'ETA', etaTurno: 'Turno ETA', dsr: 'DSR', proxCarga: 'Próx. carga', proxCargaNota: 'Próx. carga (nota)', proxCargaHora: 'Hora próx. carga' };
  try { await gravarCiclo(id, { [campo]: valor }, `${nomes[campo] || campo}: ${valor || '(vazio)'}`); }
  catch (e) { toast('Não consegui salvar: ' + (e.code || e.message), false); }
  if (campo === 'proxCarga' || campo === 'proxCargaNota') renderSobDemanda();
};
window.sobVeiculo = async (k, campo, valor) => {
  await carregarCfg();
  const atual = { estado: 'sem_pedido', disponivelEm: '', hora: '', obs: '', ...(st.cfg.veiculos?.[k] || {}) };
  atual[campo] = valor;
  try { await salvarCfg({ veiculos: { [k]: atual } }); } catch (e) { toast('Não consegui salvar: ' + (e.code || e.message), false); }
  renderSobDemanda();
};
window.sobMarcarAgora = async (id, marco) => {
  const m = MARCOS.find(x => x.id === marco);
  try { await gravarCiclo(id, { marcos: { [marco]: new Date().toISOString() } }, `${m?.label || marco}: agora`); }
  catch (e) { toast('Não consegui salvar: ' + (e.code || e.message), false); }
  renderSobDemanda();
};
// Oferta a placa no Painel de Disponibilidade (mesmo registro da tela de ofertas).
window.sobOfertar = async (carrier, placa, ds, hora, obs) => {
  if (!carrier || !placa || !ds) return;
  const id = `${carrier}__${placa}__${ds}`;
  try {
    const ref = fb().doc(fb().db, 'availability', id);
    const ex = await fb().getDoc(ref);
    if (ex.exists() && ex.data().status) {
      const d = ex.data();
      if (!confirm(`${placa} já tem registro para ${br(ds)}: ${d.oferta ? 'ofertada' : d.status}${d.time ? ' a partir das ' + d.time : ''}.\n\nSubstituir pela oferta${hora ? ' a partir das ' + hora : ''}?`)) return;
    } else if (!confirm(`Ofertar ${placa} (${carrier}) no Painel para ${br(ds)}${hora ? ' a partir das ' + hora : ''}?\n\nO Roteirizador passa a poder usar a placa nesse dia.`)) return;
    const agora = new Date().toISOString();
    await fb().setDoc(ref, { carrier, plate: placa, dateStr: ds, status: 'disponivel', time: hora || '', filled: true,
      oferta: true, baseOferta: '', obsOferta: obs || '', ofertadoPor: usuario(), ofertadoEm: agora, updatedAt: agora,
      motoristaDiurno: ex.exists() ? (ex.data().motoristaDiurno || '') : '', motoristaNoturno: ex.exists() ? (ex.data().motoristaNoturno || '') : '',
      hodometro: null, hodometroFotoUrl: null });
    window._invalidarStatusPlaca?.(carrier, placa, ds);
    try { await window.dbAddAudit?.('oferta_sob_demanda', { carrier, plate: placa, date: ds }); } catch (e) {}
    toast(`${placa} ofertada para ${br(ds)}${hora ? ' a partir das ' + hora : ''}.`, true);
  } catch (e) { toast('Não consegui ofertar: ' + (e.code || e.message), false); }
};

// ═══════════════════════════════════════════════════════════════════════════
// MODAL — detalhe do ciclo / nova prévia
// ═══════════════════════════════════════════════════════════════════════════
let _mod = null; // { id|null, entregas:[{cliente, produto, volume}] (linhas planas) }

function fecharModal() { document.getElementById('sob-modal')?.remove(); _mod = null; }
window.sobFecharModal = fecharModal;
function abrirModal(html, largura = 860) {
  document.getElementById('sob-modal')?.remove();
  const m = document.createElement('div');
  m.id = 'sob-modal';
  m.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.7);display:flex;align-items:center;justify-content:center;z-index:9999;padding:1rem';
  m.innerHTML = `<div style="background:var(--mid);border:1px solid var(--border2);border-radius:var(--radius-lg);padding:1.5rem;width:100%;max-width:${largura}px;max-height:92vh;overflow-y:auto">${html}</div>`;
  m.addEventListener('mousedown', e => { if (e.target === m) fecharModal(); });
  document.body.appendChild(m);
}
const lab = (t, campo) => `<label style="display:flex;flex-direction:column;gap:4px;font-size:11px;color:var(--muted);font-weight:600">${t}${campo}</label>`;
const inpM = 'background:var(--dark);border:1px solid var(--border2);border-radius:6px;color:var(--text);padding:6px 8px;font-size:12.5px';

// Achata entregas → linhas editáveis e volta.
const planificar = ents => (ents || []).flatMap(e => (e.itens && e.itens.length ? e.itens : [{ produto: '', volume: 0 }]).map(it => ({ cliente: e.cliente || '', pedidoId: e.pedidoId ?? '', produto: it.produto || '', volume: Number(it.volume) || 0 })));
function agrupar(linhas) {
  const map = new Map();
  linhas.filter(l => (l.cliente || '').trim() && (Number(l.volume) || 0) > 0).forEach(l => {
    const k = l.cliente.trim().toUpperCase();
    if (!map.has(k)) map.set(k, { cliente: l.cliente.trim(), pedidoId: l.pedidoId || '', itens: [] });
    const e = map.get(k);
    const ex = e.itens.find(i => i.produto === (l.produto || '').trim());
    if (ex) ex.volume += Number(l.volume) || 0; else e.itens.push({ produto: (l.produto || '').trim(), volume: Number(l.volume) || 0 });
  });
  return [...map.values()];
}

// Colar do SAP/Excel: cada linha vira cliente + produto + volume (litros → m³).
function parseColagem(txt) {
  const linhas = [];
  String(txt || '').split(/\r?\n/).forEach(l => {
    if (!l.trim()) return;
    const cel = l.split(/\t|;|\s{2,}/).map(x => x.trim()).filter(Boolean);
    if (!cel.length) return;
    // Linha "solta" com espaço simples: "P.1 S10 10", "BIG JOHN 10 GC", "CUTUBA 10m³ S10"
    if (cel.length === 1) {
      const t = cel[0];
      const reProd = /\b(S[\s-]?10|S[\s-]?500|GC|GA|EA|ETANOL|GASOLINA(?:\s+C)?(?:\s+COMUM|\s+ADITIVADA)?|DIESEL(?:\s+S10|\s+S500)?)\b/i;
      const reNum = /(?:^|\s)(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:[.,]\d+)?)\s*(?:m³|m3|l|lts?)?(?=\s|$)/i;
      const mp = reProd.exec(t), mn = reNum.exec(t);
      if (mp && mn) {
        let vs = mn[1]; vs = /^\d{1,3}(\.\d{3})+/.test(vs) ? vs.replace(/\./g, '').replace(',', '.') : vs.replace(',', '.');
        let volume = parseFloat(vs) || 0; if (volume >= 500) volume /= 1000;
        const cliente = t.replace(mp[0], ' ').replace(mn[0], ' ').replace(/\s+/g, ' ').replace(/[-–(]\s*$/, '').trim();
        if (cliente) { linhas.push({ cliente, pedidoId: '', produto: mp[1].toUpperCase().replace(/[\s-](?=\d)/g, ''), volume: Math.round(volume * 100) / 100 }); return; }
      }
    }
    const num = x => { let t = x.replace(/\s|m³|m3|L$/gi, ''); if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t = t.replace(/\./g, '').replace(',', '.'); else t = t.replace(',', '.'); return /^\d+(\.\d+)?$/.test(t) ? parseFloat(t) : NaN; };
    const ehProd = x => /\b(S[\s-]?10|S[\s-]?500|GASOLINA|ETANOL|DIESEL|ALCOOL|ÁLCOOL|EHC|GC|GA|EA)\b/i.test(x);
    let produto = cel.find(x => ehProd(x) && !/^\d+$/.test(x)) || '';
    let pedidoId = '';
    const nums = [];
    cel.forEach(x => { if (/^\d{7,}$/.test(x)) { if (!pedidoId) pedidoId = x; return; } const n = num(x); if (!isNaN(n) && n > 0) nums.push(n); });
    let volume = nums.length ? nums[nums.length - 1] : 0;
    if (volume >= 500) volume = volume / 1000; // veio em litros
    const textos = cel.filter(x => x !== produto && isNaN(num(x)) && !/^\d+$/.test(x) && /[A-Za-zÀ-ú]/.test(x) && !/^\d{2}\/\d{2}/.test(x));
    // "TREMENDÃO 10 S10" numa célula só
    if (!produto) { const m = /(.*?)[\s-]+(S[\s-]?10|S[\s-]?500|GC|GA|EA|ETANOL)\b/i.exec(cel.join(' ')); if (m) { produto = m[2].toUpperCase().replace(/[\s-]/g, ''); } }
    let cliente = textos.sort((a, b) => b.length - a.length)[0] || '';
    if (cliente && produto && cliente.toUpperCase().includes(produto.toUpperCase())) cliente = cliente.replace(new RegExp('\\s*\\d*[,.]?\\d*\\s*(m³|m3)?\\s*' + produto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '.*$', 'i'), '').trim();
    if (!cliente && !volume) return;
    linhas.push({ cliente, pedidoId, produto, volume: Math.round(volume * 100) / 100 });
  });
  return linhas;
}

function htmlLinhasEntregas(editavel) {
  const linhas = _mod.linhas;
  if (!editavel) {
    return `<table class="table" style="font-size:12px"><thead><tr><th>CLIENTE</th><th>PEDIDO</th><th>PRODUTO</th><th>SIGLA</th><th style="text-align:right">m³</th></tr></thead><tbody>
      ${linhas.map(l => `<tr><td>${esc(l.cliente)}</td><td>${esc(l.pedidoId)}</td><td>${esc(l.produto)}</td><td>${esc(sigla(l.produto))}</td><td style="text-align:right">${fmtVol(l.volume)}</td></tr>`).join('')}</tbody></table>`;
  }
  return `<table class="table" style="font-size:12px"><thead><tr><th>CLIENTE</th><th>PRODUTO</th><th style="width:60px">SIGLA</th><th style="width:90px;text-align:right">m³</th><th style="width:30px"></th></tr></thead><tbody>
    ${linhas.map((l, i) => `<tr>
      <td><input value="${esc(l.cliente)}" oninput="sobLinha(${i},'cliente',this.value)" style="${inpM};width:100%;box-sizing:border-box"></td>
      <td><input value="${esc(l.produto)}" oninput="sobLinha(${i},'produto',this.value)" placeholder="ex.: S10, Gasolina C" style="${inpM};width:100%;box-sizing:border-box"></td>
      <td style="font-size:11px;font-weight:700">${esc(sigla(l.produto))}</td>
      <td><input type="number" step="0.5" min="0" value="${l.volume || ''}" oninput="sobLinha(${i},'volume',this.value)" style="${inpM};width:80px;text-align:right"></td>
      <td><button class="btn btn-sm" style="padding:2px 6px" onclick="sobLinhaRemover(${i})">✕</button></td></tr>`).join('')}
    </tbody></table>
    <div style="display:flex;gap:8px;align-items:center;margin-top:6px"><button class="btn btn-sm" onclick="sobLinhaAdd()">+ linha</button>
      <span style="font-size:12px;color:var(--muted)">Total: <b id="sob-mod-total" style="color:var(--text)">${fmtVol(linhas.reduce((s, l) => s + (Number(l.volume) || 0), 0))} m³</b></span></div>`;
}
function redesenharLinhas() { const box = document.getElementById('sob-mod-linhas'); if (box) box.innerHTML = htmlLinhasEntregas(_mod.editavel); }
window.sobLinha = (i, campo, v) => {
  if (!_mod?.linhas[i]) return;
  _mod.linhas[i][campo] = campo === 'volume' ? (parseFloat(String(v).replace(',', '.')) || 0) : v;
  const t = document.getElementById('sob-mod-total'); if (t) t.textContent = `${fmtVol(_mod.linhas.reduce((s, l) => s + (Number(l.volume) || 0), 0))} m³`;
  if (campo === 'produto') { const tr = document.querySelectorAll('#sob-mod-linhas tbody tr')[i]; if (tr) tr.children[2].textContent = sigla(v); }
};
window.sobLinhaAdd = () => { _mod.linhas.push({ cliente: '', produto: '', volume: 0 }); redesenharLinhas(); };
window.sobLinhaRemover = i => { _mod.linhas.splice(i, 1); redesenharLinhas(); };
window.sobColar = () => {
  const ta = document.getElementById('sob-mod-colar');
  const novas = parseColagem(ta?.value);
  if (!novas.length) { toast('Não reconheci linhas no texto colado.', false); return; }
  _mod.linhas = _mod.linhas.filter(l => (l.cliente || '').trim() || l.volume).concat(novas);
  if (ta) ta.value = '';
  redesenharLinhas();
  toast(`${novas.length} linha(s) lidas — confira cliente, produto e m³.`, true);
};

function optsPlacas(sel) {
  return `<option value="">— Sem caminhão ainda (pedidos para alocar) —</option>` + st.placas.map(p => {
    const v = `${p.carrier}||${p.placa}`;
    return `<option value="${esc(v)}"${v === sel ? ' selected' : ''}>${esc(p.placa)} · ${esc(p.carrier)} · ${esc(p.capacidade || '')}m³</option>`;
  }).join('');
}

window.sobNovaPrevia = async (carrier = '', placa = '', data = '', hora = '') => {
  await carregarTudo(false);
  _mod = { id: null, editavel: true, linhas: [{ cliente: '', produto: '', volume: 0 }] };
  const sel = carrier && placa ? `${carrier}||${placa}` : '';
  abrirModal(`
    <p style="font-size:15px;font-weight:600;color:var(--lime);margin-bottom:4px">➕ Prévia de ciclo</p>
    <p style="font-size:12px;color:var(--muted);margin-bottom:14px">Para pedidos que vocês já veem no SAP mas ainda não subiram no Roteirizador. Aparece no reporte normalmente; quando a viagem for salva no Roteirizador (mesma placa, carregamento até 3 dias de diferença), a prévia é substituída pelos dados reais — mantendo nº do ciclo, status e marcos.</p>
    <div style="display:grid;grid-template-columns:2fr 1fr 1fr 1.3fr;gap:10px;margin-bottom:12px">
      ${lab('Caminhão', `<select id="sob-mod-placa" style="${inpM}">${optsPlacas(sel)}</select>`)}
      ${lab('Carregamento (data)', `<input type="date" id="sob-mod-data" value="${esc(data)}" style="${inpM}">`)}
      ${lab('Hora', `<input type="time" id="sob-mod-hora" value="${esc(hora)}" style="${inpM}">`)}
      ${lab('Base', `<input type="text" id="sob-mod-base" placeholder="ex.: São Caetano" style="${inpM}">`)}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:12px">
      ${lab('ETA destino', `<input type="date" id="sob-mod-eta" style="${inpM}">`)}
      ${lab('Turno', `<select id="sob-mod-turno" style="${inpM}"><option value=""></option>${TURNOS.map(t => `<option>${t}</option>`).join('')}</select>`)}
      ${lab('DSR', `<input type="date" id="sob-mod-dsr" style="${inpM}">`)}
    </div>
    <div style="border:1px dashed var(--border2);border-radius:8px;padding:10px;margin-bottom:10px">
      <div style="font-size:12px;font-weight:600;margin-bottom:6px">📋 Colar do SAP / Excel</div>
      <textarea id="sob-mod-colar" rows="4" placeholder="Copie as linhas do SAP ou da planilha (cliente, produto, volume em m³ ou litros) e cole aqui" style="${inpM};width:100%;box-sizing:border-box;font-family:'DM Mono',monospace;font-size:11.5px"></textarea>
      <button class="btn btn-sm" style="margin-top:6px" onclick="sobColar()">Ler linhas coladas</button>
    </div>
    <div id="sob-mod-linhas">${htmlLinhasEntregas(true)}</div>
    <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px">
      <button class="btn btn-sm" onclick="sobFecharModal()">Cancelar</button>
      <button class="btn btn-lime btn-sm" onclick="sobSalvarPrevia()">💾 Salvar prévia</button>
    </div>`);
  // base sugerida pela placa
  const ajustaBase = () => { const v = document.getElementById('sob-mod-placa').value; const p = st.placas.find(x => `${x.carrier}||${x.placa}` === v); const b = document.getElementById('sob-mod-base'); if (b && !b.value && p?.bases?.length === 1) b.value = p.bases[0]; };
  document.getElementById('sob-mod-placa').addEventListener('change', ajustaBase); ajustaBase();
};
window.sobSalvarPrevia = async () => {
  const entregas = agrupar(_mod.linhas);
  if (!entregas.length) { alert('Inclua pelo menos um cliente com volume.'); return; }
  const [carrier, placa] = (document.getElementById('sob-mod-placa').value || '').split('||');
  const ds = document.getElementById('sob-mod-data').value;
  const id = `prev__${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  const marcos = placa && ds ? await marcosDaOferta(carrier, placa, ds) : {};
  try {
    await gravarCiclo(id, {
      origem: 'previa', previa: true, cancelado: false, finalizado: false, status: 'programado',
      placa: placa || '', carrier: carrier || '', base: document.getElementById('sob-mod-base').value.trim(),
      dataCarregamento: ds, horaCarregamento: document.getElementById('sob-mod-hora').value,
      etaData: document.getElementById('sob-mod-eta').value, etaTurno: document.getElementById('sob-mod-turno').value,
      dsr: document.getElementById('sob-mod-dsr').value,
      entregas, marcos, criadoEm: new Date().toISOString(), criadoPor: usuario(),
    }, 'Prévia criada');
    if (placa) {
      await renumerarPlaca(normPlaca(placa));
      const pn = normPlaca(placa);
      if (st.cfg?.veiculos?.[pn]?.estado && st.cfg.veiculos[pn].estado !== 'sem_pedido') await salvarCfg({ veiculos: { [pn]: { ...st.cfg.veiculos[pn], estado: 'sem_pedido' } } });
    }
    fecharModal();
    toast('Prévia salva.', true);
    renderSobDemanda();
  } catch (e) { alert('Não consegui salvar: ' + (e.code || e.message)); }
};

window.sobAbrirCiclo = id => {
  const c = st.ciclos.get(id); if (!c) return;
  const editavel = !!c.previa;
  _mod = { id, editavel, linhas: planificar(c.entregas) };
  const s = stById[c.status] || stById.programado;
  const sel = c.placa ? `${c.carrier}||${c.placa}` : '';
  const placaCampo = editavel
    ? `<select id="sob-mod-placa" style="${inpM}">${optsPlacas(sel)}</select>`
    : `<div style="${inpM};opacity:.85">${esc(c.placa)} · ${esc(c.carrier)}</div>`;
  abrirModal(`
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-wrap:wrap">
      <p style="font-size:15px;font-weight:600;color:var(--lime);margin:0">${c.cicloNum ? c.cicloNum + 'º ciclo' : 'Ciclo'} · ${esc(c.placa || 'sem caminhão')}</p>
      ${chip(s)} ${c.previa ? '<span style="font-size:11px;font-weight:700;color:#A78BFA">PRÉVIA</span>' : ''}
      ${c.viagem?.petId ? `<span style="font-size:11px;color:var(--muted)">Viagem ${esc(c.viagem.petId)} · ${esc(c.viagem.arquivo || '')}</span>` : ''}
    </div>
    <div style="display:grid;grid-template-columns:2fr 1fr 1fr 1.3fr;gap:10px;margin-bottom:12px">
      ${lab('Caminhão', placaCampo)}
      ${lab('Carregamento (data)', `<input type="date" id="sob-mod-data" value="${esc(c.dataCarregamento || '')}" ${editavel ? '' : 'disabled'} style="${inpM}">`)}
      ${lab('Hora', `<input type="time" id="sob-mod-hora" value="${esc(c.horaCarregamento || '')}" ${editavel ? '' : 'disabled'} style="${inpM}">`)}
      ${lab('Base', `<input type="text" id="sob-mod-base" value="${esc(c.base || '')}" ${editavel ? '' : 'disabled'} style="${inpM}">`)}
    </div>
    ${editavel ? '' : `<p style="font-size:11px;color:var(--muted);margin:-4px 0 10px">Caminhão, carregamento e entregas vêm do Roteirizador — para mudar, edite a viagem no Histórico (✏️ volume, 🗑 entrega, 🔁 troca de veículo) e o ciclo atualiza sozinho.</p>`}
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr 0.8fr 1fr;gap:10px;margin-bottom:12px">
      ${lab('ETA destino', `<input type="date" id="sob-mod-eta" value="${esc(c.etaData || '')}" style="${inpM}">`)}
      ${lab('Turno', `<select id="sob-mod-turno" style="${inpM}"><option value=""></option>${TURNOS.map(t => `<option${t === c.etaTurno ? ' selected' : ''}>${t}</option>`).join('')}</select>`)}
      ${lab('DSR', `<input type="date" id="sob-mod-dsr" value="${esc(c.dsr || '')}" style="${inpM}">`)}
      ${lab('Próx. carga', `<input type="date" id="sob-mod-prox" value="${esc(c.proxCarga || '')}" style="${inpM}">`)}
      ${lab('Hora', `<input type="time" id="sob-mod-proxhora" value="${esc(c.proxCargaHora || '')}" style="${inpM}">`)}
      ${lab('Ou', `<select id="sob-mod-proxnota" style="${inpM}">${NOTAS_PROX.map(n => `<option value="${esc(n)}"${n === (c.proxCargaNota || '') ? ' selected' : ''}>${n || '— (usar a data)'}</option>`).join('')}</select>`)}
    </div>
    ${lab('Observação interna', `<input type="text" id="sob-mod-obs" value="${esc(c.obs || '')}" style="${inpM};margin-bottom:12px" placeholder="não sai no reporte">`)}
    <div style="font-size:12px;font-weight:600;margin:6px 0">Entregas</div>
    ${editavel ? `<div style="border:1px dashed var(--border2);border-radius:8px;padding:8px;margin-bottom:8px"><textarea id="sob-mod-colar" rows="2" placeholder="Colar mais linhas do SAP/Excel" style="${inpM};width:100%;box-sizing:border-box;font-family:'DM Mono',monospace;font-size:11.5px"></textarea><button class="btn btn-sm" style="margin-top:4px" onclick="sobColar()">Ler linhas coladas</button></div>` : ''}
    <div id="sob-mod-linhas">${htmlLinhasEntregas(editavel)}</div>
    <div style="font-size:12px;font-weight:600;margin:14px 0 6px">Marcos</div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px">
      ${MARCOS.map(m => lab(`${m.label}${m.auto ? ` <span style="font-weight:400" title="${esc(m.auto)}">ⓘ</span>` : ''}`, `<div style="display:flex;gap:4px"><input type="datetime-local" id="sob-mod-m-${m.id}" value="${esc(localDT(c.marcos?.[m.id]))}" style="${inpM};flex:1;min-width:0"><button class="btn btn-sm" style="padding:2px 6px" title="Agora" onclick="document.getElementById('sob-mod-m-${m.id}').value='${localDT(new Date())}'">⏱</button></div>`)).join('')}
      ${lab('Data carregamento', `<div style="${inpM};opacity:.85">${esc(br(c.dataCarregamento) || '—')}</div>`)}
    </div>
    <details style="margin-top:14px"><summary style="font-size:12px;cursor:pointer;color:var(--muted)">Histórico de alterações (${(c.log || []).length})</summary>
      <div style="font-size:11.5px;max-height:180px;overflow:auto;margin-top:6px">${(c.log || []).slice().reverse().map(l => `<div style="padding:3px 0;border-bottom:1px solid var(--border)"><span style="color:var(--muted)">${esc(brDataHora(l.em))} · ${esc(l.por)}</span> — ${esc(l.acao)}</div>`).join('') || '—'}</div></details>
    <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px;flex-wrap:wrap">
      <button class="btn btn-sm" style="color:#F87171;border-color:rgba(248,113,113,.4)" onclick="sobCancelarCiclo('${jsq(id)}')">🗑 Cancelar ciclo</button>
      <div style="display:flex;gap:8px"><button class="btn btn-sm" onclick="sobFecharModal()">Fechar</button><button class="btn btn-lime btn-sm" onclick="sobSalvarCiclo()">💾 Salvar</button></div>
    </div>`, 920);
};
window.sobSalvarCiclo = async () => {
  const c = st.ciclos.get(_mod?.id); if (!c) return;
  const g = id => document.getElementById(id)?.value ?? '';
  const patch = { etaData: g('sob-mod-eta'), etaTurno: g('sob-mod-turno'), dsr: g('sob-mod-dsr'), proxCarga: g('sob-mod-prox'), proxCargaHora: g('sob-mod-proxhora'), proxCargaNota: g('sob-mod-proxnota'), obs: g('sob-mod-obs').trim(), marcos: {} };
  MARCOS.forEach(m => { const v = g(`sob-mod-m-${m.id}`); patch.marcos[m.id] = v ? new Date(v).toISOString() : ''; });
  let renum = [];
  if (_mod.editavel) {
    const entregas = agrupar(_mod.linhas);
    if (!entregas.length) { alert('Inclua pelo menos um cliente com volume.'); return; }
    const [carrier, placa] = (g('sob-mod-placa') || '').split('||');
    Object.assign(patch, { entregas, placa: placa || '', carrier: carrier || '', dataCarregamento: g('sob-mod-data'), horaCarregamento: g('sob-mod-hora'), base: g('sob-mod-base').trim() });
    if (normPlaca(placa) !== normPlaca(c.placa) || patch.dataCarregamento !== c.dataCarregamento || patch.horaCarregamento !== c.horaCarregamento) renum = [normPlaca(c.placa), normPlaca(placa)];
    if (!placa) patch.cicloNum = null;
  }
  const mud = [];
  ['etaData', 'etaTurno', 'dsr', 'proxCarga', 'proxCargaHora', 'proxCargaNota', 'obs', 'placa', 'dataCarregamento', 'horaCarregamento', 'base'].forEach(k => { if (k in patch && (patch[k] || '') !== (c[k] || '')) mud.push(k); });
  MARCOS.forEach(m => { if ((patch.marcos[m.id] || '') !== (c.marcos?.[m.id] || '')) mud.push(m.label); });
  if (_mod.editavel && JSON.stringify(patch.entregas) !== JSON.stringify(c.entregas)) mud.push('entregas');
  try {
    await gravarCiclo(c.id, patch, mud.length ? `Editado: ${mud.join(', ')}` : null);
    for (const pn of renum.filter(Boolean)) await renumerarPlaca(pn);
    fecharModal(); toast('Ciclo salvo.', true); renderSobDemanda();
  } catch (e) { alert('Não consegui salvar: ' + (e.code || e.message)); }
};
window.sobCancelarCiclo = async id => {
  const c = st.ciclos.get(id); if (!c) return;
  const viaRot = c.origem === 'roteirizador' && !c.previa;
  if (!confirm(`Cancelar o ${c.cicloNum ? c.cicloNum + 'º ciclo' : 'ciclo'} de ${c.placa || 'sem caminhão'}?\n\nEle sai do reporte e os ciclos seguintes da placa são renumerados.${viaRot ? '\n\nAtenção: a viagem continua na programação do Histórico — se ela foi cancelada de fato, exclua a viagem lá também.' : ''}`)) return;
  try {
    await gravarCiclo(id, { cancelado: true, finalizado: true }, 'Ciclo cancelado');
    await renumerarPlaca(normPlaca(c.placa));
    fecharModal(); toast('Ciclo cancelado.', true); renderSobDemanda();
  } catch (e) { alert('Não consegui cancelar: ' + (e.code || e.message)); }
};

// ── Configuração: siglas, apelidos de cliente, contagem inicial ────────────
window.sobAbrirConfig = async () => {
  await carregarTudo(false);
  const produtos = new Set(), clientes = new Set();
  st.ciclos.forEach(c => (c.entregas || []).forEach(e => { if (e.cliente) clientes.add(e.cliente); (e.itens || []).forEach(it => it.produto && produtos.add(it.produto)); }));
  Object.keys(st.cfg.siglas || {}).forEach(p => produtos.add(p));
  Object.keys(st.cfg.apelidos || {}).forEach(c => clientes.add(c));
  const contagem = st.placas.map(p => { const k = normPlaca(p.placa); return { k, p }; });
  abrirModal(`
    <p style="font-size:15px;font-weight:600;color:var(--lime);margin-bottom:12px">⚙ Siglas, apelidos e contagem de ciclos</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px">
      <div>
        <div style="font-size:12px;font-weight:700;margin-bottom:4px">Produto → sigla do reporte</div>
        <p style="font-size:11px;color:var(--muted);margin-bottom:6px">Em branco = sigla automática (mostrada em cinza).</p>
        <div style="max-height:46vh;overflow:auto">${[...produtos].sort().map((p, i) => `<div style="display:flex;gap:6px;align-items:center;margin-bottom:4px"><span style="flex:1;font-size:11.5px">${esc(p)}</span><input data-sob-sigla="${esc(p)}" value="${esc(st.cfg.siglas?.[p] || '')}" placeholder="${esc(siglaAuto(p))}" style="${inpM};width:80px"></div>`).join('') || '<p style="font-size:12px;color:var(--muted)">Nenhum produto ainda.</p>'}</div>
      </div>
      <div>
        <div style="font-size:12px;font-weight:700;margin-bottom:4px">Cliente → nome curto no reporte</div>
        <p style="font-size:11px;color:var(--muted);margin-bottom:6px">Ex.: POSTO DE GASOLINA TREMENDAO LTDA → TREMENDÃO.</p>
        <div style="max-height:46vh;overflow:auto">${[...clientes].sort().map(c => `<div style="display:flex;gap:6px;align-items:center;margin-bottom:4px"><span style="flex:1;font-size:11.5px">${esc(c)}</span><input data-sob-apelido="${esc(c)}" value="${esc(st.cfg.apelidos?.[c] || '')}" placeholder="(nome completo)" style="${inpM};width:140px"></div>`).join('') || '<p style="font-size:12px;color:var(--muted)">Nenhum cliente ainda.</p>'}</div>
      </div>
    </div>
    <div style="font-size:12px;font-weight:700;margin:16px 0 4px">Contagem de ciclos já feitos por placa</div>
    <p style="font-size:11px;color:var(--muted);margin-bottom:6px">Nº do último ciclo FINALIZADO de cada caminhão (da planilha de hoje). Ex.: URP1B02 com 4 → a viagem em curso é o 5º. Os ciclos ativos são renumerados ao salvar.</p>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:6px">${contagem.map(({ k, p }) => `<label style="display:flex;gap:6px;align-items:center;font-size:12px"><span style="flex:1"><b style="font-family:'DM Mono',monospace">${esc(p.placa)}</b> <span style="color:var(--muted);font-size:10.5px">${esc(p.carrier)}</span></span><input type="number" min="0" data-sob-base="${esc(k)}" value="${Number(st.cfg.base?.[k]) || 0}" style="${inpM};width:64px;text-align:right"></label>`).join('')}</div>
    <div style="border:1px dashed var(--border2);border-radius:8px;padding:10px;margin-top:16px;display:flex;gap:10px;align-items:center;flex-wrap:wrap">
      <div style="flex:1;min-width:260px"><div style="font-size:12px;font-weight:700">📥 Importar ciclos finalizados</div><div style="font-size:11px;color:var(--muted)">Arquivo .json gerado a partir da planilha antiga. Pode importar de novo sem duplicar; a contagem de cada placa é ajustada sozinha.</div></div>
      <label class="btn btn-sm" style="cursor:pointer">Escolher arquivo…<input type="file" accept=".json,application/json" style="display:none" onchange="sobImportarArquivo(this.files[0]);this.value=''"></label>
    </div>
    <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px"><button class="btn btn-sm" onclick="sobFecharModal()">Cancelar</button><button class="btn btn-lime btn-sm" onclick="sobSalvarConfig()">💾 Salvar</button></div>`, 980);
};
window.sobSalvarConfig = async () => {
  const siglas = {}, apelidos = {}, base = {};
  document.querySelectorAll('[data-sob-sigla]').forEach(el => { const v = el.value.trim().toUpperCase(); if (v) siglas[el.dataset.sobSigla] = v; });
  document.querySelectorAll('[data-sob-apelido]').forEach(el => { const v = el.value.trim(); if (v) apelidos[el.dataset.sobApelido] = v; });
  document.querySelectorAll('[data-sob-base]').forEach(el => { base[el.dataset.sobBase] = Math.max(0, parseInt(el.value, 10) || 0); });
  try {
    // mergeFields: siglas/apelidos/base são substituídos por inteiro (permite apagar um item);
    // o resto do documento (estado dos caminhões) fica como está.
    await fb().setDoc(refCfg(), { siglas, apelidos, base: { ...(st.cfg.base || {}), ...base }, atualizadoEm: new Date().toISOString(), atualizadoPor: usuario() },
      { mergeFields: ['siglas', 'apelidos', 'base', 'atualizadoEm', 'atualizadoPor'] });
    await carregarCfg(true);
    const ativos = [...st.ciclos.values()].filter(c => !c.finalizado && !c.cancelado);
    for (const k of Object.keys(base)) await renumerarPlaca(k, ativos);
    fecharModal(); toast('Configuração salva.', true); renderSobDemanda();
  } catch (e) { alert('Não consegui salvar: ' + (e.code || e.message)); }
};

// ── Importação de ciclos finalizados (planilha antiga → .json) ─────────────
window.sobImportarArquivo = async file => {
  if (!file) return;
  let dados;
  try { dados = JSON.parse(await file.text()); } catch (e) { alert('Arquivo inválido (não é JSON).'); return; }
  const lista = Array.isArray(dados?.ciclos) ? dados.ciclos : null;
  if (dados?.tipo !== 'nexta_sob_ciclos_import' || !lista || !lista.length) { alert('Esse arquivo não é uma importação de ciclos da operação sob demanda.'); return; }
  await carregarCfg(true);
  const allP = await window.dbGetPlates();
  const { col, db, getDocs, query, where } = fb();
  const menorData = lista.map(c => c.dataCarregamento).filter(Boolean).sort()[0] || '0000';
  const existentes = [];
  const snap = await getDocs(query(col(db, 'sob_ciclos'), where('dataCarregamento', '>=', menorData)));
  snap.forEach(d => existentes.push({ id: d.id, ...d.data() }));
  const ativos = await carregarAtivos();
  ativos.forEach(a => { if (!existentes.some(e => e.id === a.id)) existentes.push(a); });
  const plano = lista.map(c => {
    const m = window._casarPlacaPainel?.(allP, c.carrier, c.placa);
    const placa = m ? m.p.placa : c.placa, carrier = m ? m.carrier : c.carrier;
    const pn = normPlaca(placa);
    const id = `imp__${pn}__${c.cicloNum}`;
    const dup = existentes.find(e => e.id !== id && !e.cancelado && e.origem !== 'importado' && normPlaca(e.placa) === pn && e.dataCarregamento === c.dataCarregamento);
    return { c, placa, carrier, pn, id, dup };
  });
  const porPlaca = {};
  plano.forEach(x => { porPlaca[x.pn] = Math.max(porPlaca[x.pn] || 0, Number(x.c.cicloNum) || 0); });
  const dups = plano.filter(x => x.dup);
  const resumo = Object.entries(porPlaca).sort().map(([k, n]) => `${k}: último finalizado ${n}º`).join('\n');
  if (!confirm(`Importar ${plano.length} ciclo(s) finalizado(s) de ${Object.keys(porPlaca).length} placa(s)?\n\n${resumo}${dups.length ? `\n\n${dups.length} já existe(m) no sistema (mesma placa e dia de carregamento) — esses não são duplicados: o ciclo do sistema fica, marcado como finalizado com o nº da planilha.` : ''}\n\nOs próximos ciclos de cada placa continuam a contagem a partir daí.`)) return;
  toast('Importando ciclos…', true);
  let novos = 0, atual = 0;
  try {
    for (const x of plano) {
      const c = x.c;
      if (x.dup) {
        const patch = { cicloNum: Number(c.cicloNum) || x.dup.cicloNum, status: 'finalizado', finalizado: true };
        if (!x.dup.etaData && c.etaData) patch.etaData = c.etaData;
        if (c.marcos) patch.marcos = Object.fromEntries(Object.entries(c.marcos).filter(([k]) => !x.dup.marcos?.[k]));
        st.ciclos.set(x.dup.id, x.dup);
        await gravarCiclo(x.dup.id, patch, 'Finalizado pela importação da planilha');
        atual++; continue;
      }
      await gravarCiclo(x.id, {
        origem: 'importado', previa: false, cancelado: false, finalizado: true, status: 'finalizado',
        cicloNum: Number(c.cicloNum) || null, placa: x.placa, carrier: x.carrier, base: c.base || '',
        dataCarregamento: c.dataCarregamento || '', horaCarregamento: c.horaCarregamento || '',
        etaData: c.etaData || '', etaTurno: c.etaTurno || '',
        entregas: (c.entregas || []).map(e => ({ cliente: e.cliente || '', pedidoId: '', itens: (e.itens || []).map(it => ({ produto: it.produto || '', volume: Number(it.volume) || 0 })) })),
        volumeInformado: Number(c.volumeInformado) || 0, produtoInformado: c.produtoInformado || '', clientesTextoOriginal: c.clientesTextoOriginal || '',
        marcos: c.marcos || {}, importadoDe: dados.origem || file.name, criadoEm: new Date().toISOString(), criadoPor: usuario(),
      }, `Importado da planilha (${dados.origem || file.name}, linha ${c.linha || '?'})`);
      novos++;
    }
    const base = {};
    Object.entries(porPlaca).forEach(([k, n]) => { base[k] = Math.max(Number(st.cfg.base?.[k]) || 0, n); });
    await salvarCfg({ base });
    const ativosAgora = await carregarAtivos();
    for (const k of Object.keys(base)) await renumerarPlaca(k, ativosAgora.filter(a => !a.cancelado));
    st.carregadoEm = 0;
    if (st.periodo < 45) st.periodo = 60;
    st.finalAberto = true;
    fecharModal();
    toast(`📥 ${novos} ciclo(s) importado(s)${atual ? `, ${atual} já existente(s) atualizado(s)` : ''}. Contagem das placas ajustada.`, true);
    renderSobDemanda();
  } catch (e) { alert(`Parei no meio da importação (${novos} gravados): ${e.code || e.message}\nPode importar de novo — não duplica.`); }
};

// ═══════════════════════════════════════════════════════════════════════════
// REPORTE — Excel formatado (ExcelJS) e PDF (jsPDF)
// ═══════════════════════════════════════════════════════════════════════════
const COLS_REP = [
  { h: 'CICLO', w: 7 }, { h: 'TRANSPORTADORA', w: 24 }, { h: 'PLACA', w: 10 }, { h: 'CLIENTE', w: 70 },
  { h: 'CARREGAMENTO', w: 17 }, { h: 'ETA DESTINO', w: 22 }, { h: 'STATUS', w: 22 }, { h: 'VOLUME', w: 9 },
  { h: 'PRODUTO', w: 10 }, { h: 'DSR', w: 12 }, { h: 'PRÓX. CARGA', w: 13 },
];
function linhaRep(r, secao) {
  if (r.tipo === 'veiculo') {
    const e = estById[r.estado] || estById.sem_pedido;
    return { v: [`${r.num}º`, (r.carrier || '').toUpperCase(), r.placa, r.obs || '', r.disponivelEm ? `${brCurto(r.disponivelEm)} às ${r.hora ? r.hora + 'h' : '--:--'}` : '-', '-', e.label, Number(r.capacidade) || '', '', '', ''], cor: e };
  }
  const c = r.c; const s = stById[c.status] || stById.programado;
  const eta = secao === 'finalizadas' ? (c.etaData ? `FINALIZADA ${br(c.etaData)}` : 'FINALIZADA') : etaTexto(c);
  return { v: [c.cicloNum ? `${c.cicloNum}º` : '', (c.carrier || '').toUpperCase(), c.placa || '', clientesTexto(c), carregTexto(c), eta, s.label, Math.round(volCiclo(c) * 10) / 10, produtoCiclo(c), c.dsr ? br(c.dsr) : '', proxTexto(c)], cor: s };
}
function secoesRep() {
  const s = montarSecoes();
  return [
    { t: 'ENTREGAS EM CURSO', k: 'emCurso', l: s.emCurso },
    { t: 'ENTREGAS PREVISTAS', k: 'previstas', l: s.previstas },
    { t: 'PEDIDOS PARA ALOCAR', k: 'alocar', l: s.alocar },
    { t: `ENTREGAS FINALIZADAS (ÚLTIMOS ${st.periodo} DIAS)`, k: 'finalizadas', l: s.finalizadas },
  ];
}
const nomeArq = () => { const d = new Date(); return `Operacao_sob_demanda_${d.getFullYear()}${p2(d.getMonth() + 1)}${p2(d.getDate())}_${p2(d.getHours())}${p2(d.getMinutes())}`; };
function carregarScript(src, teste) {
  return new Promise((ok, ko) => {
    if (teste()) return ok();
    const s = document.createElement('script'); s.src = src; s.onload = () => ok(); s.onerror = () => ko(new Error('não carregou ' + src));
    document.head.appendChild(s);
  });
}
const argb = hex => 'FF' + hex.replace('#', '').toUpperCase();

window.sobExportarExcel = async () => {
  try { await carregarTudo(false); await carregarScript('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js', () => !!window.ExcelJS); }
  catch (e) { alert('Não consegui carregar a biblioteca de Excel: ' + e.message); return; }
  const wb = new window.ExcelJS.Workbook();
  wb.creator = 'NEXTA Frota';
  const ws = wb.addWorksheet('Operação sob demanda', { views: [{ showGridLines: false }], pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } } });
  ws.columns = COLS_REP.map(c => ({ width: c.w }));
  const N = COLS_REP.length;
  const borda = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  const bordas = { top: borda, left: borda, bottom: borda, right: borda };
  const fonte = (o = {}) => ({ name: 'Calibri', size: 10, ...o });
  // Título
  ws.mergeCells(1, 1, 1, N);
  const t = ws.getCell(1, 1);
  t.value = 'OPERAÇÃO SOB DEMANDA — STATUS DOS CICLOS'; t.font = fonte({ size: 14, bold: true, color: { argb: 'FFFFFFFF' } });
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } }; t.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 26;
  ws.mergeCells(2, 1, 2, N);
  const sub = ws.getCell(2, 1);
  const agora = new Date();
  sub.value = `Atualizado em ${br(isoHoje(0))} às ${p2(agora.getHours())}:${p2(agora.getMinutes())}${st.transp ? ` · Transportadora: ${st.transp}` : ''}`;
  sub.font = fonte({ italic: true, color: { argb: 'FF6B7280' } }); sub.alignment = { horizontal: 'center' };
  let r = 4;
  for (const sec of secoesRep()) {
    ws.mergeCells(r, 1, r, N);
    const h = ws.getCell(r, 1);
    h.value = sec.t; h.font = fonte({ size: 11, bold: true }); h.alignment = { horizontal: 'center', vertical: 'middle' };
    h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; h.border = bordas;
    ws.getRow(r).height = 18; r++;
    const cab = ws.getRow(r);
    COLS_REP.forEach((c, i) => {
      const cel = cab.getCell(i + 1);
      cel.value = sec.k === 'finalizadas' && c.h === 'ETA DESTINO' ? 'ETA / FINALIZADA' : c.h;
      cel.font = fonte({ bold: true, size: 9.5 }); cel.border = { bottom: { style: 'medium', color: { argb: 'FF6B7280' } } };
      cel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
      cel.alignment = { horizontal: ['VOLUME', 'PRODUTO', 'CICLO', 'DSR', 'PRÓX. CARGA', 'STATUS'].includes(c.h) ? 'center' : 'left', vertical: 'middle' };
    });
    r++;
    if (!sec.l.length) {
      ws.mergeCells(r, 1, r, N);
      const v = ws.getCell(r, 1); v.value = '—'; v.font = fonte({ color: { argb: 'FF9CA3AF' } }); v.alignment = { horizontal: 'center' }; r++;
    }
    sec.l.forEach(item => {
      const { v, cor } = linhaRep(item, sec.k);
      const row = ws.getRow(r);
      v.forEach((val, i) => {
        const cel = row.getCell(i + 1);
        cel.value = val === '' ? null : val;
        cel.font = fonte({ bold: i === 10 && !!val });
        cel.border = bordas;
        cel.alignment = { vertical: 'middle', wrapText: i === 3, horizontal: [0, 6, 7, 8, 9, 10].includes(i) ? 'center' : 'left' };
      });
      const sc = row.getCell(7);
      sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(cor.bg) } };
      sc.font = fonte({ bold: true, size: 9, color: { argb: argb(cor.fg) } });
      if (item.tipo === 'veiculo') row.eachCell(cel => { if (cel.col !== 7) cel.font = fonte({ color: { argb: 'FF6B7280' } }); });
      const len = String(v[3] || '').length;
      row.height = Math.max(16, Math.ceil(len / 85) * 14);
      r++;
    });
    r++;
  }
  // Legenda (vertical, na coluna STATUS — mesma largura das células coloridas)
  const leg = [...STATUS.filter(s => s.id !== 'finalizado'), ...ESTADOS, stById.finalizado];
  ws.getCell(r, 7).value = 'LEGENDA'; ws.getCell(r, 7).font = fonte({ bold: true, size: 9 }); ws.getCell(r, 7).alignment = { horizontal: 'center' }; r++;
  leg.forEach(s => {
    const cel = ws.getCell(r, 7);
    cel.value = s.label; cel.font = fonte({ size: 8.5, bold: true, color: { argb: argb(s.fg) } });
    cel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(s.bg) } }; cel.alignment = { horizontal: 'center' }; cel.border = bordas;
    r++;
  });
  const buf = await wb.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a'); a.href = url; a.download = nomeArq() + '.xlsx'; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  toast('📊 Excel do reporte gerado.', true);
};

window.sobExportarPDF = async () => {
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) { alert('Biblioteca de PDF não carregada.'); return; }
  await carregarTudo(false);
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
  const W = pdf.internal.pageSize.getWidth();
  const limpa = s => String(s ?? '').replace(/³/g, '3');
  const rgb = hex => [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
  pdf.setFillColor(31, 41, 55); pdf.rect(0, 0, W, 38, 'F');
  pdf.setTextColor(255, 255, 255); pdf.setFont('helvetica', 'bold'); pdf.setFontSize(14);
  pdf.text('OPERAÇÃO SOB DEMANDA — STATUS DOS CICLOS', W / 2, 24, { align: 'center' });
  const agora = new Date();
  pdf.setTextColor(107, 114, 128); pdf.setFont('helvetica', 'italic'); pdf.setFontSize(8.5);
  pdf.text(`Atualizado em ${br(isoHoje(0))} às ${p2(agora.getHours())}:${p2(agora.getMinutes())}${st.transp ? ' · ' + st.transp : ''}`, W / 2, 52, { align: 'center' });
  // legenda
  const leg = [...STATUS.filter(s => s.id !== 'finalizado'), ...ESTADOS, stById.finalizado];
  let x = 24; pdf.setFont('helvetica', 'bold'); pdf.setFontSize(6.5);
  leg.forEach(s => { const w = pdf.getTextWidth(limpa(s.label)) + 10; pdf.setFillColor(...rgb(s.bg)); pdf.rect(x, 60, w, 12, 'F'); pdf.setTextColor(...rgb(s.fg)); pdf.text(limpa(s.label), x + 5, 68.5); x += w + 4; });
  let y = 84;
  const cw = [30, 84, 46, 214, 72, 80, 88, 38, 42, 46, 54];
  for (const sec of secoesRep()) {
    const head = [COLS_REP.map(c => sec.k === 'finalizadas' && c.h === 'ETA DESTINO' ? 'ETA / FINALIZADA' : c.h)];
    const linhas = sec.l.map(item => linhaRep(item, sec.k));
    const body = linhas.length ? linhas.map(l => l.v.map(v => limpa(v === '' ? '' : v))) : [[{ content: '—', colSpan: COLS_REP.length, styles: { halign: 'center', textColor: [156, 163, 175] } }]];
    pdf.autoTable({
      startY: y, margin: { left: 24, right: 24 },
      head: [[{ content: limpa(sec.t), colSpan: COLS_REP.length, styles: { halign: 'center', fillColor: [217, 217, 217], textColor: [17, 24, 39], fontSize: 9.5 } }], ...head],
      body, theme: 'grid',
      styles: { fontSize: 7.2, cellPadding: 3, lineColor: [210, 210, 210], lineWidth: 0.4, textColor: [31, 41, 55], valign: 'middle' },
      headStyles: { fillColor: [243, 244, 246], textColor: [17, 24, 39], fontStyle: 'bold', fontSize: 6.6, overflow: 'visible' },
      columnStyles: Object.fromEntries(cw.map((w, i) => [i, { cellWidth: w, halign: [0, 6, 7, 8, 9, 10].includes(i) ? 'center' : 'left' }])),
      didParseCell: d => {
        if (d.section !== 'body' || !linhas.length) return;
        const l = linhas[d.row.index];
        if (d.column.index === 6) { d.cell.styles.fillColor = rgb(l.cor.bg); d.cell.styles.textColor = rgb(l.cor.fg); d.cell.styles.fontStyle = 'bold'; d.cell.styles.fontSize = 6.6; }
        if (d.column.index === 10 && d.cell.raw) d.cell.styles.fontStyle = 'bold';
      },
    });
    y = pdf.lastAutoTable.finalY + 14;
  }
  const n = pdf.internal.getNumberOfPages();
  for (let i = 1; i <= n; i++) { pdf.setPage(i); pdf.setFont('helvetica', 'normal'); pdf.setFontSize(7); pdf.setTextColor(156, 163, 175); pdf.text(`NEXTA · página ${i} de ${n}`, W - 24, pdf.internal.pageSize.getHeight() - 12, { align: 'right' }); }
  pdf.save(nomeArq() + '.pdf');
  toast('📄 PDF do reporte gerado.', true);
};

window.renderSobDemanda = renderSobDemanda;
// Testes/diagnóstico
window._sobDemandaInterno = { st, parseColagem, siglaAuto, clientesTexto, produtoCiclo, montarSecoes, renumerarPlaca };
})();
