// ═══════════════════════════════════════════════════════════════════════════
// NEXTA — km-utils.js
// ═══════════════════════════════════════════════════════════════════════════
// Fonte ÚNICA de verdade para "quantos km essa viagem rodou" — usada pelo
// Dashboard, pelo relatório de Frete, e pelas exportações. Antes desse
// arquivo existir, cada tela tinha sua própria conta (ligeiramente
// diferente), o que já causou dois bugs reais em produção:
//   1) O relatório de Frete dobrava o km (modo "ida e volta") mesmo quando
//      o km já vinha do trajeto real completo (_kmAjustado), contando a
//      viagem 2x.
//   2) Uma parada com coordenada (0,0) — pedido sem lat/lon cadastrado —
//      era tratada como coordenada válida, mandando o cálculo de rota real
//      pro meio do Oceano Atlântico ("Null Island").
// Este módulo é JS puro (sem depender de `window`, DOM, Firestore ou
// variáveis globais do roteirizador) — pode ser carregado tanto no
// navegador (via <script>, expõe window.NextaKm) quanto no Node (via
// require, usado pelos testes em /tests).
// ═══════════════════════════════════════════════════════════════════════════

(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(); // Node / Vitest
  } else {
    root.NextaKm = factory(); // navegador
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ─── Distância em linha reta entre duas coordenadas (Haversine) ──────────
  function haversineKm(lat1, lon1, lat2, lon2) {
    const R = 6371;
    const toRad = (deg) => (deg * Math.PI) / 180;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    return R * 2 * Math.asin(Math.sqrt(a));
  }

  // ─── Coordenada "válida" ────────────────────────────────────────────────
  // NÃO basta checar isNaN: 0 não é NaN, então um pedido com lat/lon
  // zerado/ausente no cadastro (comum em dado antigo) passava como
  // "coordenada real" — e o cálculo de rota ia parar em (0,0), no meio do
  // Atlântico, perto da África ("Null Island"). Esse foi o bug que inflou
  // o km total do histórico em ~14x numa rodada de recálculo real.
  // Qualquer coordenada a menos de ~111m de (0,0) é tratada como ausente.
  function coordenadaValida(lat, lon) {
    const la = typeof lat === 'number' ? lat : parseFloat(lat);
    const lo = typeof lon === 'number' ? lon : parseFloat(lon);
    return !isNaN(la) && !isNaN(lo) && (Math.abs(la) > 0.001 || Math.abs(lo) > 0.001);
  }

  // ─── Monta a sequência de pontos de uma viagem (terminal → paradas → retorno) ──
  // Não usa nenhuma variável global — recebe tudo por parâmetro, pra poder
  // ser usada tanto numa roteirização "ao vivo" (dados na memória) quanto
  // processando um arquivo antigo do histórico (dados lidos do disco).
  // `terms`: array de terminais do MESMO arquivo/snapshot da viagem (nunca
  // usa a lista de terminais "atual" do cadastro ao vivo — um arquivo
  // antigo pode ter terminais que não existem mais, ou em local diferente).
  function montarPontosViagem(v, vi, terms) {
    if (!v || !vi || !vi.paradas || !vi.paradas.length) return [];
    const terminalNome = vi.terminalOrigem || v.terminal || vi.paradas[0]?.pedido?.terminal || '';
    const terminal = (terms || []).find((t) => t.nome === terminalNome);
    const pontos = [];
    const tLat = parseFloat(terminal?.lat);
    const tLon = parseFloat(terminal?.lon);
    const terminalValido = !!terminal && coordenadaValida(tLat, tLon);

    if (terminalValido) {
      pontos.push({ lat: tLat, lon: tLon, tag: `origem:${terminal.nome}` });
    }
    vi.paradas.forEach((p, i) => {
      const lat = parseFloat(p.pedido?.lat ?? p.lat);
      const lon = parseFloat(p.pedido?.lon ?? p.lon);
      if (coordenadaValida(lat, lon)) {
        pontos.push({ lat, lon, tag: `parada${i + 1}:${p.pedido?.cliente || '?'}` });
      }
    });
    const ultimaParada = vi.paradas[vi.paradas.length - 1];
    if (terminalValido && (ultimaParada?.deslocVazioMin || 0) > 0) {
      pontos.push({ lat: tLat, lon: tLon, tag: `retorno:${terminal.nome}` });
    }
    return pontos;
  }

  // ─── Estimativa "antiga" (linha reta, soma de distanciaKm por parada) ────
  // Mantida como fallback quando não há km real calculado ainda, e como
  // referência pra trava de sanidade (comparar real vs. essa estimativa).
  function estimativaLinhaReta(vi) {
    if (!vi || !vi.paradas) return 0;
    return vi.paradas.reduce((s, p) => s + (p.distanciaKm || 0), 0);
  }

  // ─── Km de uma viagem: fonte única usada por Dashboard, Frete e exports ──
  // Prioridade: vi._kmAjustado (trajeto real via rota — calculado no
  // salvamento, no recálculo em lote do Dashboard, ou por ajuste manual no
  // mapa) > estimativa antiga por linha reta.
  // `_kmAjustado`, quando presente, JÁ é o percurso completo da viagem
  // (terminal → paradas → retorno, quando há retorno) — diferente da
  // estimativa antiga, que é só a soma das distâncias diretas terminal→
  // cliente. Ver kmEfetivo() para o motivo disso importar no cálculo de
  // custo de frete.
  function obterKmViagem(vi) {
    if (vi && typeof vi._kmAjustado === 'number' && vi._kmAjustado > 0) {
      return { km: vi._kmAjustado, real: true };
    }
    return { km: estimativaLinhaReta(vi), real: false };
  }

  // ─── Km "efetivo" pro cálculo de custo de Frete ───────────────────────────
  // Contratos podem ter kmModo 'ida' (só o trecho de ida) ou 'ida_volta'
  // (dobra o km — padrão). Essa duplicação só faz sentido quando o km de
  // entrada é a estimativa antiga (só ida, em linha reta). Quando o km já
  // vem do trajeto real (obterKmViagem().real === true), ele JÁ inclui a
  // volta ao terminal quando existe — dobrar de novo conta a viagem 2x.
  // Esse foi o bug que inflou o "KM Total (est.)" do relatório de Frete em
  // ~2x depois que o km real passou a ser usado.
  function kmEfetivo(kmInfo, kmModo) {
    if (!kmInfo) return 0;
    if (kmInfo.real) return kmInfo.km;
    const modo = kmModo || 'ida_volta';
    return modo === 'ida' ? kmInfo.km : kmInfo.km * 2;
  }

  return {
    haversineKm,
    coordenadaValida,
    montarPontosViagem,
    estimativaLinhaReta,
    obterKmViagem,
    kmEfetivo,
  };
});
