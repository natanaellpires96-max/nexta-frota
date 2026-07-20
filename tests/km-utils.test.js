// ═══════════════════════════════════════════════════════════════════════════
// Testes de assets/km-utils.js — cobre os dois bugs reais que já aconteceram
// em produção (coordenada 0,0 tratada como válida, e km duplicado no Frete),
// pra impedir que voltem a acontecer sem serem pegos antes do deploy.
// Rodar com: npm test
// ═══════════════════════════════════════════════════════════════════════════
import { describe, it, expect } from 'vitest';
import NextaKm from '../assets/km-utils.js';

describe('coordenadaValida', () => {
  it('rejeita (0, 0) — o bug da "Null Island"', () => {
    expect(NextaKm.coordenadaValida(0, 0)).toBe(false);
  });

  it('rejeita coordenadas quase-zero (dentro de ~111m de 0,0)', () => {
    expect(NextaKm.coordenadaValida(0.0005, 0.0005)).toBe(false);
    expect(NextaKm.coordenadaValida(-0.0009, 0)).toBe(false);
  });

  it('rejeita NaN / undefined / string inválida', () => {
    expect(NextaKm.coordenadaValida(NaN, -46.6)).toBe(false);
    expect(NextaKm.coordenadaValida(undefined, undefined)).toBe(false);
    expect(NextaKm.coordenadaValida('abc', -46.6)).toBe(false);
  });

  it('aceita coordenada real de São Paulo', () => {
    expect(NextaKm.coordenadaValida(-23.5505, -46.6333)).toBe(true);
  });

  it('aceita coordenada real vinda como string (parseFloat)', () => {
    expect(NextaKm.coordenadaValida('-23.5505', '-46.6333')).toBe(true);
  });

  it('aceita coordenada legitimamente perto do equador/meridiano (não é 0,0 de verdade)', () => {
    // ex.: cidade a 5km do meridiano de Greenwich, bem longe do equador
    expect(NextaKm.coordenadaValida(51.5, 0.05)).toBe(true);
  });
});

describe('haversineKm', () => {
  it('distância entre dois pontos idênticos é 0', () => {
    expect(NextaKm.haversineKm(-23.55, -46.63, -23.55, -46.63)).toBeCloseTo(0, 5);
  });

  it('distância São Paulo → Rio de Janeiro é ~360km (linha reta)', () => {
    // SP: -23.5505, -46.6333 | RJ: -22.9068, -43.1729
    const d = NextaKm.haversineKm(-23.5505, -46.6333, -22.9068, -43.1729);
    expect(d).toBeGreaterThan(340);
    expect(d).toBeLessThan(380);
  });
});

describe('estimativaLinhaReta', () => {
  it('soma distanciaKm de todas as paradas', () => {
    const vi = { paradas: [{ distanciaKm: 10 }, { distanciaKm: 25.5 }, { distanciaKm: 0 }] };
    expect(NextaKm.estimativaLinhaReta(vi)).toBeCloseTo(35.5, 5);
  });

  it('trata parada sem distanciaKm como 0', () => {
    const vi = { paradas: [{ distanciaKm: 10 }, {}] };
    expect(NextaKm.estimativaLinhaReta(vi)).toBeCloseTo(10, 5);
  });

  it('viagem sem paradas retorna 0', () => {
    expect(NextaKm.estimativaLinhaReta({ paradas: [] })).toBe(0);
    expect(NextaKm.estimativaLinhaReta(null)).toBe(0);
  });
});

describe('obterKmViagem', () => {
  it('usa _kmAjustado quando presente e positivo, marcando real=true', () => {
    const vi = { _kmAjustado: 152.3, paradas: [{ distanciaKm: 40 }] };
    const info = NextaKm.obterKmViagem(vi);
    expect(info.km).toBe(152.3);
    expect(info.real).toBe(true);
  });

  it('cai pra estimativa em linha reta quando _kmAjustado ausente, marcando real=false', () => {
    const vi = { paradas: [{ distanciaKm: 40 }, { distanciaKm: 10 }] };
    const info = NextaKm.obterKmViagem(vi);
    expect(info.km).toBe(50);
    expect(info.real).toBe(false);
  });

  it('trata _kmAjustado = 0 ou negativo como ausente (não confia em valor inválido)', () => {
    const vi = { _kmAjustado: 0, paradas: [{ distanciaKm: 40 }] };
    expect(NextaKm.obterKmViagem(vi).real).toBe(false);
    expect(NextaKm.obterKmViagem(vi).km).toBe(40);
  });
});

describe('kmEfetivo — regressão do bug de km duplicado no relatório de Frete', () => {
  it('NÃO dobra quando o km já é o trajeto real completo (real=true)', () => {
    const kmInfo = { km: 344, real: true };
    expect(NextaKm.kmEfetivo(kmInfo, 'ida_volta')).toBe(344);
    expect(NextaKm.kmEfetivo(kmInfo, 'ida')).toBe(344); // real=true ignora o modo do contrato
  });

  it('dobra a estimativa antiga (real=false) no modo "ida_volta" (padrão)', () => {
    const kmInfo = { km: 100, real: false };
    expect(NextaKm.kmEfetivo(kmInfo, 'ida_volta')).toBe(200);
    expect(NextaKm.kmEfetivo(kmInfo)).toBe(200); // modo ausente = ida_volta por padrão
  });

  it('NÃO dobra a estimativa antiga no modo "ida"', () => {
    const kmInfo = { km: 100, real: false };
    expect(NextaKm.kmEfetivo(kmInfo, 'ida')).toBe(100);
  });

  it('entrada nula/vazia retorna 0 sem quebrar', () => {
    expect(NextaKm.kmEfetivo(null, 'ida_volta')).toBe(0);
  });
});

describe('montarPontosViagem', () => {
  const terms = [{ nome: 'Terminal A', lat: -23.5, lon: -46.6 }];
  const v = { terminal: 'Terminal A' };

  it('monta origem + paradas válidas, na ordem', () => {
    const vi = {
      paradas: [
        { pedido: { lat: -23.4, lon: -46.5, cliente: 'Cliente 1' } },
        { pedido: { lat: -23.3, lon: -46.4, cliente: 'Cliente 2' } },
      ],
    };
    const pontos = NextaKm.montarPontosViagem(v, vi, terms);
    expect(pontos.length).toBe(3); // origem + 2 paradas
    expect(pontos[0].tag).toContain('origem');
    expect(pontos[1].lat).toBe(-23.4);
    expect(pontos[2].lat).toBe(-23.3);
  });

  it('EXCLUI parada com coordenada (0,0) — regressão do bug da "Null Island"', () => {
    const vi = {
      paradas: [
        { pedido: { lat: -23.4, lon: -46.5, cliente: 'Cliente OK' } },
        { pedido: { lat: 0, lon: 0, cliente: 'Cliente sem coordenada' } },
      ],
    };
    const pontos = NextaKm.montarPontosViagem(v, vi, terms);
    // origem + só 1 parada válida (a de coordenada 0,0 foi excluída)
    expect(pontos.length).toBe(2);
    expect(pontos.some(p => p.lat === 0 && p.lon === 0)).toBe(false);
  });

  it('adiciona ponto de retorno ao terminal quando deslocVazioMin > 0 na última parada', () => {
    const vi = {
      paradas: [
        { pedido: { lat: -23.4, lon: -46.5, cliente: 'Cliente 1' }, deslocVazioMin: 30 },
      ],
    };
    const pontos = NextaKm.montarPontosViagem(v, vi, terms);
    expect(pontos.length).toBe(3); // origem + parada + retorno
    expect(pontos[2].tag).toContain('retorno');
  });

  it('não adiciona retorno quando deslocVazioMin é 0/ausente', () => {
    const vi = { paradas: [{ pedido: { lat: -23.4, lon: -46.5, cliente: 'Cliente 1' } }] };
    const pontos = NextaKm.montarPontosViagem(v, vi, terms);
    expect(pontos.length).toBe(2); // só origem + parada
  });

  it('terminal não encontrado na lista: não quebra, só fica sem ponto de origem', () => {
    const vOutro = { terminal: 'Terminal Inexistente' };
    const vi = { paradas: [{ pedido: { lat: -23.4, lon: -46.5, cliente: 'Cliente 1' } }] };
    const pontos = NextaKm.montarPontosViagem(vOutro, vi, terms);
    expect(pontos.length).toBe(1); // só a parada, sem origem
  });

  it('viagem sem paradas retorna array vazio', () => {
    expect(NextaKm.montarPontosViagem(v, { paradas: [] }, terms)).toEqual([]);
    expect(NextaKm.montarPontosViagem(v, null, terms)).toEqual([]);
  });
});
