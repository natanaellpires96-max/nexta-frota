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
  // ─────────────────────────────────────────────────────────────────────────
  // BASE GERADA A PARTIR DO OPENSTREETMAP (consulta Overpass em 08/07/2026,
  // raio de 150km ao redor de: São Caetano do Sul, Paulínia, São José dos
  // Campos, Ribeirão Preto, Uberlândia, Uberaba, Betim, Brasília e Duque de
  // Caxias — as 9 operações cadastradas no sistema pai na data da consulta).
  // 244 praças/pórticos com tarifa de caminhão confirmada (tag OSM
  // "BRL/hgv/axle"). Cobre praças físicas E pórticos free-flow (marcados com
  // modelo: 'free_flow' abaixo).
  // tarifas: { 2: X } — X é o preço para 2 eixos (= preço-por-eixo do OSM × 2);
  // obterTarifaPedagio() já deriva proporcionalmente para 3/4/5/6+ eixos a
  // partir disso, então não precisa (nem deve) editar esse cálculo.
  // "rodovia" nesta base geralmente traz a concessionária/operador (OSM não
  // tinha o campo de rodovia preenchido na maioria dos nós) — ainda serve pra
  // identificar a praça no alerta, só não é o nome oficial da via.
  // Praças com cobrança separada por sentido (ex.: "sentido Norte"/"sentido
  // Sul") foram mantidas como entradas DISTINTAS de propósito — são cabines
  // fisicamente separadas nas pistas opostas, então uma rota num sentido só
  // deve mesmo bater perto de uma delas.
  // 19 pontos do OSM ficaram de fora por não terem tarifa de caminhão
  // cadastrada (balsas, pedágios ambientais/municipais, links quebrados) —
  // se alguma rota real passar por um desses, me avise o nome que eu confirmo
  // a tarifa numa fonte separada e adiciono manualmente.
  // Se uma operação nova for cadastrada fora desse raio de 150km das 9 cidades
  // acima, pedágios daquela região não vão aparecer aqui — regenerar a
  // consulta Overpass com a cidade nova incluída.
  { nome: 'Aeroporto', lat: -23.457, lon: -46.4792, tarifas: { 2: 10.88 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Aguaí', lat: -22.0139, lon: -46.8506, tarifas: { 2: 13.2 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Agulha', lat: -21.3831, lon: -48.7315, tarifas: { 2: 27.6 }, rodovia: 'Eco/vias Noroeste', regiao: 'BR' },
  { nome: 'Alambari', lat: -23.5528, lon: -47.7778, tarifas: { 2: 24.2 }, rodovia: 'CCR SPVias', regiao: 'BR' },
  { nome: 'Alexânia', lat: -16.1157, lon: -48.5893, tarifas: { 2: 14.8 }, rodovia: 'Triunfo Concebra', regiao: 'BR' },
  { nome: 'Anchieta - P5', lat: -23.7565, lon: -46.542, tarifas: { 2: 10.8 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Anhembi', lat: -22.9479, lon: -48.2858, tarifas: { 2: 22.0 }, rodovia: 'Rodovias do Tietê', regiao: 'BR' },
  { nome: 'Araguari', lat: -18.8668, lon: -48.0165, tarifas: { 2: 28.0 }, rodovia: 'EPR Triângulo', regiao: 'BR' },
  { nome: 'Araguari 1 - Norte', lat: -18.5394, lon: -48.0467, tarifas: { 2: 14.6 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Araguari 1 - Sul', lat: -18.5394, lon: -48.0468, tarifas: { 2: 14.6 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Araguari 2 - Norte', lat: -18.7499, lon: -48.2355, tarifas: { 2: 11.0 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Araguari 2 - Sul', lat: -18.7497, lon: -48.2356, tarifas: { 2: 11.0 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Araraquara', lat: -21.7532, lon: -48.2617, tarifas: { 2: 45.6 }, rodovia: 'EcoVias Noroeste', regiao: 'BR' },
  { nome: 'Araras', lat: -22.3683, lon: -47.2167, tarifas: { 2: 19.6 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Araçoiaba', lat: -23.511, lon: -47.5652, tarifas: { 2: 8.86 }, rodovia: 'CCR Sorocabana', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Aricanduva', lat: -23.4958, lon: -46.5559, tarifas: { 2: 9.03 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Arujá', lat: -23.4188, lon: -46.3318, tarifas: { 2: 4.38 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Arujá - P10 - Externo', lat: -23.4205, lon: -46.3617, tarifas: { 2: 7.0 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Atibaia', lat: -23.0944, lon: -46.6261, tarifas: { 2: 21.2 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Azurita', lat: -20.0396, lon: -44.5331, tarifas: { 2: 15.2 }, rodovia: 'MG-050', regiao: 'BR' },
  { nome: 'Barueri', lat: -23.5099, lon: -46.8172, tarifas: { 2: 8.4 }, rodovia: 'Ecovias Raposo Castelo', regiao: 'BR' },
  { nome: 'Batatais', lat: -20.943, lon: -47.6335, tarifas: { 2: 23.4 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Bertioga', lat: -23.7555, lon: -46.0397, tarifas: { 2: 27.8 }, rodovia: 'CNL Novo Litoral', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Boa Esperança do Sul', lat: -22.0273, lon: -48.4032, tarifas: { 2: 24.0 }, rodovia: 'Arteris ViaPaulista', regiao: 'BR' },
  { nome: 'Boituva', lat: -23.3156, lon: -47.6395, tarifas: { 2: 27.6 }, rodovia: 'AB Colinas', regiao: 'BR' },
  { nome: 'Borda Da Mata', lat: -22.2897, lon: -46.199, tarifas: { 2: 20.2 }, rodovia: 'Km 34,5', regiao: 'BR' },
  { nome: 'Brotas', lat: -22.2638, lon: -47.9025, tarifas: { 2: 20.0 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'Cachoeiras do Macacu', lat: -22.4164, lon: -42.623, tarifas: { 2: 25.6 }, rodovia: 'Rota 116', regiao: 'BR' },
  { nome: 'Caeté', lat: -19.7439, lon: -43.6153, tarifas: { 2: 31.0 }, rodovia: 'P1C', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Cambuí', lat: -22.6286, lon: -46.0779, tarifas: { 2: 6.4 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'Campo Alegre de Goiás - Norte', lat: -17.7721, lon: -47.7498, tarifas: { 2: 19.2 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Campo Alegre de Goiás - Sul', lat: -17.7719, lon: -47.75, tarifas: { 2: 19.2 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Campo Florido', lat: -19.7769, lon: -48.4582, tarifas: { 2: 28.4 }, rodovia: 'Way-262', regiao: 'BR' },
  { nome: 'Capim Branco', lat: -19.6012, lon: -44.2219, tarifas: { 2: 31.0 }, rodovia: 'Via Cristais', regiao: 'BR' },
  { nome: 'Caraguatatuba', lat: -23.6698, lon: -45.467, tarifas: { 2: 11.0 }, rodovia: 'SPI-097/055', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Carmópolis de Minas', lat: -20.5918, lon: -44.7013, tarifas: { 2: 6.4 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'Casa Branca', lat: -21.919, lon: -47.0518, tarifas: { 2: 18.8 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Casimiro de Abreu', lat: -22.4759, lon: -42.0888, tarifas: { 2: 15.0 }, rodovia: 'Autopista Fluminense', regiao: 'BR' },
  { nome: 'Catiguá', lat: -21.0682, lon: -49.105, tarifas: { 2: 39.6 }, rodovia: 'EcoVias Noroeste', regiao: 'BR' },
  { nome: 'Caçapava', lat: -23.1559, lon: -45.6979, tarifas: { 2: 11.0 }, rodovia: 'Ecopistas', regiao: 'BR' },
  { nome: 'Colina - P8', lat: -20.6886, lon: -48.5369, tarifas: { 2: 23.4 }, rodovia: 'EcoVias Noroeste', regiao: 'BR' },
  { nome: 'Conchas', lat: -23.0311, lon: -47.9846, tarifas: { 2: 19.4 }, rodovia: 'Rodovias do Tietê', regiao: 'BR' },
  { nome: 'Conselheiro Lafaiete', lat: -20.767, lon: -43.8067, tarifas: { 2: 27.2 }, rodovia: 'EPR Via Mineira', regiao: 'BR' },
  { nome: 'Cordeiro', lat: -22.0546, lon: -42.3624, tarifas: { 2: 25.6 }, rodovia: 'Rota 116', regiao: 'BR' },
  { nome: 'Corinto', lat: -18.8673, lon: -44.4959, tarifas: { 2: 21.2 }, rodovia: 'EcoVias NorteMinas', regiao: 'BR' },
  { nome: 'Curvelo', lat: -19.0196, lon: -44.5131, tarifas: { 2: 22.3 }, rodovia: 'Via Cristais', regiao: 'BR' },
  { nome: 'Córrego das Colheres', lat: -20.1935, lon: -44.9977, tarifas: { 2: 18.2 }, rodovia: 'MG-050', regiao: 'BR' },
  { nome: 'Delta - Norte', lat: -19.9156, lon: -47.8292, tarifas: { 2: 11.2 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Delta - Sul', lat: -19.9157, lon: -47.8293, tarifas: { 2: 11.2 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Descalvado', lat: -21.8696, lon: -47.5358, tarifas: { 2: 18.4 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Diadema', lat: -23.682, lon: -46.6108, tarifas: { 2: 5.8 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Dobrada', lat: -21.5576, lon: -48.3327, tarifas: { 2: 22.4 }, rodovia: '', regiao: 'BR' },
  { nome: 'Dois Córregos', lat: -22.256, lon: -48.2438, tarifas: { 2: 22.8 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'EIXO-SP', lat: -22.4364, lon: -47.9043, tarifas: { 2: 19.12 }, rodovia: 'EIXO-SP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'EPR Sul de Minas', lat: -22.6535, lon: -45.7483, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul de Minas', regiao: 'BR' },
  { nome: 'Ecovias', lat: -23.7352, lon: -46.6006, tarifas: { 2: 14.8 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Ecovias RioMinas', lat: -22.7163, lon: -43.7167, tarifas: { 2: 34.4 }, rodovia: 'Ecovias RioMinas', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Eixo - SP', lat: -22.129, lon: -47.8081, tarifas: { 2: 14.0 }, rodovia: 'Eixo - SP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Elisiário', lat: -21.1576, lon: -49.1247, tarifas: { 2: 3.1 }, rodovia: 'Jotec', regiao: 'BR' },
  { nome: 'Engenheiro Coelho', lat: -22.5019, lon: -47.2214, tarifas: { 2: 21.8 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Espírito Santo do Pinhal', lat: -22.2523, lon: -46.8135, tarifas: { 2: 26.2 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Estiva Gerbi', lat: -22.1635, lon: -46.9919, tarifas: { 2: 21.0 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Fernão Dias', lat: -23.4978, lon: -46.5601, tarifas: { 2: 10.88 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Florestal', lat: -19.9098, lon: -44.5067, tarifas: { 2: 20.6 }, rodovia: 'Way-262', regiao: 'BR' },
  { nome: 'Fronteira', lat: -20.1425, lon: -49.1106, tarifas: { 2: 24.0 }, rodovia: 'Way-153', regiao: 'BR' },
  { nome: 'Goianápolis', lat: -16.4384, lon: -49.0191, tarifas: { 2: 10.8 }, rodovia: 'Triunfo Concebra', regiao: 'BR' },
  { nome: 'Gonçalves - P1', lat: -22.6535, lon: -45.7482, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul de Minas', regiao: 'BR' },
  { nome: 'Guapimirim', lat: -22.672, lon: -42.9804, tarifas: { 2: 42.0 }, rodovia: 'EcoVias RioMinas', regiao: 'BR' },
  { nome: 'Guararema', lat: -23.3634, lon: -46.1557, tarifas: { 2: 10.05 }, rodovia: 'Ecopistas', regiao: 'BR' },
  { nome: 'Guararema 2', lat: -23.3856, lon: -46.1566, tarifas: { 2: 10.8 }, rodovia: 'Ecopistas', regiao: 'BR' },
  { nome: 'Guarulhos', lat: -23.3913, lon: -46.408, tarifas: { 2: 9.24 }, rodovia: 'Via Appia SP Serra', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Guatapará', lat: -21.5717, lon: -47.9487, tarifas: { 2: 36.6 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Igaratá', lat: -23.1987, lon: -46.1806, tarifas: { 2: 26.6 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Imigrantes - P4', lat: -23.7638, lon: -46.5815, tarifas: { 2: 8.6 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Imigrantes Capital - P3', lat: -23.7641, lon: -46.5927, tarifas: { 2: 8.6 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Imigrantes Litoral - P2', lat: -23.7655, lon: -46.5938, tarifas: { 2: 8.6 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Indaiatuba', lat: -23.058, lon: -47.1495, tarifas: { 2: 37.5 }, rodovia: 'AB Colinas', regiao: 'BR' },
  { nome: 'Ipameri - Norte', lat: -17.1212, lon: -47.7214, tarifas: { 2: 17.8 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Ipameri - Sul', lat: -17.1212, lon: -47.7216, tarifas: { 2: 17.8 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Iracemápolis', lat: -22.6544, lon: -47.5189, tarifas: { 2: 17.2 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Itabirito', lat: -20.2729, lon: -43.9512, tarifas: { 2: 27.2 }, rodovia: 'EPR Via Mineira', regiao: 'BR' },
  { nome: 'Itaboraí', lat: -22.7157, lon: -42.8114, tarifas: { 2: 25.6 }, rodovia: 'Rota 116', regiao: 'BR' },
  { nome: 'Itaguaí', lat: -22.8487, lon: -43.8315, tarifas: { 2: 22.0 }, rodovia: 'BR-101', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Itapevi', lat: -23.5181, lon: -46.9408, tarifas: { 2: 19.6 }, rodovia: 'Ecovias Raposo Castelo', regiao: 'BR' },
  { nome: 'Itaquaquecetuba', lat: -23.4663, lon: -46.3682, tarifas: { 2: 11.4 }, rodovia: 'Ecopistas', regiao: 'BR' },
  { nome: 'Itaquaquecetuba - P9 - Externa', lat: -23.4583, lon: -46.3442, tarifas: { 2: 8.2 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Itatiaia', lat: -22.4949, lon: -44.5696, tarifas: { 2: 29.0 }, rodovia: 'CCR RioSP', regiao: 'BR' },
  { nome: 'Itatiaiuçu', lat: -20.2684, lon: -44.4238, tarifas: { 2: 6.4 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'Itatiba', lat: -22.9542, lon: -46.8595, tarifas: { 2: 30.47 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Itatinga', lat: -23.0827, lon: -48.515, tarifas: { 2: 38.8 }, rodovia: 'CCR SPVias', regiao: 'BR' },
  { nome: 'Itirapina', lat: -22.129, lon: -47.808, tarifas: { 2: 14.0 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'Itobi', lat: -21.7059, lon: -46.9467, tarifas: { 2: 26.8 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Itu', lat: -23.4065, lon: -47.3016, tarifas: { 2: 26.4 }, rodovia: 'CCR Sorocabana', regiao: 'BR' },
  { nome: 'Ituiutaba', lat: -18.9623, lon: -49.5918, tarifas: { 2: 11.8 }, rodovia: 'Ecovias do Cerrado', regiao: 'BR' },
  { nome: 'Itumbiara', lat: -18.2685, lon: -49.2459, tarifas: { 2: 29.8 }, rodovia: 'Way-153', regiao: 'BR' },
  { nome: 'Itupeva', lat: -23.2354, lon: -47.0426, tarifas: { 2: 21.2 }, rodovia: 'AB Colinas', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Itápolis - P7', lat: -21.5436, lon: -48.7881, tarifas: { 2: 19.4 }, rodovia: 'EcoVias Noroeste', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Jaboticabal', lat: -21.2294, lon: -48.2339, tarifas: { 2: 33.8 }, rodovia: '', regiao: 'BR' },
  { nome: 'Jacareí', lat: -23.2994, lon: -46.0165, tarifas: { 2: 16.2 }, rodovia: 'CCR NovaDutra', regiao: 'BR' },
  { nome: 'Jaguariuna', lat: -22.7711, lon: -47.022, tarifas: { 2: 35.2 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Jambeiro', lat: -23.2958, lon: -45.7792, tarifas: { 2: 11.6 }, rodovia: 'Tamoios', regiao: 'BR' },
  { nome: 'Jaraguá - P7', lat: -15.8271, lon: -49.2816, tarifas: { 2: 31.2 }, rodovia: 'Ecovias do Araguaia', regiao: 'BR' },
  { nome: 'Jaú', lat: -22.4083, lon: -48.5642, tarifas: { 2: 14.0 }, rodovia: 'Arteris ViaPaulista', regiao: 'BR' },
  { nome: 'João Monlevade', lat: -19.854, lon: -43.1327, tarifas: { 2: 25.8 }, rodovia: 'P02', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Jundiaí', lat: -23.0769, lon: -46.8413, tarifas: { 2: 12.2 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Juquiá', lat: -24.3893, lon: -47.7224, tarifas: { 2: 8.6 }, rodovia: 'Autopista Régis', regiao: 'BR' },
  { nome: 'Leopoldina', lat: -21.6387, lon: -42.7104, tarifas: { 2: 29.0 }, rodovia: 'EcoVias RioMinas', regiao: 'BR' },
  { nome: 'Louveira', lat: -23.0552, lon: -46.8942, tarifas: { 2: 7.8 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Luz', lat: -19.7883, lon: -45.5846, tarifas: { 2: 20.8 }, rodovia: 'Way-262', regiao: 'BR' },
  { nome: 'Magé', lat: -22.6073, lon: -43.0287, tarifas: { 2: 40.4 }, rodovia: 'EcoVias RioMinas', regiao: 'BR' },
  { nome: 'Mairiporã', lat: -23.3313, lon: -46.5775, tarifas: { 2: 6.13 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'Marginal Tietê', lat: -23.5261, lon: -46.5883, tarifas: { 2: 2.1 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Maria Dirce', lat: -23.4353, lon: -46.427, tarifas: { 2: 3.08 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Miracatu', lat: -24.1927, lon: -47.3385, tarifas: { 2: 9.46 }, rodovia: 'Autopista Régis', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Mococa', lat: -21.6408, lon: -47.0479, tarifas: { 2: 17.8 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'Mogi Mirim', lat: -22.4567, lon: -46.9022, tarifas: { 2: 22.2 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Mogi das Cruzes', lat: -23.4453, lon: -46.25, tarifas: { 2: 3.98 }, rodovia: 'CNL Novo Litoral', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Monte Alegre de Minas', lat: -18.8817, lon: -49.0352, tarifas: { 2: 11.8 }, rodovia: 'Ecovias do Cerrado', regiao: 'BR' },
  { nome: 'Monte Alto', lat: -21.2495, lon: -48.5567, tarifas: { 2: 15.0 }, rodovia: 'EcoNoroeste', regiao: 'BR' },
  { nome: 'Monte Mor', lat: -22.9781, lon: -47.3514, tarifas: { 2: 20.2 }, rodovia: 'Rodovias do Tietê', regiao: 'BR' },
  { nome: 'Moreira César', lat: -22.9302, lon: -45.3609, tarifas: { 2: 33.8 }, rodovia: 'CCR RioSP', regiao: 'BR' },
  { nome: 'Muzambinho', lat: -21.3584, lon: -46.5005, tarifas: { 2: 30.0 }, rodovia: 'EPR Vias do Café', regiao: 'BR' },
  { nome: 'Nova Friburgo', lat: -22.2135, lon: -42.4917, tarifas: { 2: 25.6 }, rodovia: 'Rota 116', regiao: 'BR' },
  { nome: 'Nova Ponte', lat: -19.0239, lon: -47.5789, tarifas: { 2: 28.0 }, rodovia: 'EPR Triângulo', regiao: 'BR' },
  { nome: 'Osasco', lat: -23.512, lon: -46.8014, tarifas: { 2: 8.4 }, rodovia: 'Ecovias Raposo Castelo', regiao: 'BR' },
  { nome: 'Ouro Fino', lat: -22.2784, lon: -46.4877, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul de Minas', regiao: 'BR' },
  { nome: 'P10 - Padroeira Externa', lat: -23.552, lon: -46.8204, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P11 - Raposo Tavares Interna', lat: -23.595, lon: -46.8106, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P12 - Raposo Tavares Externa', lat: -23.5884, lon: -46.8098, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P13 - Regis Bittencourt Externa', lat: -23.6, lon: -46.8127, tarifas: { 2: 7.0 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P13 - Régis Bittencourt Externa', lat: -23.602, lon: -46.813, tarifas: { 2: 7.0 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P2 - Bandeirantes (interna)', lat: -23.4395, lon: -46.7583, tarifas: { 2: 5.6 }, rodovia: 'CCR RodoAnel', regiao: 'BR' },
  { nome: 'P3 - Bandeirantes (externa)', lat: -23.427, lon: -46.7602, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P3 - Santa Rita do Sapucaí', lat: -22.2548, lon: -45.8225, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul de Minas', regiao: 'BR' },
  { nome: 'P4 - Anhanguera Interna Sul', lat: -23.4562, lon: -46.7853, tarifas: { 2: 5.6 }, rodovia: 'CCR RodoAnel', regiao: 'BR' },
  { nome: 'P5 - Anhanguera Interna Norte', lat: -23.4523, lon: -46.7863, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P6 - Anhanguera externa', lat: -23.448, lon: -46.7827, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P7 - Castello Branco Interna', lat: -23.518, lon: -46.8137, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P8 - Castello Branco (externa)', lat: -23.5081, lon: -46.8232, tarifas: { 2: 7.0 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'P9 - Padroeira Interna', lat: -23.5645, lon: -46.8192, tarifas: { 2: 5.6 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'PP 04 - Monte Santo de Minas', lat: -21.254, lon: -46.9503, tarifas: { 2: 30.0 }, rodovia: 'EPR Vias do Café', regiao: 'BR' },
  { nome: 'PP02 - Perdizes', lat: -19.4049, lon: -47.3148, tarifas: { 2: 28.0 }, rodovia: 'EPR Triangulo', regiao: 'BR' },
  { nome: 'Paraibuna', lat: -23.5504, lon: -45.5249, tarifas: { 2: 24.6 }, rodovia: 'Tamoios', regiao: 'BR' },
  { nome: 'Parelheiros - P1', lat: -23.7811, lon: -46.7629, tarifas: { 2: 10.8 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Patrocínio', lat: -19.0281, lon: -47.2567, tarifas: { 2: 28.0 }, rodovia: 'EPR Triângulo', regiao: 'BR' },
  { nome: 'Paulínia A', lat: -22.6902, lon: -47.155, tarifas: { 2: 24.0 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Paulínia B - Sentido Externo', lat: -22.713, lon: -47.1427, tarifas: { 2: 33.4 }, rodovia: 'Rota das Bandeiras', regiao: 'BR' },
  { nome: 'Pedro do Rio', lat: -22.2838, lon: -43.1203, tarifas: { 2: 42.0 }, rodovia: 'Concer', regiao: 'BR' },
  { nome: 'Pedágio Batistini', lat: -23.7506, lon: -46.5959, tarifas: { 2: 16.8 }, rodovia: 'Ecovias', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Pedágio Caieiras (sentido Capital/Sul)', lat: -23.3472, lon: -46.8133, tarifas: { 2: 27.4 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Campo Limpo (sentido interior)', lat: -23.3229, lon: -46.8232, tarifas: { 2: 27.4 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Ecovias do Araguaia Corumbá de Goiás P9 SAU17', lat: -16.0035, lon: -48.8412, tarifas: { 2: 28.6 }, rodovia: 'Ecovias do Araguaia', regiao: 'BR' },
  { nome: 'Pedágio Itupeva (sentido Norte)', lat: -23.0576, lon: -47.0442, tarifas: { 2: 27.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Itupeva (sentido Sul)', lat: -23.0581, lon: -47.0448, tarifas: { 2: 27.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Ituverava (sentido Norte)', lat: -20.3802, lon: -47.8134, tarifas: { 2: 34.6 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Pedágio Ituverava (sentido Sul)', lat: -20.3802, lon: -47.8137, tarifas: { 2: 34.6 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Pedágio Leme (sentido Norte)', lat: -22.2498, lon: -47.3899, tarifas: { 2: 22.4 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Pedágio Leme (sentido Sul)', lat: -22.2499, lon: -47.3902, tarifas: { 2: 22.4 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Pedágio Limeira (sentido Norte)', lat: -22.5405, lon: -47.4394, tarifas: { 2: 18.4 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Limeira (sentido Sul)', lat: -22.5101, lon: -47.3977, tarifas: { 2: 18.4 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Municipal Limeira (sentido Norte)', lat: -22.5304, lon: -47.4338, tarifas: { 2: 12.2 }, rodovia: 'Prefeitura de Limeira', regiao: 'BR' },
  { nome: 'Pedágio Municipal Limeira (sentido Sul)', lat: -22.5304, lon: -47.4338, tarifas: { 2: 12.2 }, rodovia: 'Prefeitura de Limeira', regiao: 'BR' },
  { nome: 'Pedágio Nova Odessa (sentido Norte)', lat: -22.7707, lon: -47.2387, tarifas: { 2: 24.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Nova Odessa (sentido Sul)', lat: -22.771, lon: -47.2391, tarifas: { 2: 24.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Perus (sentido Norte)', lat: -23.4182, lon: -46.7989, tarifas: { 2: 27.4 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Perus (sentido Sul)', lat: -23.4188, lon: -46.7993, tarifas: { 2: 26.7 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Pirassununga (sentido Norte)', lat: -21.9579, lon: -47.4593, tarifas: { 2: 22.4 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Pedágio Pirassununga (sentido Sul)', lat: -21.958, lon: -47.4597, tarifas: { 2: 22.4 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Pedágio Sales Oliveira (sentido Norte)', lat: -20.8509, lon: -47.8954, tarifas: { 2: 26.8 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Pedágio Sales Oliveira (sentido Sul)', lat: -20.8509, lon: -47.8957, tarifas: { 2: 28.6 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Pedágio Santa Rita do Passa Quatro (sentido Norte)', lat: -21.6584, lon: -47.6088, tarifas: { 2: 20.0 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Pedágio Santa Rita do Passa Quatro (sentido Sul)', lat: -21.6584, lon: -47.6092, tarifas: { 2: 20.0 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Pedágio Sumaré (sentido Norte)', lat: -22.8573, lon: -47.3005, tarifas: { 2: 24.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Sumaré (sentido Sul)', lat: -22.8578, lon: -47.3006, tarifas: { 2: 24.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio São Simão (sentido Norte)', lat: -21.4143, lon: -47.6641, tarifas: { 2: 19.0 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Pedágio São Simão (sentido Sul)', lat: -21.4146, lon: -47.6644, tarifas: { 2: 19.0 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Pedágio Valinhos (sentido Norte)', lat: -23.0129, lon: -47.0231, tarifas: { 2: 27.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio Valinhos (sentido Sul)', lat: -23.0203, lon: -47.0187, tarifas: { 2: 27.2 }, rodovia: 'CCR AutoBAn', regiao: 'BR' },
  { nome: 'Pedágio free flow', lat: -23.4933, lon: -46.5461, tarifas: { 2: 2.1 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Perdizes', lat: -19.6176, lon: -47.3465, tarifas: { 2: 20.6 }, rodovia: 'Way-262', regiao: 'BR' },
  { nome: 'Piracicaba', lat: -22.6065, lon: -47.7149, tarifas: { 2: 14.8 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'Pirangi - P10', lat: -21.0654, lon: -48.7086, tarifas: { 2: 25.6 }, rodovia: 'EcoVias Noroeste', regiao: 'BR' },
  { nome: 'Piratininga', lat: -23.8208, lon: -46.5827, tarifas: { 2: 77.4 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Pitangueiras', lat: -21.0093, lon: -48.1767, tarifas: { 2: 23.2 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Ponte Rio Niterói', lat: -22.8779, lon: -43.1157, tarifas: { 2: 13.2 }, rodovia: 'EcoVias Ponte', regiao: 'BR' },
  { nome: 'Porto Feliz', lat: -23.2275, lon: -47.5626, tarifas: { 2: 22.0 }, rodovia: 'AB Colinas', regiao: 'BR' },
  { nome: 'Poços de Caldas', lat: -21.9079, lon: -46.5782, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul de Minas', regiao: 'BR' },
  { nome: 'Prata', lat: -19.4715, lon: -48.8711, tarifas: { 2: 26.4 }, rodovia: 'Way-153', regiao: 'BR' },
  { nome: 'Praça 1 / Perus', lat: -23.4168, lon: -46.7375, tarifas: { 2: 6.5 }, rodovia: 'CCR Rodoanel Oeste', regiao: 'BR' },
  { nome: 'Praça P1 – Caldas – km 40,5', lat: -21.9825, lon: -46.3301, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul MG', regiao: 'BR' },
  { nome: 'Praça P2 – Caldas – km 40,5', lat: -21.9825, lon: -46.3303, tarifas: { 2: 20.2 }, rodovia: 'EPR Sul MG', regiao: 'BR' },
  { nome: 'Quadra', lat: -23.2458, lon: -48.086, tarifas: { 2: 38.8 }, rodovia: 'CCR SPVias', regiao: 'BR' },
  { nome: 'Rafard', lat: -23.0508, lon: -47.5754, tarifas: { 2: 14.4 }, rodovia: 'Rodovias do Tietê', regiao: 'BR' },
  { nome: 'Restinga', lat: -20.714, lon: -47.5061, tarifas: { 2: 23.4 }, rodovia: 'Via Paulista', regiao: 'BR' },
  { nome: 'Riacho Grande', lat: -23.7894, lon: -46.5197, tarifas: { 2: 77.4 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Ribeirão Pires - P6 externa', lat: -23.7161, lon: -46.4643, tarifas: { 2: 10.8 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Ribeirão Pires - P6 interna', lat: -23.7157, lon: -46.4644, tarifas: { 2: 8.2 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Ribeirão Pires - P7 externa', lat: -23.717, lon: -46.4633, tarifas: { 2: 11.4 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Ribeirão Pires - P7 interna', lat: -23.7144, lon: -46.4595, tarifas: { 2: 8.2 }, rodovia: 'SPMar', regiao: 'BR' },
  { nome: 'Rio Bonito', lat: -22.6869, lon: -42.5402, tarifas: { 2: 15.0 }, rodovia: 'Autopista Fluminense', regiao: 'BR' },
  { nome: 'Rio Claro', lat: -22.4302, lon: -47.5618, tarifas: { 2: 16.93 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'Rio Grande', lat: -22.8959, lon: -43.3992, tarifas: { 2: 17.2 }, rodovia: 'ViaRio', regiao: 'BR' },
  { nome: 'Rio das Pedras', lat: -22.8779, lon: -47.6344, tarifas: { 2: 25.2 }, rodovia: 'Rodovias do Tietê', regiao: 'BR' },
  { nome: 'Rod. Domingos Innocentini', lat: -22.1769, lon: -47.8821, tarifas: { 2: 10.0 }, rodovia: 'Prefeitura de Itirapina', regiao: 'BR' },
  { nome: 'Salto', lat: -23.1347, lon: -47.3607, tarifas: { 2: 9.8 }, rodovia: 'Rodovias do Tietê', regiao: 'BR' },
  { nome: 'Santa Cruz das Palmeiras', lat: -21.808, lon: -47.1934, tarifas: { 2: 18.0 }, rodovia: 'Intervias', regiao: 'BR' },
  { nome: 'Santo Antônio do Amparo', lat: -21.0003, lon: -44.9667, tarifas: { 2: 6.4 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'Santos', lat: -23.8882, lon: -46.211, tarifas: { 2: 11.6 }, rodovia: 'CNL', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Santos sentido Capital', lat: -23.8949, lon: -46.3022, tarifas: { 2: 36.6 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Senador José Bento - P2', lat: -22.1593, lon: -46.107, tarifas: { 2: 20.2 }, rodovia: 'EPG Sul de Minas', regiao: 'BR' },
  { nome: 'Sentido Sorocaba/Interior', lat: -23.4115, lon: -47.3418, tarifas: { 2: 15.0 }, rodovia: 'CCR Sorocabana', regiao: 'BR' },
  { nome: 'Seropédica', lat: -22.7151, lon: -43.7308, tarifas: { 2: 30.4 }, rodovia: 'Ecovias Rio Minas', regiao: 'BR' },
  { nome: 'Sertãozinho 1 - Entrevias', lat: -21.1668, lon: -47.9145, tarifas: { 2: 18.2 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Sertãozinho 2 - Entrevias', lat: -21.1672, lon: -47.9145, tarifas: { 2: 18.2 }, rodovia: 'Entrevias', regiao: 'BR' },
  { nome: 'Simão Pereira', lat: -21.9282, lon: -43.3163, tarifas: { 2: 42.0 }, rodovia: 'Concer', regiao: 'BR' },
  { nome: 'Sorocaba', lat: -23.5152, lon: -47.3155, tarifas: { 2: 9.9 }, rodovia: 'CCR Sorocabama', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'São Carlos', lat: -21.8214, lon: -47.9077, tarifas: { 2: 17.2 }, rodovia: 'Centrovias Arteris', regiao: 'BR' },
  { nome: 'São Gonçalo - Sentido Rio de Janeiro', lat: -22.7752, lon: -42.9455, tarifas: { 2: 15.0 }, rodovia: 'Autopista Fluminense', regiao: 'BR' },
  { nome: 'São Gonçalo do Sapucaí', lat: -21.9703, lon: -45.6309, tarifas: { 2: 6.4 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'São José dos Campos', lat: -23.2834, lon: -45.8606, tarifas: { 2: 10.8 }, rodovia: 'Ecopistas', regiao: 'BR' },
  { nome: 'São João da Boa Vista - Sentido interior', lat: -21.9482, lon: -46.8415, tarifas: { 2: 14.8 }, rodovia: 'Renovias', regiao: 'BR' },
  { nome: 'São Lourenço da Serra', lat: -23.7877, lon: -46.9139, tarifas: { 2: 8.6 }, rodovia: 'Autopista Régis', regiao: 'BR' },
  { nome: 'São Pedro I', lat: -22.6482, lon: -47.8072, tarifas: { 2: 16.6 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'São Pedro II', lat: -22.5649, lon: -48.062, tarifas: { 2: 17.2 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'São Roque', lat: -23.5899, lon: -47.0603, tarifas: { 2: 10.1 }, rodovia: 'CCR Sorocabana', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'São Sebastião do Paraíso', lat: -20.8429, lon: -46.878, tarifas: { 2: 18.2 }, rodovia: 'MG-050', regiao: 'BR' },
  { nome: 'São Vicente - Sentido Capital', lat: -23.9394, lon: -46.4637, tarifas: { 2: 20.58 }, rodovia: 'Ecovias', regiao: 'BR' },
  { nome: 'Taiúva - P5', lat: -21.1264, lon: -48.3899, tarifas: { 2: 20.4 }, rodovia: 'EcoVias Noroeste', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Tatui/Morro do Alto (sentido norte)', lat: -23.4347, lon: -47.9381, tarifas: { 2: 31.8 }, rodovia: 'CCR SPVias', regiao: 'BR' },
  { nome: 'Tatui/Morro do Alto (sentido sul)', lat: -23.4709, lon: -47.968, tarifas: { 2: 30.0 }, rodovia: 'CCR SPVias', regiao: 'BR' },
  { nome: 'Tiradentes Sentido SP', lat: -23.4844, lon: -46.5392, tarifas: { 2: 10.88 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Torrinha', lat: -22.4051, lon: -48.2617, tarifas: { 2: 15.2 }, rodovia: 'Eixo SP', regiao: 'BR' },
  { nome: 'Uberaba', lat: -19.2033, lon: -47.8507, tarifas: { 2: 28.0 }, rodovia: 'EPR Triângulo', regiao: 'BR' },
  { nome: 'Uberaba - Norte', lat: -19.1807, lon: -48.1573, tarifas: { 2: 15.8 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Uberaba - Sul', lat: -19.1808, lon: -48.1575, tarifas: { 2: 15.8 }, rodovia: 'ECO 050', regiao: 'BR' },
  { nome: 'Uberlândia', lat: -18.8764, lon: -48.5196, tarifas: { 2: 11.8 }, rodovia: 'Ecovias do Cerrado', regiao: 'BR' },
  { nome: 'Vargem', lat: -22.9088, lon: -46.4246, tarifas: { 2: 6.4 }, rodovia: 'Autopista Fernão Dias', regiao: 'BR' },
  { nome: 'ViaRio', lat: -22.9195, lon: -43.3965, tarifas: { 2: 17.2 }, rodovia: 'ViaRio', regiao: 'BR' },
  { nome: 'Vicinal Graciano da Ressurreição Affonso - Sentido Araraquara', lat: -21.6672, lon: -48.2475, tarifas: { 2: 8.0 }, rodovia: 'Prefeitura de Araraquara', regiao: 'BR' },
  { nome: 'Vila Maria', lat: -23.5096, lon: -46.5735, tarifas: { 2: 2.1 }, rodovia: 'CCR RioSP', regiao: 'BR', modelo: 'free_flow' },
  { nome: 'Vinhedo', lat: -23.0567, lon: -47.0413, tarifas: { 2: 16.8 }, rodovia: 'DERSA', regiao: 'BR' },
  { nome: 'Xerém', lat: -22.6102, lon: -43.2857, tarifas: { 2: 42.0 }, rodovia: 'Concer', regiao: 'BR' },
  { nome: 'Água Santa', lat: -22.9073, lon: -43.3089, tarifas: { 2: 6.0 }, rodovia: 'LAMSA', regiao: 'BR' },
  { nome: 'Águas da Prata', lat: -21.9306, lon: -46.6963, tarifas: { 2: 12.6 }, rodovia: 'Renovias', regiao: 'BR' },
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
// cada 150m — fino o bastante mesmo com o raio de detecção reduzido (500m),
// pra uma amostragem grossa não "pular" por cima do ponto exato onde a reta
// passa mais perto do pedágio.
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
// Raio de detecção configurável via `raioKm` (padrão 500m, ajustado a partir
// de 300m — em uso real 300m estava deixando passar pedágios legítimos: a
// coordenada cadastrada em PEDAGIOS_BR nem sempre cai exatamente em cima do
// ponto por onde a polyline do OSRM passa, e 300m tinha pouca margem pra essa
// pequena imprecisão). Com o traçado REAL vindo do OSRM (ver
// obterPontosRotaComCoords / _mvRoutePoints), um raio pequeno é o certo: se
// o trajeto de verdade passa pelo pedágio, os pontos da polyline ficam a
// poucos metros dele; se o trajeto vai por outra via (mesmo perto), a
// distância mínima sobe bem acima de 500m e o pedágio não é acusado à toa.
// IMPORTANTE: 500m só faz sentido quando `paradas` é o traçado real (polyline
// do OSRM). Para chamadas que usam só a LINHA RETA entre pontos (fallback
// antes do trajeto real carregar, ou relatórios que nunca buscam o trajeto
// real, como o relatório de custo de frete), 500m ainda é curto demais — a
// reta entre origem/destino se afasta facilmente mais que isso do ponto onde
// a estrada real (que curva) passa perto do pedágio, e o pedágio deixa de
// ser detectado. Quem só tem a linha reta deve passar um `raioKm` maior
// (ex.: 3) explicitamente — ver chamada em nexta-frota-roteirizador.js
// (relatório de frete), que já faz isso.
// ATENÇÃO: se um pedágio real deixar de ser detectado com o traçado REAL, o
// mais provável é a coordenada cadastrada em PEDAGIOS_BR estar um pouco
// deslocada da posição exata da praça/pórtico (a base atual vem do
// OpenStreetMap — ver comentário no topo do arquivo, em PEDAGIOS_BR) — ajuste
// lat/lon da entrada em vez de aumentar esse raio de volta.
function detectarPedagiosNaRota(paradas, eixos = 2, raioKm = 0.5) {
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
