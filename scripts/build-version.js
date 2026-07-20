#!/usr/bin/env node
// ═══════════════════════════════════════════════════════════════════════════
// NEXTA — build-version.js
// ═══════════════════════════════════════════════════════════════════════════
// Roda automaticamente a cada deploy (ver vercel.json → buildCommand).
// Calcula um hash de 8 caracteres do CONTEÚDO de cada .js/.css referenciado
// em assets/ no index.html, e reescreve a query string "?v=..." de cada um
// com o hash correspondente.
//
// Por quê: antes disso, a versão era uma string fixa editada à mão
// ("?v=kmreal-fix-2") — se alguém esquecesse de trocar depois de editar um
// arquivo, o navegador (ou o CDN da Vercel) podia continuar servindo a
// versão antiga em cache mesmo com o conteúdo novo já publicado. Isso já
// causou confusão real numa sessão de debug (uma correção enviada não
// rodava no navegador do usuário por causa de cache).
//
// Com hash de conteúdo: só muda a URL quando o arquivo realmente muda —
// arquivo idêntico = mesmo hash = cache continua válido (bom pra
// performance); arquivo diferente = hash diferente = navegador busca a
// versão nova automaticamente, sem precisar lembrar de nada.
// ═══════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.resolve(__dirname, '..');
const INDEX_HTML = path.join(ROOT, 'index.html');

function hashFile(filePath) {
  const conteudo = fs.readFileSync(filePath);
  return crypto.createHash('sha256').update(conteudo).digest('hex').slice(0, 8);
}

function main() {
  let html = fs.readFileSync(INDEX_HTML, 'utf8');
  let trocas = 0;

  // Casa qualquer referência a assets/<algo>.js ou .css, com ou sem query
  // string existente — funciona tanto pra primeira vez rodando (sem "?v=")
  // quanto pras próximas (troca o hash antigo pelo novo).
  const regex = /(assets\/[\w.\-]+\.(?:js|css))(\?v=[\w.\-]+)?/g;

  html = html.replace(regex, (match, caminhoRelativo) => {
    const caminhoAbsoluto = path.join(ROOT, caminhoRelativo);
    if (!fs.existsSync(caminhoAbsoluto)) {
      console.warn(`[build-version] aviso: ${caminhoRelativo} referenciado no index.html mas não encontrado em disco — deixei sem alterar.`);
      return match;
    }
    const hash = hashFile(caminhoAbsoluto);
    trocas++;
    return `${caminhoRelativo}?v=${hash}`;
  });

  fs.writeFileSync(INDEX_HTML, html, 'utf8');
  console.log(`[build-version] ${trocas} referência(s) de assets atualizada(s) com hash de conteúdo em ${path.relative(ROOT, INDEX_HTML)}.`);
}

main();
