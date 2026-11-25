// Abre e fecha a gaveta da O.S. a partir da linha ja filtrada na listagem.
//
// A gaveta abre por cima da listagem, sem mudar a URL: a confirmacao de que
// abriu e a barra de abas dela (Geral, Formularios...).

const { fecharPopups, clicar } = require("./popups.js");

const SELETOR_EDITAR = "palantir-button.palantir-table__action-button";

/** A aba "Formularios" da gaveta, usada como sinal de que ela esta aberta. */
function localizarAbaFormularios(page) {
  return page
    .locator("button.palantir-tabs__button")
    .filter({ hasText: /Formul[aá]rios/i })
    .first();
}

/** A gaveta esta aberta? */
async function gavetaAberta(page, timeout = 12000) {
  return localizarAbaFormularios(page)
    .waitFor({ state: "visible", timeout })
    .then(() => true)
    .catch(() => false);
}

/** Escapa o texto para usar literalmente dentro de uma RegExp. */
function escaparRegex(texto) {
  return String(texto).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Abre a gaveta clicando no botao de edicao da linha.
 * Pressupoe que a listagem ja esta filtrada pelo identificador.*/
async function abrirGaveta(page, identificador) {
  const linhas = page.locator('tr, [role="row"], palantir-table-row');
  const qualquerLinha = linhas.filter({ hasText: identificador }).first();

  const achou = await qualquerLinha
    .waitFor({ state: "visible", timeout: 5000 })
    .then(() => true)
    .catch(() => false);
  if (!achou) {
    throw new Error(`Identificador ${identificador} nao encontrado na listagem.`);
  }

  // Linha com uma celula igual ao identificador; sem ela, a que contem o texto.
  const exato = new RegExp(`^\\s*${escaparRegex(identificador)}\\s*$`, "i");
  const linhaExata = linhas
    .filter({ has: page.locator('td, [role="cell"]', { hasText: exato }) })
    .first();
  const linha = (await linhaExata.count().catch(() => 0)) ? linhaExata : qualquerLinha;

  // Um aviso do Field por cima engole o clique sem dar erro.
  await fecharPopups(page);

  // Lapis de edicao; sem ele, clica na propria linha.
  const botao = linha.locator(SELETOR_EDITAR).last();
  const alvo = (await botao.count().catch(() => 0)) ? botao : linha;

  // Duas tentativas de abrir a gaveta.
  for (let tentativa = 1; tentativa <= 2; tentativa += 1) {
    await clicar(alvo);
    if (await gavetaAberta(page, 15000)) return;
    await page.waitForTimeout(1500);
  }

  throw new Error(`O clique em Editar nao abriu a O.S. ${identificador}`);
}

/** Fecha a gaveta e espera a listagem voltar a ser a tela de frente. */
async function fecharGaveta(page) {
  await page.keyboard.press("Escape").catch(() => {});
  await localizarAbaFormularios(page)
    .waitFor({ state: "hidden", timeout: 5000 })
    .catch(() => {});
}

module.exports = {
  abrirGaveta,
  fecharGaveta,
  localizarAbaFormularios,
};
