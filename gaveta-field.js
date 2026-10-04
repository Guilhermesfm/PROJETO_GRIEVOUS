// Abre e fecha a gaveta da O.S. a partir da linha ja filtrada na listagem.
//
// A O.S. abre como uma gaveta por cima da listagem: mesma aba, mesma URL.
// A listagem continua no DOM atras dela, entao a confirmacao de que abriu e
// a barra de abas da gaveta (Geral, Formularios, Vinculos...), nao a URL.

const { fecharPopups } = require("./popups-field.js");

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

/**
 * Abre a gaveta clicando no botao de edicao da linha.
 * Pressupoe que a listagem ja esta filtrada pelo identificador.
 */
async function abrirGaveta(page, identificador) {
  const linha = page
    .locator('tr, [role="row"], palantir-table-row')
    .filter({ hasText: identificador })
    .first();

  const achou = await linha
    .waitFor({ state: "visible", timeout: 20000 })
    .then(() => true)
    .catch(() => false);

  if (!achou) return false;

  // Um aviso do Field por cima engole o clique sem dar erro.
  await fecharPopups(page);

  const botao = linha.locator(SELETOR_EDITAR).last();
  const alvo = (await botao.count().catch(() => 0)) ? botao : linha;

  for (let tentativa = 1; tentativa <= 2; tentativa += 1) {
    // Clique normal primeiro: espera o elemento ficar acionavel de verdade.
    await alvo.click({ timeout: 8000 }).catch(async () => {
      await alvo.click({ force: true }).catch(() => {});
    });

    if (await gavetaAberta(page)) return true;
    await page.waitForTimeout(1500);
  }

  return false;
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
  gavetaAberta,
  localizarAbaFormularios,
  SELETOR_EDITAR,
};
