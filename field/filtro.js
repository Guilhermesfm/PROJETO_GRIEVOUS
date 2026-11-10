// Filtro "Identificador" da listagem de atividades do Field.
// Usado pelo fechamento e pela automacao de placas.

const { fecharPopups, clicar } = require("./popups.js");

const SELETORES = {
  maisFiltros: "button.palantir-filter-bar__more",
  chip: "palantir-filter-chip",
  chipTrigger: "button.palantir-filter-chip__trigger",
  item: "palantir-filter-item, .filter-item",
  itemLabel: ".filter-item__label",
  // O placeholder varia entre versoes do Field.
  busca: 'input[placeholder="Busque aqui..."], input[placeholder="Buscar..."]',
  // Painel do filtro (overlay do CDK).
  painel: ".cdk-overlay-container",
};

/**
 * Retorna o campo de busca do painel de filtro, ou null.
 * Restrito ao painel: a pesquisa global tem o mesmo placeholder.
 */
async function acharCampoBusca(page, timeout = 6000) {
  const campo = page.locator(`${SELETORES.painel} ${SELETORES.busca}`).first();

  const visivel = await campo
    .waitFor({ state: "visible", timeout })
    .then(() => true)
    .catch(() => false);

  return visivel ? campo : null;
}

/** Espera a listagem terminar de carregar (a barra "Mais filtros" aparecer). */
async function esperarListagemPronta(
  page,
  tentativas = 3,
  porTentativa = 45000,
) {
  for (let tentativa = 1; tentativa <= tentativas; tentativa += 1) {
    const limite = Date.now() + porTentativa;

    while (Date.now() < limite) {
      if (
        await page
          .locator(SELETORES.maisFiltros)
          .first()
          .isVisible()
          .catch(() => false)
      ) {
        return true;
      }

      await fecharPopups(page, 2);
      await page.waitForTimeout(1500);
    }

    // A lista pode travar no esqueleto de carregamento: recarrega a pagina.
    if (tentativa < tentativas) {
      await page.reload({ waitUntil: "domcontentloaded" }).catch(() => {});
      await page.waitForTimeout(4000);
    }
  }

  return false;
}

/** Abre o menu "Mais filtros". */
async function abrirMenuFiltros(page) {
  const botao = page.locator(SELETORES.maisFiltros).first();

  const visivel = await botao
    .waitFor({ state: "visible", timeout: 20000 })
    .then(() => true)
    .catch(() => false);

  if (!visivel) {
    throw new Error(`Botao "Mais filtros" nao encontrado. URL: ${page.url()}`);
  }

  const itens = page.locator(SELETORES.item);

  for (let tentativa = 1; tentativa <= 3; tentativa += 1) {
    // Se ja estiver aberto, nao clica de novo: o clique fecharia o menu.
    if (
      await itens
        .first()
        .isVisible()
        .catch(() => false)
    )
      return itens;

    await fecharPopups(page);

    await clicar(botao);

    const abriu = await itens
      .first()
      .waitFor({ state: "visible", timeout: 6000 })
      .then(() => true)
      .catch(() => false);

    if (abriu) return itens;

    // O clique pode ter fechado o menu; tenta de novo na proxima volta.
    await page.waitForTimeout(800);
  }

  const expandido = await botao.getAttribute("aria-expanded").catch(() => null);
  throw new Error(
    `Menu "Mais filtros" nao abriu apos 3 tentativas ` +
      `(aria-expanded=${expandido}). URL: ${page.url()}`,
  );
}

/**
 * Deixa o campo de busca do filtro "Identificador" visivel e retorna o locator.
 */
async function abrirFiltroIdentificador(page) {
  // 1. Painel ja aberto da iteracao anterior.
  const jaAberto = await acharCampoBusca(page, 400);
  if (jaAberto) return jaAberto;

  // 2. A listagem precisa estar carregada.
  await esperarListagemPronta(page);

  // 3. Filtro fixado como chip na barra (botao de pin do Field).
  const chip = page
    .locator(SELETORES.chip)
    .filter({ hasText: /^\s*Identificador\s*$/ })
    .first();
  if (await chip.count().catch(() => 0)) {
    await chip
      .locator(SELETORES.chipTrigger)
      .click({ force: true })
      .catch(() => {});
    const campo = await acharCampoBusca(page, 4000);
    if (campo) return campo;
  }

  // 4. Caminho completo: "Mais filtros" -> item "Identificador".
  await abrirMenuFiltros(page);

  // Rotulo exato (evita "Identificador do cliente"), depois o container clicavel.
  const rotulo = page
    .locator(SELETORES.itemLabel)
    .filter({ hasText: /^\s*Identificador\s*$/ })
    .first();

  const temRotulo = await rotulo
    .waitFor({ state: "visible", timeout: 10000 })
    .then(() => true)
    .catch(() => false);

  if (!temRotulo) {
    const disponiveis = await page
      .locator(SELETORES.itemLabel)
      .allInnerTexts()
      .catch(() => []);
    throw new Error(
      `Filtro "Identificador" nao encontrado. ` +
        `Disponiveis: ${disponiveis.map((t) => t.trim()).join(" | ") || "nenhum"}`,
    );
  }

  const item = rotulo
    .locator('xpath=ancestor::*[contains(@class,"filter-item")][last()]')
    .first();

  const alvo = (await item.count().catch(() => 0)) ? item : rotulo;
  await alvo.click({ force: true });
  await page.waitForTimeout(1200);

  const campo = await acharCampoBusca(page, 8000);
  if (!campo) {
    throw new Error(
      "Campo de busca do filtro Identificador nao apareceu depois do clique.",
    );
  }

  return campo;
}

/**
 * Filtra a listagem por um identificador (ou placa) e espera a tabela atualizar.
 */
async function filtrarPorIdentificador(page, valor) {
  const busca = await abrirFiltroIdentificador(page);

  await busca.click({ force: true });
  await busca.fill("");
  await busca.fill(valor);

  await esperarTabelaFiltrada(page, valor);

  // Se o filtro listar opcoes para marcar, seleciona a do valor buscado.
  const opcao = page
    .locator('palantir-checkbox, [role="option"], palantir-list-item')
    .filter({ hasText: valor })
    .first();

  if (await opcao.isVisible().catch(() => false)) {
    await opcao.click({ force: true }).catch(() => {});
    await page.keyboard.press("Escape").catch(() => {});
    await esperarTabelaFiltrada(page, valor);
  }
}

/** Espera a linha do valor buscado aparecer na tabela. */
async function esperarTabelaFiltrada(page, valor, timeout = 15000) {
  const linha = page
    .locator('tr, [role="row"], palantir-table-row')
    .filter({ hasText: valor })
    .first();

  return linha
    .waitFor({ state: "visible", timeout })
    .then(() => true)
    .catch(() => false);
}

module.exports = {
  filtrarPorIdentificador,
  esperarListagemPronta,
  SELETORES,
};
