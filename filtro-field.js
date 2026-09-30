// Filtro "Identificador" da listagem de atividades do Field.
// Fica em um modulo separado porque o fechamento e a automacao de placas
// usam a mesma tela — quando o Field muda esse filtro, so este arquivo muda.

const fs = require("node:fs");
const path = require("node:path");
const { fecharPopups } = require("./popups-field.js");

const DEBUG_DIR = path.join(__dirname, "debug-field");

const SELETORES = {
  maisFiltros: "button.palantir-filter-bar__more",
  chip: "palantir-filter-chip",
  chipTrigger: "button.palantir-filter-chip__trigger",
  item: "palantir-filter-item, .filter-item",
  itemLabel: ".filter-item__label",
  // O placeholder ja foi "Buscar..."; hoje e "Busque aqui...".
  busca: 'input[placeholder="Busque aqui..."], input[placeholder="Buscar..."]',
};

/** Salva tela e HTML para descobrir o que o Field mostrou no momento da falha. */
async function salvarDiagnostico(page, nome) {
  try {
    fs.mkdirSync(DEBUG_DIR, { recursive: true });
    const base = path.join(DEBUG_DIR, nome.replace(/[\\/:*?"<>|]/g, "_"));
    await page.screenshot({ path: `${base}.png`, fullPage: true });
    fs.writeFileSync(`${base}.html`, await page.content(), "utf8");
    return base;
  } catch {
    return null;
  }
}

/**
 * Abre o menu "Mais filtros".
 *
 * O menu é um overlay do CDK: ele só existe no DOM depois que o botão é
 * acionado, e o botão alterna (um segundo clique fecharia). Por isso a
 * confirmação é o aria-expanded do botão mais a presença dos itens.
 */
async function abrirMenuFiltros(page) {
  const botao = page.locator(SELETORES.maisFiltros).first();

  const visivel = await botao
    .waitFor({ state: "visible", timeout: 20000 })
    .then(() => true)
    .catch(() => false);

  if (!visivel) {
    const caminho = await salvarDiagnostico(page, "sem-botao-mais-filtros");
    throw new Error(
      `Botao "Mais filtros" nao encontrado. URL: ${page.url()}` +
        (caminho ? ` | Diagnostico: ${caminho}.png` : ""),
    );
  }

  const itens = page.locator(SELETORES.item);

  for (let tentativa = 1; tentativa <= 3; tentativa += 1) {
    // Se ja estiver aberto, nao clica de novo: o clique fecharia o menu.
    if (await itens.first().isVisible().catch(() => false)) return itens;

    // Um modal do Field ("Boas-vindas", avisos) cobre a tela e engole o
    // clique sem dar erro — o menu simplesmente nao abre.
    await fecharPopups(page);

    // Clique normal primeiro (espera o elemento ficar acionavel de verdade);
    // force so como ultimo recurso.
    await botao.click({ timeout: 5000 }).catch(async () => {
      await botao.click({ force: true }).catch(() => {});
    });

    const abriu = await itens
      .first()
      .waitFor({ state: "visible", timeout: 6000 })
      .then(() => true)
      .catch(() => false);

    if (abriu) return itens;

    // O clique pode ter alternado o menu para fechado; reabre na proxima volta.
    await page.waitForTimeout(800);
  }

  const expandido = await botao.getAttribute("aria-expanded").catch(() => null);
  const caminho = await salvarDiagnostico(page, "menu-filtros-nao-abriu");
  throw new Error(
    `Menu "Mais filtros" nao abriu apos 3 tentativas ` +
      `(aria-expanded=${expandido}). URL: ${page.url()}` +
      (caminho ? ` | Diagnostico: ${caminho}.png` : ""),
  );
}

/**
 * Deixa o campo de busca do filtro "Identificador" visivel e retorna o locator.
 */
async function abrirFiltroIdentificador(page) {
  const busca = page.locator(SELETORES.busca).first();

  const buscaVisivel = (timeout = 6000) =>
    busca
      .waitFor({ state: "visible", timeout })
      .then(() => true)
      .catch(() => false);

  // 1. Painel ja aberto da iteracao anterior.
  if (await busca.isVisible().catch(() => false)) return busca;

  // 2. Filtro fixado como chip na barra (botao de pin do Field).
  const chip = page
    .locator(SELETORES.chip)
    .filter({ hasText: /^\s*Identificador\s*$/ })
    .first();
  if (await chip.count().catch(() => 0)) {
    await chip
      .locator(SELETORES.chipTrigger)
      .click({ force: true })
      .catch(() => {});
    if (await buscaVisivel(4000)) return busca;
  }

  // 3. Caminho completo: "Mais filtros" -> item "Identificador".
  await abrirMenuFiltros(page);

  // Localiza pelo rotulo exato e sobe para o container clicavel.
  // "Identificador do cliente" nao pode ser escolhido por engano.
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
    const caminho = await salvarDiagnostico(page, "sem-filtro-identificador");
    throw new Error(
      `Filtro "Identificador" nao encontrado. ` +
        `Disponiveis: ${disponiveis.map((t) => t.trim()).join(" | ") || "nenhum"}` +
        (caminho ? ` | Diagnostico: ${caminho}.png` : ""),
    );
  }

  // O clicavel e o div.filter-item (role="button"), nao o span do rotulo.
  const item = rotulo
    .locator('xpath=ancestor::*[contains(@class,"filter-item")][last()]')
    .first();

  const alvo = (await item.count().catch(() => 0)) ? item : rotulo;
  await alvo.click({ force: true });
  await page.waitForTimeout(1200);

  if (!(await buscaVisivel(8000))) {
    const caminho = await salvarDiagnostico(page, "sem-campo-busca");
    throw new Error(
      "Campo de busca do filtro Identificador nao apareceu depois do clique." +
        (caminho ? ` | Diagnostico: ${caminho}.png` : ""),
    );
  }

  return busca;
}

/**
 * Filtra a listagem por um identificador (ou placa) e espera a tabela atualizar.
 */
async function filtrarPorIdentificador(page, valor) {
  const busca = await abrirFiltroIdentificador(page);

  await busca.click({ force: true });
  await busca.fill("");
  await busca.fill(valor);
  await page.waitForTimeout(2500);

  // Alguns filtros listam as opcoes encontradas para marcar; se for o caso,
  // seleciona a que corresponde ao valor buscado.
  const opcao = page
    .locator('palantir-checkbox, [role="option"], palantir-list-item')
    .filter({ hasText: valor })
    .first();

  if (await opcao.count().catch(() => 0)) {
    await opcao.click({ force: true }).catch(() => {});
    await page.waitForTimeout(1200);
    // Fecha o painel para liberar a tabela.
    await page.keyboard.press("Escape").catch(() => {});
  }

  await page.waitForTimeout(1500);
}

module.exports = {
  filtrarPorIdentificador,
  abrirFiltroIdentificador,
  salvarDiagnostico,
  SELETORES,
  DEBUG_DIR,
};
