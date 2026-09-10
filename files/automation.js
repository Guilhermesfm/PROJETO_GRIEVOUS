const { chromium } = require("playwright");
const XLSX = require("xlsx");

(async () => {
  // Preencha aqui as placas que serao consultadas.
  const placas = ["OPZ4G70", "QMT7J41", "TDO8D16", "RWV6H65", "GFS4C62"];

  const resultados = [];

  const browser = await chromium.launch({
    headless: false,
    executablePath: chromium.executablePath(),
  });
  const page = await browser.newPage();

  async function fecharPopups() {
    const botoesFechamento = page.locator(
      '.cdk-overlay-container palantir-button[icon="close-outline"]',
    );
    for (let tentativa = 0; tentativa < 2; tentativa += 1) {
      const botaoFechamento = botoesFechamento.last();
      if (
        await botaoFechamento
          .waitFor({ state: "visible", timeout: 10000 })
          .then(() => true)
          .catch(() => false)
      ) {
        await botaoFechamento.click({ force: true });
        await page.waitForTimeout(500);
      }
    }

    const maisTarde = page.locator('button[data-cy="fc-modal-right-button"]');
    const botaoMaisTarde = maisTarde.filter({ hasText: "Mais tarde" });
    if (
      await botaoMaisTarde
        .waitFor({ state: "visible", timeout: 10000 })
        .then(() => true)
        .catch(() => false)
    ) {
      await botaoMaisTarde.click({ force: true });
      await page.waitForTimeout(500);
    }
  }

  // 1. Acessa a página de login
  await page.goto(
    "https://app.fieldcontrol.com.br/autenticador-v2/#/login?to=:hash:%2Fatividades",
  );

  // 2. Clica no campo de e-mail e preenche
  await page.click('input[name="email"]');
  await page.fill('input[name="email"]', "guilherme.miguel@inoprime.com.br");

  // 3. Clica em "Continuar"
  await page.click('palantir-button:has-text("Continuar")');

  // 4. Clica no campo de senha e preenche
  await page.click('input[aria-label="password"]');
  await page.fill('input[aria-label="password"]', "Inoprime@1234");

  // 5. Clica em "Continuar" de novo
  await page.click('palantir-button:has-text("Continuar")');

  // 6. Aguarda 30 segundos para o login finalizar
  await page.waitForTimeout(30000);

  // 7. Aguarda e fecha todos os pop-ups conhecidos
  await page.waitForTimeout(1000);
  await fecharPopups();
  await page.waitForTimeout(5000);

  // 8. Repete a busca para cada placa da lista
  for (const placa of placas) {
    // Clica no ícone de filtro
    const filtros = page.locator(
      'palantir-button[data-cy="filter-button"]:visible',
    );
    const filtroPorIdentificador = page
      .locator('palantir-text[type="title"]')
      .filter({ hasText: /^Por identificador$/ })
      .locator("xpath=ancestor::palantir-accordion[1]");
    const campoBusca = filtroPorIdentificador.locator(
      'input[placeholder="Buscar..."]',
    );
    let filtroAberto = await campoBusca.isVisible().catch(() => false);
    const quantidadeFiltros = await filtros.count();

    if (!filtroAberto) {
      for (let indice = 0; indice < quantidadeFiltros; indice += 1) {
        const filtro = filtros.nth(indice);
        await filtro.getByText("Filtrar", { exact: true }).click({
          force: true,
        });
        await page.waitForTimeout(2500);
        if (
          await campoBusca
            .waitFor({ state: "visible", timeout: 10000 })
            .then(() => true)
            .catch(() => false)
        ) {
          filtroAberto = true;
          break;
        }
      }
    }

    if (!filtroAberto) {
      console.log(
        "Aguardando o campo Buscar...; se necessario, abra Filtrar na janela do navegador.",
      );
      filtroAberto = await campoBusca
        .waitFor({ state: "visible", timeout: 30000 })
        .then(() => true)
        .catch(() => false);
    }

    if (!filtroAberto) {
      const estruturaFiltro = quantidadeFiltros
        ? await filtros.first().evaluate((elemento) => {
            const pais = [];
            let atual = elemento;
            for (let nivel = 0; nivel < 4 && atual; nivel += 1) {
              pais.push(atual.outerHTML);
              atual = atual.parentElement;
            }
            return pais.join("\n---\n");
          })
        : "nenhum filtro encontrado";
      throw new Error(
        `Nao foi possivel abrir o filtro de placas. URL: ${page.url()} | ` +
          `Titulo: ${await page.title()} | Filtros: ${quantidadeFiltros}\n${estruturaFiltro}`,
      );
    }

    // Preenche diretamente o campo exibido dentro do acordeao de filtros.
    await campoBusca.waitFor({ state: "visible", timeout: 10000 });
    await campoBusca.click({ force: true });
    await campoBusca.fill(placa);
    await page.waitForTimeout(2000);

    // Clica no resultado que aparece com o texto da placa
    await page.locator(`text=${placa}`).first().click({ force: true });
    await page.waitForTimeout(3000);

    const dataCriacao = (
      await page.locator('palantir-text[type="text"]').allInnerTexts()
    )
      .map((texto) => texto.trim())
      .find((texto) => /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$/.test(texto));

    if (!dataCriacao) {
      throw new Error(`Data de criacao nao encontrada para a placa ${placa}`);
    }

    // Captura o texto da situação da OS
    const situacao = (
      await page
        .locator('palantir-text[type="title"][size="xs"]')
        .filter({ hasText: /\S/ })
        .first()
        .innerText()
    ).trim();

    resultados.push({
      Placa: placa,
      Situacao: situacao,
      "Data de criacao": dataCriacao,
    });

    console.log(`Placa ${placa} -> ${situacao} | Criada em ${dataCriacao}`);
  }

  await browser.close();

  // 9. Gera a planilha Excel com os resultados
  const planilha = XLSX.utils.json_to_sheet(resultados);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, planilha, "Situacoes");
  XLSX.writeFile(workbook, "situacoes_placas.xlsx");

  console.log("Planilha gerada: situacoes_placas.xlsx");
})();
