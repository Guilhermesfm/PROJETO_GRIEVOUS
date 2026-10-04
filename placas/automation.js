const path = require("node:path");
const { chromium } = require("playwright");
const XLSX = require("xlsx");
const { carregarCredenciais, carregarUrlsField } = require("../variaveis.js");
const { filtrarPorIdentificador } = require("../filtro-field.js");
const { fecharPopups } = require("../popups-field.js");
const { capturarLinkOs } = require("../link-field.js");
const { abrirGaveta, fecharGaveta } = require("../gaveta-field.js");

// Tudo da automacao de placas fica nesta pasta.
const OUTPUT_DIR = __dirname;

function resolveOutputPath(filename = "situacoes_placas.xlsx") {
  return path.join(OUTPUT_DIR, filename);
}

async function runAutomation(placas = [], deveParar = () => false) {
  // Marca se a execucao foi interrompida pelo usuario.
  let interrompida = false;
  const { email, senha } = carregarCredenciais();
  if (!email || !senha) {
    throw new Error(
      "Credenciais nao configuradas. Informe o login do Field na tela inicial.",
    );
  }

  const resultados = [];

  // Sem executablePath: o Playwright acha o Chromium que ele mesmo instalou,
  // em qualquer maquina. HEADLESS=false abre o navegador para acompanhar.
  const browser = await chromium.launch({
    headless: process.env.HEADLESS !== "false",
  });
  const page = await browser.newPage();

  try {
    // 1. Acessa a página de login
    await page.goto(carregarUrlsField().login);

    // 2. Clica no campo de e-mail e preenche
    await page.click('input[name="email"]');
    await page.fill('input[name="email"]', email);

    // 3. Clica em "Continuar"
    await page.click('palantir-button:has-text("Continuar")');

    // 4. Clica no campo de senha e preenche
    await page.click('input[aria-label="password"]');
    await page.fill('input[aria-label="password"]', senha);

    // 5. Clica em "Continuar" de novo
    await page.click('palantir-button:has-text("Continuar")');

    // 6. Aguarda 30 segundos para o login finalizar
    await page.waitForTimeout(30000);

    // 7. Aguarda e fecha todos os pop-ups conhecidos
    await page.waitForTimeout(1000);
    await fecharPopups(page);
    await page.waitForTimeout(5000);

    // 8. Repete a busca para cada placa da lista
    for (const placa of placas) {
      // Parada solicitada: para na placa atual e salva o que ja foi consultado.
      if (deveParar()) {
        interrompida = true;
        console.log("Execucao interrompida pelo usuario.");
        break;
      }

      // Filtro "Identificador" da barra nova do Field (modulo compartilhado).
      await filtrarPorIdentificador(page, placa);

      const candidatosResultado = [
        page.getByText(placa, { exact: true }),
        page.locator(`text=${placa}`),
        page.locator("tr, li, div, palantir-item, palantir-card").filter({
          hasText: placa,
        }),
        page
          .locator(
            '[role="row"], [role="option"], [data-cy*="row"], [data-cy*="item"]',
          )
          .filter({ hasText: placa }),
      ];

      let placaSelecionada = false;
      for (const candidato of candidatosResultado) {
        const total = await candidato.count().catch(() => 0);
        if (!total) continue;

        const alvo = candidato.filter({ hasText: placa }).first();
        if (
          await alvo
            .waitFor({ state: "visible", timeout: 15000 })
            .then(() => true)
            .catch(() => false)
        ) {
          await alvo.click({ force: true });
          placaSelecionada = true;
          break;
        }
      }

      if (!placaSelecionada) {
        const pagina = await page.locator("body").innerText();
        throw new Error(
          `Nao foi possivel localizar a placa ${placa} na tela. ` +
            `URL: ${page.url()}\nConteudo da pagina:\n${pagina.slice(0, 2000)}`,
        );
      }

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

      // So as concluidas interessam para o link. O botao Compartilhar vive
      // dentro da gaveta da O.S., que ate aqui nao chegava a ser aberta —
      // por isso o link saia sempre vazio.
      let link = "";
      if (/conclu/i.test(situacao)) {
        if (await abrirGaveta(page, placa)) {
          link = await capturarLinkOs(page);
          await fecharGaveta(page);
        } else {
          console.log(`Nao foi possivel abrir a O.S. ${placa} para pegar o link.`);
        }
      }

      resultados.push({
        Placa: placa,
        Situacao: situacao,
        Link: link,
        "Data de criacao": dataCriacao,
      });

      console.log(`Placa ${placa} -> ${situacao} | Criada em ${dataCriacao}`);
    }
  } finally {
    await browser.close();
  }

  // 9. Gera a planilha Excel com os resultados
  const planilha = XLSX.utils.json_to_sheet(resultados);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, planilha, "Situacoes");
  const outputPath = resolveOutputPath();
  XLSX.writeFile(workbook, outputPath);

  console.log(`Planilha gerada: ${outputPath}`);

  return { outputPath, resultados, interrompida };
}

async function main() {
  await runAutomation([]);
}

if (require.main === module) {
  main().catch((error) => {
    console.error("Erro na execucao da automacao:", error);
    process.exit(1);
  });
}

module.exports = {
  resolveOutputPath,
  runAutomation,
};
