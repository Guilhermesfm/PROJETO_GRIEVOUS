/**
 * placas/automation.js - consulta de placas no Field.
 *
 * Para cada placa: filtra a listagem pelo identificador, le a situacao e a
 * data de criacao e, se a O.S. estiver concluida, pega o link de
 * compartilhamento. O resultado vai para placas/situacoes_placas.xlsx.
 */

const path = require("node:path");
const { chromium } = require("playwright");
const {
  carregarCredenciais,
  carregarUrlsField,
} = require("../field/credenciais.js");
const {
  filtrarPorIdentificador,
  esperarListagemPronta,
} = require("../field/filtro.js");
const { fecharPopups } = require("../field/popups.js");
const { capturarLinkOs } = require("../field/link.js");
const { abrirGaveta, fecharGaveta } = require("../field/gaveta.js");
const { gravarPlanilha, SEM_RETORNO } = require("../planilha.js");

// A planilha de saida fica nesta pasta.
const OUTPUT_DIR = __dirname;

/** Caminho da planilha de saida, dentro desta pasta. */
function resolveOutputPath(filename = "situacoes_placas.xlsx") {
  return path.join(OUTPUT_DIR, filename);
}

/**
 * Faz login no Field, consulta cada placa e grava a planilha.
 * deveParar() e checada antes de cada placa.
 * Retorna { outputPath, resultados, interrompida }.
 */
async function runAutomation(placas = [], deveParar = () => false) {
  let interrompida = false;
  const { email, senha } = carregarCredenciais();
  const urls = carregarUrlsField();
  const destinoListagem = new URL(urls.listagem);
  if (!email || !senha) {
    throw new Error(
      "Credenciais nao configuradas. Informe o login do Field na tela inicial.",
    );
  }

  const resultados = [];

  // HEADLESS=false abre o navegador na tela.
  const browser = await chromium.launch({
    headless: process.env.HEADLESS !== "false",
  });
  const page = await browser.newPage();

  try {
    // Login: e-mail, Continuar, senha, Continuar.
    await page.goto(urls.login);

    await page.click('input[name="email"]');
    await page.fill('input[name="email"]', email);

    await page.click('palantir-button:has-text("Continuar")');

    await page.click('input[aria-label="password"]');
    await page.fill('input[aria-label="password"]', senha);

    await page.click('palantir-button:has-text("Continuar")');

    // Espera o login redirecionar para a listagem antes de seguir.
    await page.waitForURL(
      (url) =>
        url.origin === destinoListagem.origin &&
        url.hash.startsWith(destinoListagem.hash),
      { timeout: 60000 },
    );
    await fecharPopups(page);

    if (!(await esperarListagemPronta(page))) {
      throw new Error("A listagem de atividades nao carregou depois do login.");
    }

    for (const placa of placas) {
      // Parada solicitada: salva o que ja foi consultado.
      if (deveParar()) {
        interrompida = true;
        console.log("Execucao interrompida pelo usuario.");
        break;
      }

      // Falha numa placa nao derruba o lote.
      try {
        await filtrarPorIdentificador(page, placa);

        const linha = page
          .locator('tr, [role="row"], palantir-table-row')
          .filter({ hasText: placa })
          .first();

        const achou = await linha
          .waitFor({ state: "visible", timeout: 20000 })
          .then(() => true)
          .catch(() => false);

        if (!achou) {
          throw new Error("placa nao encontrada na listagem");
        }

        await linha.click({ force: true });

        // Espera a data de criacao aparecer (ate 15s).
        const padraoData = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$/;
        const textos = page.locator('palantir-text[type="text"]');
        let dataCriacao = "";
        const limite = Date.now() + 15000;

        while (!dataCriacao && Date.now() < limite) {
          dataCriacao =
            (await textos.allInnerTexts().catch(() => []))
              .map((texto) => texto.trim())
              .find((texto) => padraoData.test(texto)) || "";
          if (!dataCriacao) await page.waitForTimeout(400);
        }

        const situacao = (
          await page
            .locator('palantir-text[type="title"][size="xs"]')
            .filter({ hasText: /\S/ })
            .first()
            .innerText()
            .catch(() => "")
        ).trim();

        // So as concluidas interessam para o link.
        let link = "";
        if (/conclu/i.test(situacao)) {
          if (await abrirGaveta(page, placa)) {
            link = await capturarLinkOs(page);
            await fecharGaveta(page);
          }
        }

        resultados.push({
          Placa: placa,
          Situacao: situacao || SEM_RETORNO,
          Link: link || (/conclu/i.test(situacao) ? SEM_RETORNO : ""),
          "Data de criacao": dataCriacao || SEM_RETORNO,
          Observacao: situacao ? "" : "Nada retornado para esta placa",
        });

        console.log(
          `Placa ${placa} -> ${situacao || SEM_RETORNO} | ${dataCriacao || SEM_RETORNO}`,
        );
      } catch (erro) {
        console.error(`Falha em ${placa}: ${erro.message}`);
        resultados.push({
          Placa: placa,
          Situacao: SEM_RETORNO,
          Link: "",
          "Data de criacao": SEM_RETORNO,
          Observacao: `ERRO: ${erro.message}`,
        });
      }
    }
  } finally {
    await browser.close();
  }

  // Gera a planilha com os resultados.
  const outputPath = resolveOutputPath();
  gravarPlanilha(outputPath, [
    {
      nome: "Situacoes",
      linhas: resultados,
      ordem: ["Placa", "Situacao", "Data de criacao", "Link", "Observacao"],
    },
  ]);

  console.log(`Planilha gerada: ${outputPath}`);

  return { outputPath, resultados, interrompida };
}

/** Execucao direta pelo terminal (as placas vem pelo Grievous.js). */
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
