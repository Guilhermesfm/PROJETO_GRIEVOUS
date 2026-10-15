const path = require("node:path");
const { chromium } = require("playwright");
const { carregarCredenciais, carregarUrlsField } = require("../field/credenciais.js");
const {
  filtrarPorIdentificador,
  esperarListagemPronta,
} = require("../field/filtro.js");
const { fecharPopups } = require("../field/popups.js");
const { capturarLinkOs } = require("../field/link.js");
const { abrirGaveta, fecharGaveta } = require("../field/gaveta.js");

// Tudo da automacao de placas fica nesta pasta.
const { gravarPlanilha, SEM_RETORNO } = require("../planilha.js");

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

    // 6. Eram 36s de espera cega aqui. O esperarListagemPronta abaixo aguarda
    //    a tela ficar pronta de fato, e recarrega se a consulta do Field
    //    falhar; fica so o tempo de sair da tela de login.
    await page.waitForTimeout(5000);
    await fecharPopups(page);

    if (!(await esperarListagemPronta(page))) {
      throw new Error("A listagem de atividades nao carregou depois do login.");
    }

    // 8. Repete a busca para cada placa da lista
    for (const placa of placas) {
      // Parada solicitada: para na placa atual e salva o que ja foi consultado.
      if (deveParar()) {
        interrompida = true;
        console.log("Execucao interrompida pelo usuario.");
        break;
      }

      // Uma placa problematica nao pode derrubar o lote: antes qualquer
      // falha encerrava a execucao e perdia as placas seguintes.
      try {
        // Filtro "Identificador" da barra nova do Field (modulo compartilhado).
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

        // Espera a data aparecer em vez de um tempo fixo de 3s.
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

        // So as concluidas interessam para o link. O botao Compartilhar vive
        // dentro da gaveta da O.S., que ate aqui nao chegava a ser aberta.
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

  // 9. Gera a planilha Excel com os resultados
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
