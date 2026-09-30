/**
 * Pagamento.js - Fechamento da folha de pagamento.
 *
 * Le a planilha de O.S. exportada do Field (colocada nesta mesma pasta),
 * busca cada Identificador no filtro do Field, abre a O.S. pelo botao de
 * edicao e extrai as respostas dos formularios conforme o tipo de servico.
 *
 * Uso:
 *   node folha-de-pagamento/Pagamento.js [planilha.xlsx] [saida.xlsx]
 *
 * Credenciais: EMAIL e PASSWORD no arquivo .env (nunca no codigo).
 */

const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");
const XLSX = require("xlsx");
const { carregarCredenciais, carregarUrlsField } = require("../variaveis.js");
const { filtrarPorIdentificador, salvarDiagnostico } = require("../filtro-field.js");
const { fecharPopups } = require("../popups-field.js");

// ---------------------------------------------------------------------------
// Configuracao
// ---------------------------------------------------------------------------

// Tudo do pagamento fica nesta pasta: planilha de entrada, saida e debug.
const BASE_DIR = __dirname;
const DEBUG_DIR = path.join(BASE_DIR, "debug-formularios");
const SAIDA_PADRAO = path.join(BASE_DIR, "pagamento_formularios.xlsx");

// Numeros das perguntas de formulario retornadas por tipo de O.S.
const PERGUNTAS_INSTALACAO = [11, 15, 34, 45];
const PERGUNTAS_REMOCAO = [13, 31, 34, 37, 39, 40, 42];

const MAPA_PERGUNTAS = {
  instalacao: PERGUNTAS_INSTALACAO,
  remocao: PERGUNTAS_REMOCAO,
  // Remanejamento = instalacao + remocao. Regra definitiva ainda a definir;
  // por enquanto exporta a uniao dos dois conjuntos.
  remanejamento: [
    ...new Set([...PERGUNTAS_INSTALACAO, ...PERGUNTAS_REMOCAO]),
  ].sort((a, b) => a - b),
  // Manutencao com avaria / violacao / extravio.
  manutencao: PERGUNTAS_REMOCAO,
};

// Seletores do Field (marcacao Palantir).
const SELETORES = {
  botaoEditarLinha: "palantir-button.palantir-table__action-button",
  abaFormularios: 'div.tw-flex:has(palantir-badge):text-matches("Formul", "i")',
};


const HEADLESS = process.env.HEADLESS !== "false";
const TIMEOUT_LOGIN = 30000;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Remove acentos e normaliza para comparacao. */
function normalizar(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .trim();
}

/**
 * Classifica a O.S. a partir do Tipo e do Status da planilha.
 * Retorna a chave do MAPA_PERGUNTAS, "impedido" ou null (fora do escopo).
 */
function classificarOs({ tipo, status }) {
  if (normalizar(status).includes("impedido")) return "impedido";

  const alvo = normalizar(tipo);
  if (alvo.includes("remanejamento")) return "remanejamento";
  if (alvo.includes("instalacao")) return "instalacao";
  if (alvo.includes("remocao")) return "remocao";
  if (
    alvo.includes("manutencao") ||
    alvo.includes("avaria") ||
    alvo.includes("violacao") ||
    alvo.includes("extravio")
  ) {
    return "manutencao";
  }
  // Preparacao de materiais, revisao preventiva, laudo etc. nao entram.
  return null;
}

/**
 * Le a planilha exportada do Field. O cabecalho nao fica na primeira linha,
 * entao procuramos a linha que contem a coluna "Identificador".
 */
function lerOrdensServico(caminhoPlanilha) {
  const workbook = XLSX.readFile(caminhoPlanilha);
  const aba = workbook.Sheets[workbook.SheetNames[0]];
  const linhas = XLSX.utils.sheet_to_json(aba, { header: 1, defval: "" });

  const indiceCabecalho = linhas.findIndex((linha) =>
    linha.some((celula) => normalizar(celula) === "identificador"),
  );
  if (indiceCabecalho === -1) {
    throw new Error('Coluna "Identificador" nao encontrada na planilha.');
  }

  const cabecalho = linhas[indiceCabecalho].map((celula) => normalizar(celula));
  const coluna = (nome) => cabecalho.indexOf(normalizar(nome));

  const colunas = {
    identificador: coluna("Identificador"),
    tipo: coluna("Tipo"),
    cliente: coluna("Cliente"),
    status: coluna("Status"),
    tecnico: coluna("Técnico"),
    avaria: coluna("Avaria"),
  };

  return linhas
    .slice(indiceCabecalho + 1)
    .map((linha) => ({
      identificador: String(linha[colunas.identificador] ?? "").trim(),
      tipo: String(linha[colunas.tipo] ?? "").trim(),
      cliente: String(linha[colunas.cliente] ?? "").trim(),
      status: String(linha[colunas.status] ?? "").trim(),
      tecnico: String(linha[colunas.tecnico] ?? "").trim(),
      avaria: String(linha[colunas.avaria] ?? "").trim(),
    }))
    .filter((os) => os.identificador);
}

/** Sem argumento, pega a planilha mais recente desta pasta. */
function descobrirPlanilhaEntrada() {
  return (
    fs
      .readdirSync(BASE_DIR)
      .filter(
        (nome) =>
          /\.(xlsx|xlsm|csv)$/i.test(nome) &&
          !nome.startsWith("~$") &&
          path.join(BASE_DIR, nome) !== SAIDA_PADRAO,
      )
      .map((nome) => path.join(BASE_DIR, nome))
      .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0] || null
  );
}

// ---------------------------------------------------------------------------
// Automacao
// ---------------------------------------------------------------------------

async function login(page) {
  const { email, senha } = carregarCredenciais();

  await page.goto(carregarUrlsField().login);

  await page.click('input[name="email"]');
  await page.fill('input[name="email"]', email);
  await page.click('palantir-button:has-text("Continuar")');

  await page.click('input[aria-label="password"]');
  await page.fill('input[aria-label="password"]', senha);
  await page.click('palantir-button:has-text("Continuar")');

  await page.waitForTimeout(TIMEOUT_LOGIN);
  await fecharPopups(page);
  await page.waitForTimeout(3000);
}

/**
 * Garante que a tela esteja na listagem antes da proxima O.S.
 *
 * Nao usar page.goBack(): quando a O.S. nao chega a abrir, o historico volta
 * para as telas do login e a aba termina em about:blank, derrubando todas as
 * O.S. seguintes em cascata.
 */
async function voltarParaListagem(page) {
  const URL_ATIVIDADES = carregarUrlsField().listagem;
  // Fecha um painel ou modal que tenha ficado aberto.
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(500);

  if (!page.url().startsWith(URL_ATIVIDADES)) {
    await page
      .goto(URL_ATIVIDADES, { waitUntil: "domcontentloaded" })
      .catch(() => {});
    await page.waitForTimeout(3000);
    return;
  }

  await page.waitForTimeout(1500);
}

/** Localiza a linha do identificador na tabela e clica no botao de edicao. */
async function abrirOs(page, identificador) {
  await filtrarPorIdentificador(page, identificador);

  const linha = page
    .locator('tr, [role="row"], palantir-table-row')
    .filter({ hasText: identificador })
    .first();

  const encontrou = await linha
    .waitFor({ state: "visible", timeout: 20000 })
    .then(() => true)
    .catch(() => false);

  if (!encontrou) {
    throw new Error(`Identificador ${identificador} nao encontrado na listagem.`);
  }

  // O botao de edicao fica na coluna "Editar" da propria linha.
  const botaoEditar = linha.locator(SELETORES.botaoEditarLinha).last();
  const temBotao = await botaoEditar.count().catch(() => 0);

  // O Field abre a O.S. em uma aba nova. Sem escutar esse evento, a automacao
  // continuaria lendo a aba da listagem — que foi o que aconteceu antes:
  // todas as paginas salvas ainda eram a lista, com zero formularios.
  const novaAba = page
    .context()
    .waitForEvent("page", { timeout: 10000 })
    .catch(() => null);

  const urlAntes = page.url();

  if (temBotao) {
    await botaoEditar.click({ force: true });
  } else {
    // Fallback: abre pela propria linha quando a coluna de acoes nao aparece.
    await linha.click({ force: true });
  }

  const aba = await novaAba;
  if (aba) {
    await aba.waitForLoadState("domcontentloaded").catch(() => {});
    await aba.waitForTimeout(4000);
    await fecharPopups(aba);
    return { pagina: aba, abaNova: true };
  }

  await page.waitForTimeout(4000);

  // Sem aba nova e sem mudar de endereco: o clique nao abriu nada.
  if (page.url() === urlAntes) {
    const aindaNaListagem = await page
      .locator("task-list-filter-bar")
      .first()
      .isVisible()
      .catch(() => false);

    if (aindaNaListagem) {
      const caminho = await salvarDiagnostico(
        page,
        `nao-abriu-${identificador}`,
      );
      throw new Error(
        `O clique em Editar nao abriu a O.S. ${identificador}` +
          (caminho ? ` | Diagnostico: ${caminho}.png` : ""),
      );
    }
  }

  return { pagina: page, abaNova: false };
}

/**
 * Separa o numero impresso ao lado da pergunta.
 * "11. Placa instalada" -> { numero: 11, pergunta: "Placa instalada" }
 * Sem numero impresso, numero fica null.
 */
function separarNumeroPergunta(texto) {
  const casamento = String(texto || "").match(/^\s*(\d{1,3})\s*[).:\-–]?\s+(.+)$/);
  if (!casamento) return { numero: null, pergunta: String(texto || "").trim() };
  return { numero: Number(casamento[1]), pergunta: casamento[2].trim() };
}

/**
 * Abre o campo "Formularios" (a aba com o badge de quantidade, a direita)
 * e extrai todos os pares pergunta/resposta, usando o numero impresso
 * ao lado de cada pergunta.
 */
async function extrairFormularios(page, identificador) {
  const aba = page
    .locator("div.tw-flex.tw-items-center")
    .filter({ hasText: /^\s*Formul[aá]rios/i })
    .first();

  const temAba = await aba
    .waitFor({ state: "visible", timeout: 15000 })
    .then(() => true)
    .catch(() => false);
  if (temAba) {
    await aba.click({ force: true }).catch(() => {});
    await page.waitForTimeout(2500);
  }

  // Expande todos os acordeoes de formulario visiveis.
  const acordeoes = page.locator("palantir-accordion:visible");
  const totalAcordeoes = await acordeoes.count().catch(() => 0);
  for (let indice = 0; indice < totalAcordeoes; indice += 1) {
    await acordeoes
      .nth(indice)
      .click({ force: true })
      .catch(() => {});
    await page.waitForTimeout(300);
  }
  await page.waitForTimeout(1500);

  // Salva o HTML para calibrar os seletores caso a extracao venha vazia.
  fs.mkdirSync(DEBUG_DIR, { recursive: true });
  fs.writeFileSync(
    path.join(DEBUG_DIR, `${identificador.replace(/[\\/:*?"<>|]/g, "_")}.html`),
    await page.content(),
    "utf8",
  );

  // Cada pergunta e um componente <app-form-*-question>.
  const itens = await page.evaluate(() => {
    const limpar = (valor) =>
      (valor || "")
        .replace(/\s+/g, " ")
        // O Field imprime a pontuacao ao lado da opcao: "Cameras (0 pts.)".
        .replace(/\(\s*-?\d+(?:[.,]\d+)?\s*pts?\.?\s*\)/gi, "")
        .trim();

    const perguntas = Array.from(document.querySelectorAll("*")).filter((no) =>
      /^APP-FORM-.*-QUESTION$/.test(no.tagName),
    );

    return perguntas.map((bloco) => {
      // Enunciado: primeiro <p>, sem o asterisco de campo obrigatorio.
      const enunciado = bloco.querySelector("p");
      const pergunta = limpar(enunciado ? enunciado.textContent : "").replace(
        /\s*\*$/,
        "",
      );

      // 1) Opcoes marcadas (checkbox / radio / select).
      const marcadas = Array.from(
        bloco.querySelectorAll(
          'palantir-checkbox[aria-selected="true"], palantir-radio[aria-selected="true"], [aria-selected="true"], [aria-checked="true"]',
        ),
      )
        .map((marcada) => {
          const botao = marcada.closest("button") || marcada.parentElement;
          const rotulo = botao && botao.querySelector("span.tw-break-all");
          return limpar(rotulo ? rotulo.textContent : botao && botao.textContent);
        })
        .filter(Boolean);

      // Opcoes disponiveis, para mostrar o que existia e nao foi marcado.
      const opcoes = Array.from(bloco.querySelectorAll("span.tw-break-all"))
        .map((span) => limpar(span.textContent))
        .filter(Boolean);

      if (marcadas.length) {
        const unicas = marcadas.filter((v, i) => marcadas.indexOf(v) === i);
        return {
          pergunta,
          resposta: unicas.join(" | "),
          tipo: "selecao",
          opcoes: opcoes.join(" | "),
        };
      }

      // 2) Sem nada marcado: le o que foi escrito (observacao, texto, numero).
      const escritos = Array.from(
        bloco.querySelectorAll("input, textarea"),
      )
        .map((campo) => limpar(campo.value))
        .filter(Boolean);

      if (escritos.length) {
        return {
          pergunta,
          resposta: escritos.join(" | "),
          tipo: "texto",
          opcoes: "",
        };
      }

      // 3) Ultimo recurso: texto do bloco sem o enunciado (valores ja salvos).
      const corpo = limpar(bloco.innerText || bloco.textContent)
        .replace(pergunta, "")
        .replace(/^\s*\*\s*/, "")
        .trim();

      return {
        pergunta,
        resposta: opcoes.length ? "" : corpo,
        tipo: opcoes.length ? "selecao" : "texto",
        opcoes: opcoes.join(" | "),
      };
    });
  });

  return itens.map((item, indice) => {
    const { numero, pergunta } = separarNumeroPergunta(item.pergunta);
    return {
      // Quando o enunciado nao traz o numero impresso, cai na posicao.
      numero: numero ?? indice + 1,
      numeroImpresso: numero,
      pergunta,
      resposta: item.resposta,
      tipoCampo: item.tipo,
      opcoes: item.opcoes,
      ordem: indice + 1,
    };
  });
}

// ---------------------------------------------------------------------------
// Orquestracao
// ---------------------------------------------------------------------------

/**
 * Executa o fechamento.
 * onEvento recebe o progresso ({ tipo, ... }) para o front acompanhar.
 */
async function runPagamento(
  ordens,
  saidaPath = SAIDA_PADRAO,
  onEvento = () => {},
  deveParar = () => false,
) {
  // Marca se a execucao foi interrompida pelo usuario.
  let interrompida = false;
  const emitir = (evento) => {
    try {
      onEvento(evento);
    } catch {
      // O progresso nunca pode derrubar a automacao.
    }
  };

  const credenciais = carregarCredenciais();
  if (!credenciais.email || !credenciais.senha) {
    throw new Error(
      "Credenciais nao configuradas. Informe o login do Field na tela inicial.",
    );
  }

  const resultados = [];
  // Conteudo bruto dos formularios, para o usuario avaliar o que entra na folha.
  const respostasCompletas = [];

  // O.S. impedidas e tipos fora do escopo nao precisam abrir o navegador.
  const paraVisitar = [];
  for (const os of ordens) {
    const categoria = classificarOs(os);
    if (categoria === "impedido") {
      resultados.push({ ...montarLinhaBase(os, "impedido"), Observacao: "O.S. impedida" });
      continue;
    }
    if (!categoria) {
      resultados.push({
        ...montarLinhaBase(os, "fora do escopo"),
        Observacao: "Tipo sem formularios de pagamento",
      });
      continue;
    }
    paraVisitar.push({ ...os, categoria });
  }

  emitir({
    tipo: "inicio",
    total: ordens.length,
    aVisitar: paraVisitar.length,
    ignoradas: resultados.length,
  });

  if (paraVisitar.length) {
    const browser = await chromium.launch({ headless: HEADLESS });
    const page = await browser.newPage();

    try {
      emitir({ tipo: "login" });
      await login(page);
      emitir({ tipo: "logado" });

      let processadas = 0;
      for (const os of paraVisitar) {
        // Parada solicitada: encerra aqui, mas o que ja foi lido vai para a
        // planilha na sequencia — nao se perde o trabalho das O.S. anteriores.
        if (deveParar()) {
          interrompida = true;
          emitir({
            tipo: "parado",
            processadas,
            restantes: paraVisitar.length - processadas,
          });
          break;
        }

        processadas += 1;
        emitir({
          tipo: "os-inicio",
          identificador: os.identificador,
          categoria: os.categoria,
          atual: processadas,
          total: paraVisitar.length,
        });
        try {
          const { pagina, abaNova } = await abrirOs(page, os.identificador);

          let formularios;
          try {
            formularios = await extrairFormularios(pagina, os.identificador);
          } finally {
            // A aba da O.S. nao pode ficar acumulando entre as 213 iteracoes.
            if (abaNova) await pagina.close().catch(() => {});
          }

          // Zero respostas quase sempre significa seletor errado, nao O.S.
          // vazia — antes isso era reportado como "OK" e passava batido.
          if (!formularios.length) {
            throw new Error(
              "A O.S. abriu, mas nenhuma resposta de formulario foi lida " +
                "(conferir os seletores no HTML salvo em debug-formularios/)",
            );
          }

          const numerosDesejados = MAPA_PERGUNTAS[os.categoria];

          const linha = montarLinhaBase(os, os.categoria);
          linha.Observacao = formularios.length
            ? ""
            : "Nenhuma resposta de formulario lida";

          for (const numero of numerosDesejados) {
            const item = formularios.find((form) => form.numero === numero);
            const rotulo = `F${numero} - ${item ? item.pergunta : "nao encontrada"}`;
            linha[rotulo] = item ? item.resposta : "";
          }

          // Guarda tudo que veio do formulario, inclusive o que nao foi mapeado.
          for (const item of formularios) {
            respostasCompletas.push({
              Identificador: os.identificador,
              Tipo: os.tipo,
              Categoria: os.categoria,
              Numero: item.numero,
              "Numero impresso": item.numeroImpresso ?? "",
              Ordem: item.ordem,
              Pergunta: item.pergunta,
              Resposta: item.resposta,
              "Tipo do campo": item.tipoCampo,
              "Opcoes disponiveis": item.opcoes,
              "Entra na folha": numerosDesejados.includes(item.numero)
                ? "Sim"
                : "Avaliar",
            });
          }

          resultados.push(linha);
          console.log(
            `${os.identificador} [${os.categoria}] ${formularios.length} respostas lidas, ` +
              `${numerosDesejados.length} exportadas`,
          );
          emitir({
            tipo: "os-ok",
            identificador: os.identificador,
            categoria: os.categoria,
            respostas: formularios.length,
            exportadas: numerosDesejados.length,
          });
        } catch (erro) {
          console.error(`Falha em ${os.identificador}: ${erro.message}`);
          resultados.push({
            ...montarLinhaBase(os, os.categoria),
            Observacao: `ERRO: ${erro.message}`,
          });
          emitir({
            tipo: "os-erro",
            identificador: os.identificador,
            categoria: os.categoria,
            erro: erro.message,
          });
        }

        await voltarParaListagem(page);
      }
    } finally {
      await browser.close();
    }
  }

  const workbook = XLSX.utils.book_new();
  // Aba 1: uma linha por O.S., so com as perguntas mapeadas.
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(resultados),
    "Pagamento",
  );
  // Aba 2: tudo que veio dos formularios, para avaliacao manual.
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(respostasCompletas),
    "Formularios",
  );
  XLSX.writeFile(workbook, saidaPath);
  console.log(
    `Planilha gerada: ${saidaPath} ` +
      `(${resultados.length} O.S., ${respostasCompletas.length} respostas)`,
  );
  emitir({
    tipo: "fim",
    outputPath: saidaPath,
    ordens: resultados.length,
    respostas: respostasCompletas.length,
    interrompida,
  });

  return { outputPath: saidaPath, resultados, respostasCompletas, interrompida };
}

/** Colunas fixas presentes em toda linha da planilha de saida. */
function montarLinhaBase(os, categoria) {
  return {
    Identificador: os.identificador,
    Tipo: os.tipo,
    Categoria: categoria,
    Cliente: os.cliente,
    Status: os.status,
    Tecnico: os.tecnico,
    Avaria: os.avaria,
  };
}

async function main() {
  const entrada = process.argv[2] || descobrirPlanilhaEntrada();
  if (!entrada) {
    console.error(
      `Nenhuma planilha encontrada em ${BASE_DIR}.\n` +
        "Coloque o arquivo la ou rode: node folha-de-pagamento/Pagamento.js <planilha.xlsx>",
    );
    process.exit(1);
  }

  const saida = process.argv[3] ? path.resolve(process.argv[3]) : SAIDA_PADRAO;
  const ordens = lerOrdensServico(path.resolve(entrada));

  if (!ordens.length) {
    console.error("Nenhuma O.S. encontrada na planilha de entrada.");
    process.exit(1);
  }

  console.log(`${ordens.length} O.S. lidas de ${path.basename(entrada)}`);
  await runPagamento(ordens, saida);
}

if (require.main === module) {
  main().catch((erro) => {
    console.error("Erro na execucao:", erro);
    process.exit(1);
  });
}

module.exports = {
  runPagamento,
  lerOrdensServico,
  classificarOs,
  descobrirPlanilhaEntrada,
  MAPA_PERGUNTAS,
  SAIDA_PADRAO,
};
