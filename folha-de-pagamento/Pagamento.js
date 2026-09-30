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
  // A aba "Formularios" da gaveta, com o contador ao lado. O primeiro
  // seletor veio da marcacao real; os outros sao alternativas.
  abaFormularios:
    'div.tw-flex:has(palantir-badge):text-matches("^\s*Formul", "i"), ' +
    '[role="tab"]:has-text("Formul"), ' +
    'palantir-tab:has-text("Formul")',
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

/**
 * Localiza a linha do identificador e abre a O.S. pelo botao de edicao.
 *
 * A O.S. abre como uma GAVETA por cima da listagem — mesma aba, mesma URL.
 * A listagem continua no DOM atras dela, entao "ainda estou na listagem" nao
 * serve como sinal de falha: o que confirma a abertura e a barra de abas da
 * gaveta (Geral, Formularios, Vinculos...).
 */
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

  // Um aviso do Field por cima engole o clique sem dar erro.
  await fecharPopups(page);

  const botaoEditar = linha.locator(SELETORES.botaoEditarLinha).last();
  const alvo = (await botaoEditar.count().catch(() => 0)) ? botaoEditar : linha;

  for (let tentativa = 1; tentativa <= 2; tentativa += 1) {
    await alvo.click({ timeout: 5000 }).catch(async () => {
      await alvo.click({ force: true }).catch(() => {});
    });

    if (await esperarGaveta(page)) return page;
    await page.waitForTimeout(1500);
  }

  const caminho = await salvarDiagnostico(page, `nao-abriu-${identificador}`);
  throw new Error(
    `O clique em Editar nao abriu a O.S. ${identificador}` +
      (caminho ? ` | Diagnostico: ${caminho}.png` : ""),
  );
}

/** A gaveta da O.S. esta aberta? Confirma pela aba "Formularios". */
async function esperarGaveta(page, timeout = 12000) {
  return page
    .locator(SELETORES.abaFormularios)
    .first()
    .waitFor({ state: "visible", timeout })
    .then(() => true)
    .catch(() => false);
}

/**
 * Separa o numero impresso ao lado da pergunta.
 * "10. NECESSARIO TROCA DO EQUIPAMENTO *" -> { numero: 10, pergunta: "..." }
 */
function separarNumeroPergunta(texto) {
  const limpo = String(texto || "")
    .replace(/\s+/g, " ")
    .replace(/\s*\*\s*$/, "")
    .trim();
  const casamento = limpo.match(/^(\d{1,3})\s*[).:\-–]?\s+(.+)$/);
  if (!casamento) return { numero: null, pergunta: limpo };
  return { numero: Number(casamento[1]), pergunta: casamento[2].trim() };
}

/**
 * Carrega todas as respostas do formulario.
 *
 * O Field mostra so as primeiras e deixa um "Carregar mais" no fim. Sem
 * clicar ate o fim, perguntas de numero alto (34, 42, 45) nunca chegam ao
 * DOM — e sairiam da planilha como se nao existissem.
 */
async function carregarTodasAsRespostas(page, limite = 30) {
  for (let volta = 0; volta < limite; volta += 1) {
    const botao = page
      .locator("button, palantir-button")
      .filter({ hasText: /^\s*Carregar mais\s*$/i })
      .first();

    if (!(await botao.isVisible().catch(() => false))) return volta;

    await botao.click({ timeout: 5000 }).catch(async () => {
      await botao.click({ force: true }).catch(() => {});
    });
    await page.waitForTimeout(1200);
  }
  return limite;
}

/**
 * Abre a aba "Formularios" da gaveta e extrai as respostas.
 *
 * Retorna [{ numero, pergunta, resposta, marcado, comentario, fotos }]. O
 * numero e o impresso ao lado da pergunta, que e o que o mapeamento por tipo
 * de O.S. usa.
 */
async function extrairFormularios(page, identificador) {
  const aba = page.locator(SELETORES.abaFormularios).first();
  if (await aba.isVisible().catch(() => false)) {
    await aba.click({ force: true }).catch(() => {});
    await page.waitForTimeout(2500);
  }

  await carregarTodasAsRespostas(page);
  await page.waitForTimeout(800);

  fs.mkdirSync(DEBUG_DIR, { recursive: true });
  fs.writeFileSync(
    path.join(DEBUG_DIR, `${identificador.replace(/[\\/:*?"<>|]/g, "_")}.html`),
    await page.content(),
    "utf8",
  );

  const itens = await page.evaluate(() => {
    const limpar = (valor) =>
      (valor || "")
        .replace(/\s+/g, " ")
        // "SIM (0 pts.)" -> "SIM"
        .replace(/\(\s*-?\d+(?:[.,]\d+)?\s*pts?\.?\s*\)/gi, "")
        .trim();

    // O enunciado vem numerado: "10. NECESSARIO TROCA DO EQUIPAMENTO *".
    const ENUNCIADO = /^\s*\d{1,3}\s*[).:\-–]?\s+\S/;

    // Acha o cartao da pergunta: sobe do enunciado ate o elemento que tambem
    // contem os campos de resposta. Nao depende do nome das classes, que o
    // Field troca sem aviso.
    function blocoDaPergunta(enunciado) {
      let atual = enunciado.parentElement;
      for (let nivel = 0; nivel < 6 && atual; nivel += 1) {
        const temCampos = atual.querySelector(
          'input, textarea, select, palantir-checkbox, palantir-radio, [role="radio"], [role="checkbox"], img',
        );
        if (temCampos) return atual;
        atual = atual.parentElement;
      }
      return enunciado.parentElement || enunciado;
    }

    const vistos = new Set();
    const resultado = [];

    const candidatos = Array.from(
      document.querySelectorAll("p, h1, h2, h3, h4, h5, span, div, legend, label"),
    );

    for (const no of candidatos) {
      const texto = (no.textContent || "").replace(/\s+/g, " ").trim();
      if (!ENUNCIADO.test(texto) || texto.length > 200) continue;
      // So o elemento que carrega o texto, nao os ancestrais nem os campos.
      if (no.querySelector("input, textarea, select, img")) continue;

      const bloco = blocoDaPergunta(no);
      if (vistos.has(bloco)) continue;
      vistos.add(bloco);

      // 1. Opcoes marcadas (radio SIM/NAO, checkbox de multipla escolha).
      const marcadas = [];
      for (const campo of bloco.querySelectorAll(
        'input[type="radio"], input[type="checkbox"]',
      )) {
        if (!campo.checked) continue;
        const rotulo =
          campo.closest("label") ||
          (campo.id && bloco.querySelector(`label[for="${campo.id}"]`)) ||
          campo.parentElement;
        marcadas.push(limpar(rotulo ? rotulo.textContent : campo.value));
      }
      for (const campo of bloco.querySelectorAll(
        '[aria-selected="true"], [aria-checked="true"]',
      )) {
        const botao = campo.closest("button") || campo.parentElement;
        const rotulo = botao && botao.querySelector("span");
        marcadas.push(
          limpar(rotulo ? rotulo.textContent : botao && botao.textContent),
        );
      }

      // 2. Texto escrito (observacao, comentario, valor digitado).
      const escritos = [];
      for (const campo of bloco.querySelectorAll("textarea, input")) {
        if (campo.type === "radio" || campo.type === "checkbox") continue;
        if (campo.value && campo.value.trim()) escritos.push(limpar(campo.value));
      }

      // 3. Fotos: a resposta e a propria imagem; registra quantas.
      const fotos = bloco.querySelectorAll('img:not([src^="data:image/svg"])')
        .length;

      // Todas as opcoes oferecidas, para conferir o que ficou sem marcar.
      const opcoes = Array.from(
        bloco.querySelectorAll("label, span.tw-break-all"),
      )
        .map((e) => limpar(e.textContent))
        .filter((t) => t && t.length < 80);

      const unicas = (lista) => lista.filter((v, i) => v && lista.indexOf(v) === i);

      resultado.push({
        enunciado: texto,
        marcadas: unicas(marcadas),
        comentario: unicas(escritos).join(" | "),
        fotos,
        opcoes: unicas(opcoes).join(" | "),
      });
    }

    return resultado;
  });

  return itens.map((item, indice) => {
    const { numero, pergunta } = separarNumeroPergunta(item.enunciado);
    const marcado = item.marcadas.join(" | ");

    // A resposta util e o que foi marcado; sem marcacao, o que foi escrito;
    // sem nenhum dos dois, a contagem de fotos.
    const resposta =
      marcado || item.comentario || (item.fotos ? `${item.fotos} foto(s)` : "");

    return {
      numero,
      pergunta,
      resposta,
      marcado,
      comentario: item.comentario,
      fotos: item.fotos,
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
          const pagina = await abrirOs(page, os.identificador);
          const formularios = await extrairFormularios(pagina, os.identificador);

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
            // Uma pergunta pode ter marcacao E comentario ("NAO" mais a
            // observacao do tecnico); os dois importam para o lancamento.
            linha[rotulo] = item
              ? [item.marcado, item.comentario].filter(Boolean).join(" — ") ||
                item.resposta
              : "";
          }

          // Guarda tudo que veio do formulario, inclusive o que nao foi mapeado.
          for (const item of formularios) {
            respostasCompletas.push({
              Identificador: os.identificador,
              Tipo: os.tipo,
              Categoria: os.categoria,
              Numero: item.numero ?? "",
              Ordem: item.ordem,
              Pergunta: item.pergunta,
              Resposta: item.resposta,
              Marcado: item.marcado,
              Comentario: item.comentario,
              Fotos: item.fotos || "",
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
  // Exportados para poder testar a leitura sem abrir o Field.
  extrairFormularios,
  separarNumeroPergunta,
  carregarTodasAsRespostas,
  MAPA_PERGUNTAS,
  SAIDA_PADRAO,
};
