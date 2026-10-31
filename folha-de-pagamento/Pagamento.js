/**
 * Pagamento.js - Fechamento da folha de pagamento.
 *
 * Le a planilha de O.S. exportada do Field (colocada nesta mesma pasta),
 * busca cada Identificador no filtro do Field, abre a O.S. e le as respostas
 * dos formularios. Gera a folha no formato usado com as empresas: Cliente,
 * Identificador, Tipo, Status, Data Execucao, Avaria, Violacao, Obs e Link.
 *
 * Regra:
 *   - Remocao / instalacao (e remanejamento, que e as duas): equipamentos
 *     removidos/instalados, dano nos equipamentos e extravio.
 *   - Manutencao: violacao (coluna Violacao) e extravio.
 *   - Formulario enviado so com fotos, sem as informacoes: "Formulário
 *     incompleto" na Obs.
 *   - Impedimento: motivo do servico nao ter sido executado.
 *
 * Fluxo de cada O.S.:
 *   1. lerOrdensServico     le a planilha de entrada (coluna Identificador).
 *   2. classificarOs        separa o que entra na folha do que e fora do escopo.
 *   3. abrirOs              filtra a listagem e abre a O.S. com o
 *                           identificador exato.
 *   4. lerDadosOs           clica em "Formularios" e pega, pelo coletor, os
 *                           dados que o proprio Field carregou (atividade,
 *                           O.S. e todos os formularios respondidos).
 *   5. montarLinhaPagamento aplica a regra (analisarOs/analisarFormulario) e
 *                           monta a linha da folha.
 *   6. gravarPlanilha       grava as abas Pagamento, Formularios e Fora do
 *                           escopo.
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
const {
  carregarCredenciais,
  carregarUrlsField,
} = require("../field/credenciais.js");
const {
  filtrarPorIdentificador,
  esperarListagemPronta,
} = require("../field/filtro.js");
const { fecharPopups } = require("../field/popups.js");
const { gravarPlanilha } = require("../planilha.js");

// ---------------------------------------------------------------------------
// Configuracao
// ---------------------------------------------------------------------------

// Planilha de entrada e saida ficam nesta pasta.
const BASE_DIR = __dirname;
const SAIDA_PADRAO = path.join(BASE_DIR, "pagamento_formularios.xlsx");

// Relatorio publico da O.S. — e o mesmo link que o "Compartilhar" devolve.
const URL_RELATORIO =
  "https://app.fieldcontrol.com.br/relacionamento-cliente/#/relatorio/";

// Seletores do Field (marcacao Palantir).
const SELETORES = {
  botaoEditarLinha: "palantir-button.palantir-table__action-button",
};

// Status da atividade no Field -> como a folha escreve.
const STATUS_FIELD = {
  done: "Concluída",
  reported: "Impedido",
  canceled: "Cancelada",
  cancelled: "Cancelada",
};

const HEADLESS = process.env.HEADLESS !== "false";

// ---------------------------------------------------------------------------
// Regras dos formularios
// ---------------------------------------------------------------------------

/**
 * Perguntas usadas na folha, localizadas pelo INICIO do texto (sem acento,
 * minusculo). Nao pelo numero: o numero muda entre formularios — o extravio
 * e a 37 na "REMOCAO - GERAL" e a 34 na "REMANEJAMENTO - REMOCAO".
 */
const PERGUNTAS = {
  removidos: "quais equipamentos estao sendo removidos",
  instalados: "quais equipamentos foram instalados",
  dano: "foi identificado algum dano nos equipamentos removidos",
  danoDescricao: "descreva os danos identificados",
  extravio: "algum equipamento extraviado",
  extravioQual: "qual equipamento foi extraviado",
  extravioOutros: 'qual "outros" equipamento foi extraviado',
  // "MANUTENCAO INOPRIME V2" lista o extravio direto, sem o "houve?".
  extravioLista: "liste quais os itens da inoprime que foram extraviados",
  // Violacao na instalacao: vai para a coluna Violacao.
  violacao: "alguma violacao na instalacao encontrada",
  violacaoDescricao: "descreva a violacao",
  impedimentoMotivo: "informe o motivo do servico nao ser executado",
  impedimentoDescricao:
    "descreva detalhadamente apresentando toda especificacao",
  impedimentoDescricao2:
    "descreva detalhadamente o motivo do servico nao ter sido executado",
  impedimentoSemAssinatura:
    "descreva melhor o motivo do cliente nao ter assinado",
};

/**
 * Equipamentos com o nome que a folha usa, na ordem em que ela escreve.
 *
 * GalileoSky/JR12/JR11 sao a telemetria avancada; o rastreador basico e uma
 * telemetria basica. Sozinhos viram "telemetria" e "telemetria basica"; os
 * dois juntos, "telemetria basica e telemetria avançada" (ver
 * rotularEquipamentos).
 */
const ROTULOS_EQUIPAMENTO = [
  [/rastreador/, "telemetria basica"],
  [/galileo|jr1[12]/, "telemetria"],
  [/camera/, "cam"],
  [/bafometro/, "bafometro"],
  [/satelital/, "satelital"],
];

/**
 * Perguntas que todo formulario daquele tipo precisa ter respondidas para a
 * folha. Se alguma existe no formulario, aparece sempre (nao depende de outra
 * resposta) e ficou em branco, o formulario e "incompleto" — o tecnico
 * enviou so com fotos, sem as informacoes.
 */
const PERGUNTAS_ESSENCIAIS = [
  "removidos",
  "instalados",
  "dano",
  "extravio",
  "violacao",
];

/** Remove acentos e normaliza para comparacao. */
function normalizar(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Texto da resposta de uma pergunta do formulario ("" se nao respondida).
 *
 * `formulario.questions` traz TODAS as perguntas do formulario, inclusive as
 * condicionais que nao apareceram para o tecnico (answer = null). Por isso
 * "pergunta nao encontrada" e "pergunta sem resposta" dao o mesmo "".
 */
function respostaDe(formulario, chave) {
  const pergunta = perguntaDe(formulario, chave);
  return pergunta ? textoDaResposta(pergunta) : "";
}

/** A pergunta do formulario cujo texto comeca com PERGUNTAS[chave], ou null. */
function perguntaDe(formulario, chave) {
  const inicio = PERGUNTAS[chave];
  return (
    formulario.questions.find((q) => normalizar(q.title).startsWith(inicio)) ||
    null
  );
}

/**
 * O formulario foi enviado sem as informacoes da folha?
 *
 * So contam as perguntas essenciais que existem neste formulario e que nao
 * tem condicao para aparecer: uma pergunta condicional (ex.: "descreva o
 * dano", que so aparece depois de um SIM) em branco e normal.
 */
function formularioIncompleto(formulario) {
  return PERGUNTAS_ESSENCIAIS.some((chave) => {
    const pergunta = perguntaDe(formulario, chave);
    if (!pergunta) return false;
    const sempreAparece = !(pergunta.conditions || []).length;
    return sempreAparece && !textoDaResposta(pergunta);
  });
}

/**
 * Converte o `answer` do Field em texto.
 *
 * O formato muda com o tipo da pergunta:
 *   - texto/sim-nao/multipla escolha: string ("SIM", "GalileoSky&JR12");
 *   - foto unica ("picture"): string com a URL da imagem;
 *   - varias fotos ("multiple-picture"): lista de URLs.
 * Fotos viram contagem — a URL nao interessa para a folha.
 */
function textoDaResposta(pergunta) {
  const { answer } = pergunta;
  if (answer == null) return "";
  if (Array.isArray(answer)) return `${answer.length} foto(s)`;
  if (typeof answer === "string" && /picture|signature/.test(pergunta.type)) {
    return answer ? "1 foto" : "";
  }
  // Qualquer outro formato inesperado vai como JSON, para nao sumir.
  return String(
    typeof answer === "string" ? answer : JSON.stringify(answer),
  ).trim();
}

// "SIM" / "Sim" / "sim" -> true. "NÃO", vazio e nao respondida -> false.
const ehSim = (valor) => normalizar(valor) === "sim";

// O Field junta as opcoes marcadas com "&": "GalileoSky&Câmeras" -> ["GalileoSky", "Câmeras"].
const separarOpcoes = (valor) =>
  String(valor || "")
    .split("&")
    .map((v) => v.trim())
    .filter(Boolean);

/**
 * "GalileoSky&Câmeras&JR12" -> ["telemetria", "cam"], na ordem da folha.
 * "GalileoSky&Rastreador básico" -> ["telemetria basica", "telemetria avançada"].
 */
function rotularEquipamentos(valor) {
  // A ordem de ROTULOS_EQUIPAMENTO e a ordem em que a folha escreve
  // ("telemetria e cam"), nao a ordem em que o tecnico marcou.
  const ordem = ROTULOS_EQUIPAMENTO.map(([, rotulo]) => rotulo);
  const rotulos = [];
  for (const item of separarOpcoes(valor)) {
    const achado = ROTULOS_EQUIPAMENTO.find(([padrao]) =>
      padrao.test(normalizar(item)),
    );
    // Equipamento sem rotulo conhecido ("OUTROS") vai com o nome original.
    const rotulo = achado ? achado[1] : item.toLowerCase();
    // GalileoSky e JR12 viram os dois "telemetria": entra uma vez so.
    if (!rotulos.includes(rotulo)) rotulos.push(rotulo);
  }
  const posicao = (r) => (ordem.includes(r) ? ordem.indexOf(r) : ordem.length);
  rotulos.sort((a, b) => posicao(a) - posicao(b));

  // Basica e avancada juntas: "telemetria" sozinho ficaria ambiguo, entao a
  // avancada ganha o nome completo.
  if (rotulos.includes("telemetria basica") && rotulos.includes("telemetria")) {
    rotulos[rotulos.indexOf("telemetria")] = "telemetria avançada";
  }
  return rotulos;
}

/** ["telemetria", "cam", "bafometro"] -> "telemetria, cam e bafometro" */
function juntarComE(itens) {
  if (itens.length <= 1) return itens.join("");
  return `${itens.slice(0, -1).join(", ")} e ${itens[itens.length - 1]}`;
}

/** "telemetria" -> "Telemetria extraviada", como a equipe escreve na folha. */
function descreverExtravio(item) {
  // Concordancia: "Telemetria basica extraviada", "Bafometro extraviado".
  const feminino = /telemetria|cam|camera|antena|leitora|placa/.test(
    normalizar(item),
  );
  const nome = item.charAt(0).toUpperCase() + item.slice(1).toLowerCase();
  return `${nome} ${feminino ? "extraviada" : "extraviado"}`;
}

/**
 * Avaria e trechos da Obs de UM formulario respondido.
 *
 * O impedimento e reconhecido pelo nome do formulario. Os demais, pelas
 * perguntas que tem: cada regra abaixo so produz algo se a pergunta existir
 * e estiver respondida. Assim o "REMANEJAMENTO - REMOCAO" cai na mesma regra
 * do "FORMULARIO DE REMOCAO - GERAL" sem precisar ser cadastrado, e um
 * formulario que nao tem a pergunta simplesmente nao contribui.
 */
function analisarFormulario(formulario) {
  // avaria/violacao: marcam as colunas Avaria e Violacao; obs: trechos que
  // vao para a coluna Obs, na ordem em que sao empilhados aqui.
  const resultado = { avaria: false, violacao: false, obs: [] };

  // --- Impedimento de visita: so o motivo interessa ------------------------
  if (normalizar(formulario.name).includes("impedimento")) {
    // A descricao detalhada (pergunta 2) e o que a folha manual usa. Se o
    // tecnico nao escreveu, cai na outra descricao (3) e, por ultimo, no
    // motivo marcado na lista (1).
    const descricao =
      respostaDe(formulario, "impedimentoDescricao") ||
      respostaDe(formulario, "impedimentoDescricao2") ||
      respostaDe(formulario, "impedimentoMotivo");
    // Pergunta 8: so aparece quando ninguem no local assinou a O.S.
    const semAssinatura = respostaDe(formulario, "impedimentoSemAssinatura");
    if (descricao) resultado.obs.push(descricao);
    if (semAssinatura) resultado.obs.push(semAssinatura);
    return resultado;
  }

  // --- Remocao e instalacao: o que saiu e o que entrou ---------------------
  // Vira "Remoção de telemetria e cam" / "Instalação de telemetria basica".
  const removidos = rotularEquipamentos(respostaDe(formulario, "removidos"));
  if (removidos.length)
    resultado.obs.push(`Remoção de ${juntarComE(removidos)}`);

  const instalados = rotularEquipamentos(respostaDe(formulario, "instalados"));
  if (instalados.length)
    resultado.obs.push(`Instalação de ${juntarComE(instalados)}`);

  // --- Dano nos equipamentos removidos = avaria ----------------------------
  if (ehSim(respostaDe(formulario, "dano"))) {
    resultado.avaria = true;
    // A descricao do dano e o que vai para a Obs ("Leitora quebrada...").
    const descricao = respostaDe(formulario, "danoDescricao");
    if (descricao) resultado.obs.push(descricao);
  }

  // --- Extravio na remocao: so na Obs, nao marca avaria --------------------
  if (ehSim(respostaDe(formulario, "extravio"))) {
    // "Qual equipamento foi extraviado?" vem da lista de equipamentos; se o
    // tecnico marcou "OUTROS", o nome esta na pergunta seguinte, escrito.
    const itens = [
      ...rotularEquipamentos(respostaDe(formulario, "extravioQual")).filter(
        (item) => item !== "outros",
      ),
      respostaDe(formulario, "extravioOutros"),
    ].filter(Boolean);
    // Disse que houve extravio mas nao disse o que: registra assim mesmo.
    resultado.obs.push(
      itens.length
        ? itens.map(descreverExtravio).join(", ")
        : "Equipamento extraviado",
    );
  }

  // --- Extravio na "MANUTENCAO INOPRIME V2": lista direta ------------------
  // Esse formulario nao pergunta "houve extravio?"; lista os itens. Respostas
  // do tipo "NAO" / "NENHUM" sao descartadas.
  const extraviados = separarOpcoes(
    respostaDe(formulario, "extravioLista"),
  ).filter((item) => !/^(nao|nenhum)/.test(normalizar(item)));
  if (extraviados.length) {
    resultado.obs.push(extraviados.map(descreverExtravio).join(", "));
  }

  // --- Violacao na manutencao: coluna Violacao ------------------------------
  // Violacao respondida como violacao conta como violacao, nao como avaria.
  if (ehSim(respostaDe(formulario, "violacao"))) {
    resultado.violacao = true;
    // "sem Buzzer", "Antena GPS violada"...
    const descricao = respostaDe(formulario, "violacaoDescricao");
    if (descricao) resultado.obs.push(descricao);
  }

  // --- Enviado so com fotos, sem as informacoes ----------------------------
  if (formularioIncompleto(formulario))
    resultado.obs.push("Formulário incompleto");

  return resultado;
}

/**
 * Junta os formularios da O.S. numa linha da folha.
 *
 * Uma O.S. pode ter mais de um formulario: o remanejamento tem o de remocao
 * e o de instalacao, e as vezes o tecnico abre um impedimento por engano
 * junto com o formulario certo. Avaria e violacao valem se QUALQUER
 * formulario marcar; a Obs junta os trechos de todos, separados por " | ".
 */
function analisarOs(formularios) {
  // Remocao antes da instalacao, como a folha escreve o remanejamento:
  // "Remoção de telemetria e cam | Instalação de telemetria e cam".
  const ordenados = [...formularios].sort(
    (a, b) =>
      Number(normalizar(b.name).includes("remocao")) -
      Number(normalizar(a.name).includes("remocao")),
  );

  let avaria = false;
  let violacao = false;
  const obs = [];
  for (const formulario of ordenados) {
    // Formulario aberto por engano e largado sem resposta nao conta.
    // status "done" = finalizado; "in-progress" = aberto e nao enviado.
    if (formulario.status !== "done") {
      obs.push(`Formulário "${formulario.name}" aberto e não finalizado`);
      continue;
    }
    const resultado = analisarFormulario(formulario);
    avaria = avaria || resultado.avaria;
    violacao = violacao || resultado.violacao;
    obs.push(...resultado.obs);
  }

  // O.S. concluida sem nenhum formulario: avisa em vez de deixar a Obs vazia.
  if (!formularios.length) obs.push("Sem formulário respondido no Field");

  return { avaria, violacao, obs: obs.join(" | ") };
}

/**
 * Data de conclusao no fuso de Brasilia, como serial do Excel.
 *
 * O Field manda a data em UTC isso converte para o
 * fuso de Brasilia*/
function dataExcel(iso) {
  if (!iso) return "";
  const [dia, mes, ano] = new Date(iso)
    .toLocaleDateString("pt-BR", { timeZone: "America/Sao_Paulo" })
    .split("/");
  return (Date.UTC(+ano, +mes - 1, +dia) - Date.UTC(1899, 11, 30)) / 86400000;
}

// ---------------------------------------------------------------------------
// Planilha de entrada
// ---------------------------------------------------------------------------

/**
 * Classifica a O.S. a partir do Tipo e do Status da planilha.
 * Retorna "remocao", "instalacao", "remanejamento", "manutencao", "impedido"
 * ou null (fora do escopo).
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

/**
 * Sem argumento, pega a planilha mais recente desta pasta.
 * Ignora as saidas do proprio fechamento — inclusive as copias com data e
 * hora gravadas quando a principal estava aberta no Excel.
 */
function descobrirPlanilhaEntrada() {
  const prefixoSaida = path.basename(SAIDA_PADRAO, ".xlsx");
  return (
    fs
      .readdirSync(BASE_DIR)
      .filter(
        (nome) =>
          /\.(xlsx|xlsm|csv)$/i.test(nome) &&
          !nome.startsWith("~$") &&
          !nome.startsWith(prefixoSaida),
      )
      .map((nome) => path.join(BASE_DIR, nome))
      .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0] ||
    null
  );
}

// ---------------------------------------------------------------------------
// Automacao
// ---------------------------------------------------------------------------

async function login(page) {
  const { email, senha } = carregarCredenciais();
  const urls = carregarUrlsField();
  const destinoListagem = new URL(urls.listagem);

  await page.goto(urls.login);

  await page.click('input[name="email"]');
  await page.fill('input[name="email"]', email);
  await page.click('palantir-button:has-text("Continuar")');

  await page.click('input[aria-label="password"]');
  await page.fill('input[aria-label="password"]', senha);
  await page.click('palantir-button:has-text("Continuar")');

  // O redirecionamento depois da senha pode levar mais de 5s. Navegar para a
  // listagem antes dele terminar interrompe o fluxo de autenticacao.
  await page.waitForURL(
    (url) =>
      url.origin === destinoListagem.origin &&
      url.hash.startsWith(destinoListagem.hash),
    { timeout: 60000 },
  );
  await fecharPopups(page);

  // Espera a listagem aparecer de fato, em vez de um tempo fixo: o Field
  // demora a trocar o esqueleto pela barra de filtros nova.
  if (!(await esperarListagemPronta(page))) {
    throw new Error("A listagem de atividades nao carregou depois do login.");
  }
}

/**
 * Escuta as respostas GraphQL que o proprio Field recebe.*/
function criarColetor(page) {
  // O que o Field devolveu desde o ultimo limpar():
  //   tasks       atividades da busca na listagem (status, data de conclusao);
  //   orders      a O.S. aberta (id, cliente, tipo de servico);
  //   formAnswers formularios respondidos, com todas as perguntas e respostas.
  const dados = { tasks: [], orders: [], formAnswers: [] };

  // Chamado pelo Playwright a cada resposta de rede da pagina. So interessa
  // o GraphQL do Field; imagens, scripts e telemetria (Sentry) sao ignorados.
  const ouvir = async (resposta) => {
    if (!resposta.url().includes("/graphql")) return;
    // Resposta sem corpo JSON (ou ja descartada pelo navegador): ignora.
    const corpo = await resposta.json().catch(() => null);
    const data = corpo && corpo.data;
    if (!data) return;
    // Cada consulta GraphQL devolve um campo diferente; guarda pelo nome.
    if (Array.isArray(data.tasks?.items)) dados.tasks.push(...data.tasks.items);
    if (data.order?.identifier) dados.orders.push(data.order);
    // A primeira consulta de formAnswers so traz o total, sem os itens.
    if (Array.isArray(data.formAnswers?.items))
      dados.formAnswers.push(data.formAnswers);
  };

  page.on("response", ouvir);

  return {
    dados,
    // Chamado antes de cada O.S., para nao misturar com a anterior.
    limpar() {
      dados.tasks.length = 0;
      dados.orders.length = 0;
      dados.formAnswers.length = 0;
    },
    // Chamado no fim do fechamento, antes de fechar o navegador.
    parar() {
      page.off("response", ouvir);
    },
  };
}

/** Escapa o texto para usar literalmente dentro de uma RegExp. */
function escaparRegex(texto) {
  return String(texto).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Espera uma condicao sobre os dados coletados.*/
async function esperarDados(page, condicao, timeout) {
  const limite = Date.now() + timeout;
  while (Date.now() < limite) {
    const valor = condicao();
    if (valor) return valor;
    await page.waitForTimeout(200);
  }
  return condicao();
}

/**
 * Garante que a tela esteja na listagem antes da proxima O.S.*/
async function voltarParaListagem(page) {
  const URL_ATIVIDADES = carregarUrlsField().listagem;
  // Fecha um painel ou modal que tenha ficado aberto.
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(500);

  if (!page.url().startsWith(URL_ATIVIDADES)) {
    await page
      .goto(URL_ATIVIDADES, { waitUntil: "domcontentloaded" })
      .catch(() => {});
    await esperarListagemPronta(page);
    return;
  }

  // Confirma que a gaveta fechou, em vez de esperar um tempo fixo.
  await localizarAbaFormularios(page)
    .waitFor({ state: "hidden", timeout: 5000 })
    .catch(() => {});
}

/**
 * Localiza a linha do identificador e abre a O.S. pelo botao de edicao.*/
async function abrirOs(page, identificador) {
  await filtrarPorIdentificador(page, identificador);

  // Qualquer linha que contenha o texto: serve para saber se o filtro
  // respondeu.
  const qualquerLinha = page
    .locator('tr, [role="row"], palantir-table-row')
    .filter({ hasText: identificador })
    .first();

  const encontrou = await qualquerLinha
    .waitFor({ state: "visible", timeout: 20000 })
    .then(() => true)
    .catch(() => false);

  if (!encontrou) {
    throw new Error(
      `Identificador ${identificador} nao encontrado na listagem.`,
    );
  }

  // Vale a linha com uma celula igual ao
  //identificador; so se nenhuma bater cai na primeira que contem o texto.
  const exato = new RegExp(`^\\s*${escaparRegex(identificador)}\\s*$`, "i");
  const linhaExata = page
    .locator('tr, [role="row"], palantir-table-row')
    .filter({ has: page.locator('td, [role="cell"]', { hasText: exato }) })
    .first();
  const linha = (await linhaExata.count().catch(() => 0))
    ? linhaExata
    : qualquerLinha;

  // Um aviso do Field por cima engole o clique sem dar erro.
  await fecharPopups(page);

  // O lapis de edicao, no fim da linha, e o que abre a gaveta. Se a marcacao
  // mudar e ele sumir, clicar na propria linha costuma abrir tambem.
  const botaoEditar = linha.locator(SELETORES.botaoEditarLinha).last();
  const alvo = (await botaoEditar.count().catch(() => 0)) ? botaoEditar : linha;

  // Duas tentativas: o primeiro clique as vezes chega antes da tabela
  // terminar de redesenhar depois do filtro e nao abre nada.
  for (let tentativa = 1; tentativa <= 2; tentativa += 1) {
    await alvo.click({ timeout: 5000 }).catch(async () => {
      await alvo.click({ force: true }).catch(() => {});
    });

    if (await esperarGaveta(page)) return page;
    await page.waitForTimeout(1500);
  }

  throw new Error(`O clique em Editar nao abriu a O.S. ${identificador}`);
}

/**
 * A aba "Formularios" da gaveta.*/
function localizarAbaFormularios(page) {
  return page
    .locator("button.palantir-tabs__button")
    .filter({ hasText: /Formul[aá]rios/i })
    .first();
}

/** A gaveta da O.S. esta aberta? Confirma pela barra de abas. */
async function esperarGaveta(page, timeout = 15000) {
  return localizarAbaFormularios(page)
    .waitFor({ state: "visible", timeout })
    .then(() => true)
    .catch(() => false);
}

/**
 * Com a gaveta aberta, le a O.S. e os formularios que o Field carregou.
 * O clique na aba "Formularios" e o que dispara a consulta das respostas.
 */
async function lerDadosOs(page, coletor, identificador) {
  const ident = normalizar(identificador);

  // 1. A O.S. aberta. Chega sozinha quando a gaveta abre. Confere pelo
  //    identificador exato, para garantir que a gaveta e a da O.S. certa.
  const order = await esperarDados(
    page,
    () => coletor.dados.orders.find((o) => normalizar(o.identifier) === ident),
    15000,
  );
  if (!order) {
    throw new Error(`O Field nao devolveu os dados da O.S. ${identificador}.`);
  }

  // 2. Os formularios. So aceita respostas desta O.S.: uma consulta atrasada
  //    da anterior nao pode entrar no lugar. Lista vazia (O.S. sem formulario)
  //    tambem e aceita — e um resultado valido, nao uma falha.
  const formAnswersDaOs = () =>
    coletor.dados.formAnswers.find((fa) =>
      fa.items.every((item) => !item.order || item.order.id === order.id),
    );

  const aba = localizarAbaFormularios(page);
  // As vezes o proprio Field falha e mostra "Erro ao carregar os formularios
  // da OS" com um botao "Tentar novamente" (aconteceu no SHX9B84). Clicar de
  // novo na aba nao refaz a consulta; o botao, sim.
  const tentarNovamente = page
    .locator("button, palantir-button")
    .filter({ hasText: /Tentar novamente/i })
    .first();

  // Ate 3 tentativas: clica na aba (ou no "Tentar novamente", se o Field
  // mostrou o erro) e espera ate 10s pela resposta.
  let formAnswers = null;
  for (let tentativa = 1; tentativa <= 3 && !formAnswers; tentativa += 1) {
    const falhouNoField = await page
      .getByText(/Erro ao carregar os formul[aá]rios/i)
      .first()
      .isVisible()
      .catch(() => false);
    const alvo = falhouNoField ? tentarNovamente : aba;

    await alvo.click({ timeout: 5000 }).catch(async () => {
      await alvo.click({ force: true }).catch(() => {});
    });
    formAnswers = await esperarDados(page, formAnswersDaOs, 10000);
  }

  if (!formAnswers) {
    throw new Error(
      `O Field nao devolveu os formularios da O.S. ${identificador}`,
    );
  }
  // Nunca aconteceu (as O.S. tem 1 ou 2 formularios), mas se o Field um dia
  // paginar, o aviso aparece no terminal em vez de faltar formulario calado.
  if (formAnswers.pageInfo?.hasNextPage) {
    console.warn(
      `${identificador}: o Field paginou os formularios; so a 1a pagina foi lida.`,
    );
  }

  // 3. A atividade (status e data de conclusao). Veio na busca da listagem.
  //    A busca pode listar mais de uma atividade (O.S. antigas com o mesmo
  //    identificador): vale a desta O.S.
  const task =
    coletor.dados.tasks.find((t) => t.order?.id === order.id) ||
    coletor.dados.tasks.find(
      (t) => normalizar(t.order?.identifier) === ident,
    ) ||
    null;

  return { order, task, formularios: formAnswers.items };
}

// ---------------------------------------------------------------------------
// Orquestracao
// ---------------------------------------------------------------------------

// Colunas da folha, na ordem da planilha usada com as empresas.
const COLUNAS_PAGAMENTO = [
  "Cliente",
  "Identificador",
  "Tipo",
  "Status",
  "Data Execução",
  "Avaria",
  "Violação",
  "Obs",
  "Link",
];

const COLUNAS_FORMULARIOS = [
  "Identificador",
  "Tipo",
  "Formulario",
  "Status do formulario",
  "Numero",
  "Pergunta",
  "Resposta",
];

/**
 * Linha da folha a partir dos dados lidos do Field.
 *
 * `os` e a linha da planilha de entrada; order/task/formularios vieram do
 * Field (lerDadosOs). Onde os dois tem a informacao, a origem escolhida esta
 * comentada em cada coluna.
 */
function montarLinhaPagamento(os, { order, task, formularios }) {
  // Avaria, Violacao e Obs saem da regra dos formularios.
  const { avaria, violacao, obs } = analisarOs(formularios);
  // Data de execucao = conclusao da atividade. Sem atividade, usa o
  // formulario finalizado por ultimo.
  const concluidoEm =
    task?.completedAt ||
    formularios
      .map((f) => f.completedAt)
      .filter(Boolean)
      .sort()
      .pop() ||
    "";

  return {
    // O cliente da planilha exportada e o que a folha usa (ex.: "InoPrime"
    // em instalacoes faturadas para a InoPrime); o cadastro do Field e reserva.
    Cliente: os.cliente || order.customer?.name || "",
    Identificador: os.identificador,
    // Tipo de servico cadastrado no Field ("Remoção", "Remanejamento de Veiculo").
    Tipo: order.service?.description || os.tipo,
    // Status atual no Field, traduzido (done -> Concluída, reported -> Impedido).
    // Status desconhecido vai como o Field mandou, para nao sumir.
    Status: STATUS_FIELD[task?.status] || task?.status || os.status,
    "Data Execução": dataExcel(concluidoEm),
    Avaria: avaria ? "SIM" : "NÃO",
    Violação: violacao ? "SIM" : "NÃO",
    Obs: obs,
    // O link do relatorio e sempre .../relatorio/<id da O.S.> (conferido nas
    // O.S. de exemplo): monta direto, sem clicar em "Compartilhar".
    Link: `${URL_RELATORIO}${order.id}`,
  };
}

/**
 * Linha da folha so com o que veio da planilha de entrada.
 * Usada quando a O.S. nao pode ser lida do Field (erro) e na aba "Fora do
 * escopo"; quem chama completa a Obs com o motivo.
 */
function montarLinhaBase(os) {
  return {
    Cliente: os.cliente,
    Identificador: os.identificador,
    Tipo: os.tipo,
    Status: os.status,
    "Data Execução": "",
    Avaria: "",
    Violação: "",
    Obs: "",
    Link: "",
  };
}

/**
 * Executa o fechamento.
 *
 * ordens     linhas da planilha de entrada (lerOrdensServico).
 * saidaPath  onde gravar a planilha; se estiver aberta no Excel, grava uma
 *            copia com data e hora e devolve esse caminho em outputPath.
 * onEvento   recebe o progresso ({ tipo, ... }) para a tela acompanhar:
 *            inicio, login, logado, os-inicio, os-ok, os-erro, parado,
 *            aviso e fim.
 * deveParar  consultada entre uma O.S. e a proxima (botao Parar).
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
  // Tipos que nao entram na folha ficam numa aba separada, para conferencia.
  const foraDoEscopo = [];
  // Todas as respostas, para conferir o que gerou cada linha da folha.
  const respostasCompletas = [];

  // Impedidas tambem sao abertas: o motivo esta no formulario de impedimento.
  const paraVisitar = [];
  for (const os of ordens) {
    const categoria = classificarOs(os);
    if (!categoria) {
      foraDoEscopo.push({
        ...montarLinhaBase(os),
        Obs: "Tipo sem formularios de pagamento",
      });
      continue;
    }
    paraVisitar.push({ ...os, categoria });
  }

  emitir({
    tipo: "inicio",
    total: ordens.length,
    aVisitar: paraVisitar.length,
    ignoradas: foraDoEscopo.length,
  });

  if (paraVisitar.length) {
    const browser = await chromium.launch({ headless: HEADLESS });
    const page = await browser.newPage();
    // Liga o coletor antes do login: ele fica escutando a pagina a execucao
    // toda e e limpo a cada O.S.
    const coletor = criarColetor(page);

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
          // Limpa antes do filtro: a busca ja traz a atividade (tasks).
          coletor.limpar();
          await abrirOs(page, os.identificador);
          const dados = await lerDadosOs(page, coletor, os.identificador);

          // Linha da aba Pagamento.
          const linha = montarLinhaPagamento(os, dados);
          resultados.push(linha);

          // Linhas da aba Formularios: toda pergunta respondida, de todos os
          // formularios da O.S., para conferir de onde saiu cada Obs/Avaria/
          // Violacao.
          // Perguntas sem resposta (condicionais que nao apareceram) ficam de
          // fora — so ocupariam espaco.
          let respondidas = 0;
          for (const formulario of dados.formularios) {
            for (const pergunta of formulario.questions) {
              const resposta = textoDaResposta(pergunta);
              if (!resposta) continue;
              respondidas += 1;
              respostasCompletas.push({
                Identificador: os.identificador,
                Tipo: linha.Tipo,
                Formulario: formulario.name,
                "Status do formulario": formulario.status,
                // position comeca em 0; o numero mostrado no Field e +1.
                Numero: pergunta.position + 1,
                Pergunta: pergunta.title,
                Resposta: resposta,
              });
            }
          }

          console.log(
            `${os.identificador} [${linha.Status}] Avaria ${linha.Avaria} ` +
              `Violação ${linha.Violação} | ` +
              `${dados.formularios.map((f) => f.name).join(" + ") || "sem formulario"} | ${linha.Obs}`,
          );
          emitir({
            tipo: "os-ok",
            identificador: os.identificador,
            categoria: os.categoria,
            respostas: respondidas,
            exportadas: dados.formularios.length,
          });
        } catch (erro) {
          // Falha numa O.S. nao derruba o fechamento: ela entra na folha com
          // o erro na Obs (para refazer a mao) e a execucao segue para a
          // proxima.
          console.error(`Falha em ${os.identificador}: ${erro.message}`);
          resultados.push({
            ...montarLinhaBase(os),
            Obs: `ERRO: ${erro.message}`,
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
      coletor.parar();
      await browser.close();
    }
  }

  const abas = [
    {
      nome: "Pagamento",
      linhas: resultados,
      ordem: COLUNAS_PAGAMENTO,
      formatos: { "Data Execução": "dd/mm/yyyy" },
    },
    {
      nome: "Formularios",
      linhas: respostasCompletas,
      ordem: COLUNAS_FORMULARIOS,
    },
    ...(foraDoEscopo.length
      ? [
          {
            nome: "Fora do escopo",
            linhas: foraDoEscopo,
            ordem: COLUNAS_PAGAMENTO,
          },
        ]
      : []),
  ];

  // Planilha anterior aberta no Excel: o Windows trava o arquivo (EBUSY) e a
  // gravacao falhava no fim, jogando fora todas as O.S. lidas. Nesse caso
  // grava uma copia com data e hora no nome e avisa onde ficou.
  try {
    gravarPlanilha(saidaPath, abas);
  } catch (erro) {
    if (!["EBUSY", "EPERM", "EACCES"].includes(erro.code)) throw erro;
    const carimbo = new Date()
      .toLocaleString("sv-SE", { timeZone: "America/Sao_Paulo" })
      .replace(/[-: ]/g, "")
      .slice(0, 12);
    const alternativa = saidaPath.replace(/\.xlsx$/i, `_${carimbo}.xlsx`);
    gravarPlanilha(alternativa, abas);
    console.warn(
      `${path.basename(saidaPath)} esta aberta em outro programa; ` +
        `a planilha foi gravada em ${alternativa}`,
    );
    emitir({
      tipo: "aviso",
      mensagem:
        `A planilha anterior está aberta no Excel. O resultado foi salvo em ` +
        `${path.basename(alternativa)}, na pasta folha-de-pagamento.`,
    });
    saidaPath = alternativa;
  }

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

  return {
    outputPath: saidaPath,
    resultados,
    respostasCompletas,
    interrompida,
  };
}

/**
 * Execucao direta: node Pagamento.js [planilha]. Sem argumento, usa a
 * planilha mais recente da pasta.
 */
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
  // Exportados para poder testar as regras sem abrir o Field.
  analisarFormulario,
  analisarOs,
  montarLinhaPagamento,
  rotularEquipamentos,
  PERGUNTAS,
  SAIDA_PADRAO,
};
