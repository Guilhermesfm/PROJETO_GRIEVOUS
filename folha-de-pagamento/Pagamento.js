/**
 * Pagamento.js - Financeiro Inoprime (fechamento da folha de pagamento).
 *
 * Le a planilha de servicos internos, abre cada O.S. e le as respostas dos
 * formularios: as de origem Field no Field Control, as de origem HubPrime
 * pelo relatorio no TrackerPrime (hubprime.js). Gera a folha com: Cliente,
 * Identificador, Tipo, Status, Data Execucao, Avaria, Violacao, Obs e Link.
 *
 * Regra:
 *   - Remocao / instalacao / remanejamento: equipamentos, dano e extravio.
 *   - Manutencao: violacao e extravio.
 *   - Impedimento: motivo do servico nao ter sido executado.
 *   - Revisao Rouxinol, Entrega e Devolucao e Laudo tecnico V2: so entram
 *     com avaria, extravio ou violacao (no laudo, foto de avaria).
 *   - Avaria no veiculo (antes ou depois do servico) nao entra.
 *   - Formulario so com fotos: "Formulário incompleto".
 *
 * Antes das O.S., o login le os formularios do Field e confere com o
 * registro (formularios.js).
 *
 * Fluxo de cada O.S.:
 *   1. lerOrdensServico     le a planilha de entrada.
 *   2. classificarOs        separa o que entra do que e fora do escopo.
 *   3. Field:    abrirOs + lerDadosOs + montarLinhaPagamento.
 *      HubPrime: lerRelatorio + comoFormulario (hubprime.js).
 *   4. registrarOs          decide se entra na folha.
 *   5. gravarPlanilha       grava as abas Pagamento, Formularios, Fora do
 *                           escopo e Conferencia de formularios.
 *
 * Uso:
 *   node folha-de-pagamento/Pagamento.js [planilha.xlsx] [saida.xlsx]
 *
 * Credenciais no .env: EMAIL e PASSWORD (Field), TRACKER_EMAIL e
 * TRACKER_PASSWORD (TrackerPrime).
 */

const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");
const XLSX = require("xlsx");
const { carregarUrlsField } = require("../field/credenciais.js");
const { entrarNoField } = require("../field/login.js");
const {
  filtrarPorIdentificador,
  esperarListagemPronta,
} = require("../field/filtro.js");
const { fecharPopups, clicar } = require("../field/popups.js");
const {
  abrirGaveta,
  fecharGaveta,
  localizarAbaFormularios,
} = require("../field/gaveta.js");
const { gravarPlanilha } = require("../planilha.js");
const {
  URL_FORMULARIOS,
  carregarRegistro,
  nomesRegistrados,
  normalizarNome,
  salvarListaField,
  conferirFormularios,
} = require("./formularios.js");
const {
  loginTracker,
  listarOsHubPrime,
  lerRelatorio,
  comoFormulario,
} = require("./hubprime.js");

// ---------------------------------------------------------------------------
// Configuracao
// ---------------------------------------------------------------------------

// Planilha de entrada e saida ficam nesta pasta.
const BASE_DIR = __dirname;
const SAIDA_PADRAO = path.join(BASE_DIR, "pagamento_formularios.xlsx");

// Link do relatorio da O.S.
const URL_RELATORIO =
  "https://app.fieldcontrol.com.br/relacionamento-cliente/#/relatorio/";

// Status do Field -> como a folha escreve.
const STATUS_FIELD = {
  done: "Concluída",
  reported: "Impedido",
  canceled: "Cancelada",
  cancelled: "Cancelada",
};

const HEADLESS = process.env.HEADLESS !== "false";

// Relatorios do HubPrime lidos ao mesmo tempo.
const ABAS_HUBPRIME = 3;

// ---------------------------------------------------------------------------
// Regras dos formularios
// ---------------------------------------------------------------------------

/**
 * Perguntas usadas na folha, achadas pelo inicio do texto (sem acento,
 * minusculo). Lista = textos diferentes para a mesma pergunta.*/
const PERGUNTAS = {
  removidos: [
    "quais equipamentos estao sendo removidos",
    "qual tipo de equipamento esta sendo removido", // Remocao - Saritur
    "quais os perifericos estao sendo retirados", // Retirada Terrabel
  ],
  instalados: [
    "quais equipamentos foram instalados",
    "equipamento que sera instalado", // Novo formulario de instalacao
    "qual produto esta sendo instalado", // Instalacao - tecnico terceiro
    "quais perifericos estao sendo instalados", // Instalacao Terrabel
    "selecione os itens que estao sendo instalados", // HubPrime
  ],
  // HubPrime: remocao so tem a legenda da foto dos seriais.
  removidosLegenda: "insira fotos dos seriais dos aparelhos retirados",
  dano: "foi identificado algum dano nos equipamentos removidos",
  danoDescricao: "descreva os danos identificados",
  extravio: [
    "algum equipamento extraviado",
    "houve extravio de algum equipamento inoprime", // HubPrime manutencao
    "algum periferico foi extraviado", // HubPrime remocao
  ],
  // HubPrime: descricao do extravio.
  extravioEvidencia: "evidencie o extravio",
  extravioQual: "qual equipamento foi extraviado",
  extravioOutros: 'qual "outros" equipamento foi extraviado',
  // Manutencao InoPrime V2: lista dos itens extraviados.
  extravioLista: "liste quais os itens da inoprime que foram extraviados",
  violacao: [
    "alguma violacao na instalacao encontrada",
    "alguma violacao foi encontrada na instalacao", // Novo formulario de manutencao
    "violacao na instalacao", // Revisao Rouxinol
    "houve violacao em alguns dos lacres", // HubPrime
  ],
  // "VIOACAO" esta escrito assim no Field.
  violacaoDescricao: ["descreva a violacao", "descreva detalhadamente a vioacao"],
  perifericoAvariado: [
    "algum periferico inoprime esta avariado", // Remanejamento InoPrime V2
    "algum periferico da inoprime esta avariado", // HubPrime manutencao
    "existe avarias nos perifericos da inoprime", // HubPrime remocao
  ],
  // Entrega e devolucao de materiais.
  devolvidosAvaria: "nos materiais devolvidos existe alguma avaria",
  devolvidosAvariaDescricao: "relate as avarias encontradas nos equipamentos",
  devolvidosExtravio: "existe algum equipamento ou periferico extraviado",
  devolvidosExtravioDescricao: "descreva oque foi extraviado",
  // Laudo tecnico InoPrime V2: foto de avaria = avaria.
  laudoAparelho: "qual aparelho esta na manutencao",
  laudoConstatado: "descreva o que foi constatado",
  laudoFotosAvaria: "tire fotos das avarias",
  impedimentoMotivo: "informe o motivo do servico nao ser executado",
  impedimentoDescricao:
    "descreva detalhadamente apresentando toda especificacao",
  impedimentoDescricao2:
    "descreva detalhadamente o motivo do servico nao ter sido executado",
  impedimentoSemAssinatura:
    "descreva melhor o motivo do cliente nao ter assinado",
};

/**
 * Nome de cada equipamento na folha, na ordem em que ela escreve.
 * "GALIELO" esta escrito assim na remocao da Saritur.*/
const ROTULOS_EQUIPAMENTO = [
  [/rastreador/, "telemetria basica"],
  [/galileo|galielo|j-?r1[12]/, "telemetria"],
  [/camera|momenta|\bv5\b|nt407|videometria/, "cam"],
  [/bafometro/, "bafometro"],
  [/satelital/, "satelital"],
];

/**
 * Perguntas que nao podem ficar em branco. Se uma delas existe, nao depende
 * de outra resposta e esta vazia, o formulario e "incompleto".*/
const PERGUNTAS_ESSENCIAIS = [
  "removidos",
  "instalados",
  "dano",
  "extravio",
  "violacao",
];

// Remove acentos e normaliza para comparacao.
const normalizar = normalizarNome;

/** Resposta de uma pergunta do formulario ("" se nao existe ou nao respondida). */
function respostaDe(formulario, chave) {
  const pergunta = perguntaDe(formulario, chave);
  return pergunta ? textoDaResposta(pergunta) : "";
}

/** A pergunta do formulario que comeca com o texto de PERGUNTAS[chave]. */
function perguntaDe(formulario, chave) {
  const inicios = [].concat(PERGUNTAS[chave]);
  return (
    formulario.questions.find((q) => {
      const titulo = normalizar(q.title);
      return inicios.some((inicio) => titulo.startsWith(inicio));
    }) || null
  );
}

/** Checa se o formulario foi enviado sem as informacoes da folha. */
function formularioIncompleto(formulario) {
  return PERGUNTAS_ESSENCIAIS.some((chave) => {
    const pergunta = perguntaDe(formulario, chave);
    if (!pergunta) return false;
    const sempreAparece = !(pergunta.conditions || []).length;
    return sempreAparece && !textoDaResposta(pergunta);
  });
}

/**
 * Converte a resposta do Field em texto. Fotos viram contagem.*/
function textoDaResposta(pergunta) {
  const { answer } = pergunta;
  if (answer == null) return "";
  if (Array.isArray(answer)) return `${answer.length} foto(s)`;
  if (typeof answer === "string" && /picture|signature/.test(pergunta.type)) {
    return answer ? "1 foto" : "";
  }
  // Formato inesperado vai como JSON.
  return String(
    typeof answer === "string" ? answer : JSON.stringify(answer),
  ).trim();
}

/** Quantidade de fotos anexadas numa pergunta de foto. */
function quantidadeFotos(pergunta) {
  const { answer } = pergunta || {};
  if (Array.isArray(answer)) return answer.filter(Boolean).length;
  return answer ? 1 : 0;
}

// "SIM" -> true. "NÃO" ou vazio -> false. No HubPrime vem "SIM / legenda".
const ehSim = (valor) => normalizar(valor).split(" / ")[0] === "sim";

// Field: "GalileoSky&Câmeras". HubPrime: "TELEMETRIA, BUZZER".
const separarOpcoes = (valor) =>
  String(valor || "")
    .split(/&|,\s+/)
    .map((v) => v.trim())
    .filter(Boolean);

/**
 * "GalileoSky&Câmeras&JR12" -> ["telemetria", "cam"].
 * "GalileoSky&Rastreador básico" -> ["telemetria basica", "telemetria avançada"].
 */
function rotularEquipamentos(valor) {
  const ordem = ROTULOS_EQUIPAMENTO.map(([, rotulo]) => rotulo);
  const rotulos = [];
  for (const item of separarOpcoes(valor)) {
    const achado = ROTULOS_EQUIPAMENTO.find(([padrao]) =>
      padrao.test(normalizar(item)),
    );
    // Sem rotulo, vai com o proprio nome.
    const rotulo = achado ? achado[1] : item.toLowerCase();
    if (!rotulos.includes(rotulo)) rotulos.push(rotulo);
  }
  const posicao = (r) => (ordem.includes(r) ? ordem.indexOf(r) : ordem.length);
  rotulos.sort((a, b) => posicao(a) - posicao(b));

  // Basica e avancada juntas.
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

/** "telemetria" -> "Telemetria extraviada". */
function descreverExtravio(item) {
  const feminino = /telemetria|cam|camera|antena|leitora|placa/.test(
    normalizar(item),
  );
  const nome = item.charAt(0).toUpperCase() + item.slice(1).toLowerCase();
  return `${nome} ${feminino ? "extraviada" : "extraviado"}`;
}

/**
 * Avaria, violacao, extravio e Obs de um formulario.
 * Cada regra so conta se a pergunta existir no formulario.*/
function analisarFormulario(formulario) {
  const resultado = { avaria: false, violacao: false, extravio: false, obs: [] };

  // --- Impedimento: descricao do motivo e motivo de nao ter assinado -------
  if (normalizar(formulario.name).includes("impedimento")) {
    const descricao =
      respostaDe(formulario, "impedimentoDescricao") ||
      respostaDe(formulario, "impedimentoDescricao2") ||
      respostaDe(formulario, "impedimentoMotivo");
    const semAssinatura = respostaDe(formulario, "impedimentoSemAssinatura");
    if (descricao) resultado.obs.push(descricao);
    if (semAssinatura) resultado.obs.push(semAssinatura);
    return resultado;
  }

  // --- Equipamentos removidos e instalados ---------------------------------
  const removidos = rotularEquipamentos(respostaDe(formulario, "removidos"));
  if (removidos.length)
    resultado.obs.push(`Remoção de ${juntarComE(removidos)}`);

  // HubPrime: legenda da foto dos seriais ("V5 Galileo"); generica vira so
  // "Remoção".
  const legenda = respostaDe(formulario, "removidosLegenda");
  if (perguntaDe(formulario, "removidosLegenda")) {
    const generica =
      !legenda ||
      /^(ok+|equipamentos?|aparelhos?|seriais|serial|fotos?|placa|sim|nao|remocao ok)$/.test(
        normalizar(legenda),
      );
    resultado.obs.push(generica ? "Remoção" : `Remoção: ${legenda}`);
  }

  const instalados = rotularEquipamentos(respostaDe(formulario, "instalados"));
  if (instalados.length)
    resultado.obs.push(`Instalação de ${juntarComE(instalados)}`);

  // --- Dano nos equipamentos removidos: avaria -----------------------------
  if (ehSim(respostaDe(formulario, "dano"))) {
    resultado.avaria = true;
    const descricao = respostaDe(formulario, "danoDescricao");
    if (descricao) resultado.obs.push(descricao);
  }

  // --- Extravio na remocao -------------------------------------------------
  if (ehSim(respostaDe(formulario, "extravio"))) {
    // "OUTROS" tem o nome escrito na pergunta seguinte.
    const itens = [
      ...rotularEquipamentos(respostaDe(formulario, "extravioQual")).filter(
        (item) => item !== "outros",
      ),
      respostaDe(formulario, "extravioOutros"),
    ].filter(Boolean);
    resultado.extravio = true;
    // HubPrime: o que foi extraviado vem escrito em "Evidencie o extravio".
    const evidencia = respostaDe(formulario, "extravioEvidencia");
    resultado.obs.push(
      itens.length
        ? itens.map(descreverExtravio).join(", ")
        : evidencia
          ? `Extraviado: ${evidencia}`
          : "Equipamento extraviado",
    );
  }

  // --- Extravio na Manutencao InoPrime V2 ----------------------------------
  const extraviados = separarOpcoes(
    respostaDe(formulario, "extravioLista"),
  ).filter((item) => !/^(nao|nenhum)/.test(normalizar(item)));
  if (extraviados.length) {
    resultado.extravio = true;
    resultado.obs.push(extraviados.map(descreverExtravio).join(", "));
  }

  // --- Periferico InoPrime avariado (Remanejamento V2) ---------------------
  if (ehSim(respostaDe(formulario, "perifericoAvariado"))) {
    resultado.avaria = true;
    resultado.obs.push("Periférico InoPrime avariado");
  }

  // --- Entrega e devolucao: avaria e extravio no material devolvido --------
  if (ehSim(respostaDe(formulario, "devolvidosAvaria"))) {
    resultado.avaria = true;
    resultado.obs.push(
      respostaDe(formulario, "devolvidosAvariaDescricao") ||
        "Material devolvido com avaria",
    );
  }
  if (ehSim(respostaDe(formulario, "devolvidosExtravio"))) {
    resultado.extravio = true;
    const descricao = respostaDe(formulario, "devolvidosExtravioDescricao");
    resultado.obs.push(
      descricao ? `Extraviado: ${descricao}` : "Equipamento extraviado",
    );
  }

  // --- Laudo tecnico V2: foto de avaria marca avaria -----------------------
  const fotosAvaria = quantidadeFotos(perguntaDe(formulario, "laudoFotosAvaria"));
  if (fotosAvaria) {
    resultado.avaria = true;
    const aparelho = juntarComE(
      rotularEquipamentos(respostaDe(formulario, "laudoAparelho")),
    );
    const constatado = respostaDe(formulario, "laudoConstatado");
    resultado.obs.push(
      `Avaria${aparelho ? ` em ${aparelho}` : ""}` +
        `${constatado ? `: ${constatado}` : ""} (${fotosAvaria} foto${fotosAvaria > 1 ? "s" : ""})`,
    );
  }

  // --- Violacao: coluna Violacao -------------------------------------------
  if (ehSim(respostaDe(formulario, "violacao"))) {
    resultado.violacao = true;
    resultado.obs.push(
      respostaDe(formulario, "violacaoDescricao") || "Violação na instalação",
    );
  }

  // --- Enviado so com fotos ------------------------------------------------
  if (formularioIncompleto(formulario))
    resultado.obs.push("Formulário incompleto");

  return resultado;
}

/**
 * Junta todos os formularios da O.S. numa linha.
 * Avaria e violacao valem se qualquer formulario marcar.*/
function analisarOs(formularios) {
  // Remocao antes da instalacao (remanejamento).
  const ordenados = [...formularios].sort(
    (a, b) =>
      Number(normalizar(b.name).includes("remocao")) -
      Number(normalizar(a.name).includes("remocao")),
  );

  let avaria = false;
  let violacao = false;
  let extravio = false;
  const obs = [];
  for (const formulario of ordenados) {
    // Formulario aberto e nao finalizado nao conta.
    if (formulario.status !== "done") {
      obs.push(`Formulário "${formulario.name}" aberto e não finalizado`);
      continue;
    }
    const resultado = analisarFormulario(formulario);
    avaria = avaria || resultado.avaria;
    violacao = violacao || resultado.violacao;
    extravio = extravio || resultado.extravio;
    obs.push(...resultado.obs);
  }

  if (!formularios.length) obs.push("Sem formulário respondido no Field");

  return { avaria, violacao, extravio, obs: obs.join(" | ") };
}

/** Data UTC do Field -> serial do Excel, no fuso de Brasilia. */
function dataExcel(iso) {
  if (!iso) return "";
  const [dia, mes, ano] = new Date(iso)
    .toLocaleDateString("pt-BR", { timeZone: "America/Sao_Paulo" })
    .split("/");
  return (Date.UTC(+ano, +mes - 1, +dia) - Date.UTC(1899, 11, 30)) / 86400000;
}

/** "10/11/2026 11:07" (horario de Brasilia) -> ISO; "" se vazio. */
function dataBrasilParaIso(texto) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/.exec(texto || "");
  return m ? `${m[3]}-${m[2]}-${m[1]}T${m[4]}:${m[5]}:00-03:00` : "";
}

// ---------------------------------------------------------------------------
// Planilha de entrada
// ---------------------------------------------------------------------------

/**
 * Classifica a O.S. pelo Tipo e Status.
 * "ocorrencia": so entra na folha com avaria, extravio ou violacao.
 * null: fora do escopo.*/
function classificarOs({ tipo, status }) {
  if (normalizar(status).includes("impedido")) return "impedido";

  const alvo = normalizar(tipo);
  if (alvo.includes("remanejamento")) return "remanejamento";
  if (alvo.includes("instalacao")) return "instalacao";
  // "3 - RETIRADA EXCLUSIVO TERRABEL".
  if (alvo.includes("remocao") || alvo.includes("retirada")) return "remocao";
  if (
    alvo.includes("manutencao") ||
    // "MANUTECAO" esta escrito assim no Field (Terrabel).
    alvo.includes("manutecao") ||
    alvo.includes("avaria") ||
    alvo.includes("violacao") ||
    alvo.includes("extravio")
  ) {
    return "manutencao";
  }
  if (
    alvo.includes("revisao rouxinol") ||
    alvo.includes("entrega e devolucao") ||
    alvo.includes("laudo tecnico inoprime v2")
  ) {
    return "ocorrencia";
  }
  return null;
}

/**
 * Le a planilha exportada do Field. O cabecalho e a linha que tem a coluna
 * "Identificador".*/
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
    // ID do TrackerPrime e Origem (Field ou HubPrime).
    id: coluna("ID"),
    origem: coluna("Origem"),
  };

  return linhas
    .slice(indiceCabecalho + 1)
    .map((linha) => ({
      identificador: String(linha[colunas.identificador] ?? "").trim(),
      tipo: String(linha[colunas.tipo] ?? "").trim(),
      cliente: String(linha[colunas.cliente] ?? "").trim(),
      status: String(linha[colunas.status] ?? "").trim(),
      id: String(linha[colunas.id] ?? "").trim(),
      // Sem a coluna Origem, tudo e do Field.
      origem: String(linha[colunas.origem] ?? "").trim() || "Field",
    }))
    .filter((os) => os.identificador);
}

/**
 * Planilha mais recente desta pasta, ignorando as saidas do fechamento.*/
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

/**
 * Faz login pela pagina de formularios, le a lista de formularios e vai
 * para a listagem de atividades. Retorna os formularios (null se nao vieram).*/
async function login(page) {
  const urls = carregarUrlsField();
  const loginFormularios =
    `${urls.login.split("?")[0]}?to=${encodeURIComponent(URL_FORMULARIOS)}`;

  // Pede a lista inteira de uma vez (o Field pagina de 25 em 25).
  await page.route("**/graphql", async (rota) => {
    let corpo = null;
    try {
      corpo = JSON.parse(rota.request().postData() || "null");
    } catch {
      // Pedido sem JSON segue como veio.
    }
    if (corpo?.operationName !== "FormsQuery") return rota.continue();
    corpo.variables = { ...corpo.variables, take: 200, skip: 0 };
    return rota.continue({ postData: JSON.stringify(corpo) });
  });

  let formularios = null;
  const ouvir = async (resposta) => {
    if (!resposta.url().includes("/graphql")) return;
    const corpo = await resposta.json().catch(() => null);
    if (Array.isArray(corpo?.data?.forms?.items)) formularios = corpo.data.forms.items;
  };
  page.on("response", ouvir);

  await entrarNoField(page, loginFormularios);

  // Espera a lista de formularios (o login termina junto).
  const limite = Date.now() + 60000;
  while (!formularios && Date.now() < limite) await page.waitForTimeout(300);
  page.off("response", ouvir);
  await page.unroute("**/graphql");
  if (!formularios && page.url().includes("autenticador")) {
    throw new Error("Login no Field nao concluido. Confira o e-mail e a senha.");
  }

  await page.goto(urls.listagem);
  await fecharPopups(page);

  if (!(await esperarListagemPronta(page))) {
    throw new Error("A listagem de atividades nao carregou depois do login.");
  }
  return formularios;
}

/**
 * Escuta as respostas GraphQL que o proprio Field recebe.*/
function criarColetor(page) {
  // tasks: atividades. orders: a O.S. aberta. formAnswers: formularios.
  const dados = { tasks: [], orders: [], formAnswers: [] };

  // Guarda so as respostas do GraphQL do Field.
  const ouvir = async (resposta) => {
    if (!resposta.url().includes("/graphql")) return;
    const corpo = await resposta.json().catch(() => null);
    const data = corpo && corpo.data;
    if (!data) return;
    if (Array.isArray(data.tasks?.items)) dados.tasks.push(...data.tasks.items);
    if (data.order?.identifier) dados.orders.push(data.order);
    if (Array.isArray(data.formAnswers?.items))
      dados.formAnswers.push(data.formAnswers);
  };

  page.on("response", ouvir);

  return {
    dados,
    // Limpa antes de cada O.S.
    limpar() {
      dados.tasks.length = 0;
      dados.orders.length = 0;
      dados.formAnswers.length = 0;
    },
    // Para de escutar no fim do fechamento.
    parar() {
      page.off("response", ouvir);
    },
  };
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
  const listagem = carregarUrlsField().listagem;
  if (page.url().startsWith(listagem)) return fecharGaveta(page);

  await page.goto(listagem, { waitUntil: "domcontentloaded" }).catch(() => {});
  await esperarListagemPronta(page);
}

/** Filtra a listagem pelo identificador e abre a O.S. */
async function abrirOs(page, identificador) {
  await filtrarPorIdentificador(page, identificador);
  await abrirGaveta(page, identificador);
}

/**
 * Le a O.S., a atividade e os formularios que o Field carregou.*/
async function lerDadosOs(page, coletor, identificador) {
  const ident = normalizar(identificador);

  // 1. A O.S. aberta, conferida pelo identificador.
  const order = await esperarDados(
    page,
    () => coletor.dados.orders.find((o) => normalizar(o.identifier) === ident),
    15000,
  );
  if (!order) {
    throw new Error(`O Field nao devolveu os dados da O.S. ${identificador}.`);
  }

  // 2. Os formularios desta O.S. Lista vazia tambem vale.
  const formAnswersDaOs = () =>
    coletor.dados.formAnswers.find((fa) =>
      fa.items.every((item) => !item.order || item.order.id === order.id),
    );

  const aba = localizarAbaFormularios(page);
  // Botao que o Field mostra quando falha ao carregar os formularios.
  const tentarNovamente = page
    .locator("button, palantir-button")
    .filter({ hasText: /Tentar novamente/i })
    .first();

  // Ate 3 tentativas de 10s.
  let formAnswers = null;
  for (let tentativa = 1; tentativa <= 3 && !formAnswers; tentativa += 1) {
    const falhouNoField = await page
      .getByText(/Erro ao carregar os formul[aá]rios/i)
      .first()
      .isVisible()
      .catch(() => false);
    await clicar(falhouNoField ? tentarNovamente : aba);
    formAnswers = await esperarDados(page, formAnswersDaOs, 10000);
  }

  if (!formAnswers) {
    throw new Error(
      `O Field nao devolveu os formularios da O.S. ${identificador}`,
    );
  }
  // Avisa se o Field paginar os formularios.
  if (formAnswers.pageInfo?.hasNextPage) {
    console.warn(
      `${identificador}: o Field paginou os formularios; so a 1a pagina foi lida.`,
    );
  }

  // 3. A atividade desta O.S. (status e data de conclusao).
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

// Colunas da folha.
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
 * Monta a linha da folha com a planilha de entrada e os dados do Field.*/
function montarLinhaPagamento(
  os,
  { order, task, formularios },
  { avaria, violacao, obs } = analisarOs(formularios),
) {
  // Conclusao da atividade; sem ela, o ultimo formulario finalizado.
  const concluidoEm =
    task?.completedAt ||
    formularios
      .map((f) => f.completedAt)
      .filter(Boolean)
      .sort()
      .pop() ||
    "";

  return {
    // Cliente e Tipo da planilha; o Field so se a planilha nao tiver.
    Cliente: os.cliente || order.customer?.name || "",
    Identificador: os.identificador,
    Tipo: os.tipo || order.service?.description || "",
    Status: STATUS_FIELD[task?.status] || task?.status || os.status,
    "Data Execução": dataExcel(concluidoEm),
    Avaria: avaria ? "SIM" : "NÃO",
    Violação: violacao ? "SIM" : "NÃO",
    Obs: obs,
    Link: `${URL_RELATORIO}${order.id}`,
  };
}

/**
 * Linha so com os dados da planilha de entrada (erro e fora do escopo).*/
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
 * ordens     linhas da planilha de entrada.
 * saidaPath  onde gravar a planilha.
 * onEvento   recebe o progresso para a tela.
 * deveParar  checada antes de cada O.S. (botao Parar).
 */
async function runPagamento(
  ordens,
  saidaPath = SAIDA_PADRAO,
  onEvento = () => {},
  deveParar = () => false,
) {
  let interrompida = false;
  const emitir = (evento) => {
    try {
      onEvento(evento);
    } catch {
      // Erro no progresso nao para o fechamento.
    }
  };

  const resultados = [];
  const foraDoEscopo = [];
  // Todas as respostas, para conferencia.
  const respostasCompletas = [];
  // Formularios conhecidos e o resultado da conferencia com o Field.
  const registro = carregarRegistro();
  const registrados = nomesRegistrados(registro);
  let conferencia = [];

  // Separa as O.S. que serao abertas.
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

  // Field e HubPrime sao processados em etapas separadas.
  const doHubPrime = (os) => normalizar(os.origem) === "hubprime";
  const paraField = paraVisitar.filter((os) => !doHubPrime(os));
  const paraHubPrime = paraVisitar.filter(doHubPrime);
  let processadas = 0;

  /** Pede parada: grava o que ja foi lido. */
  const parar = () => {
    if (!deveParar()) return false;
    interrompida = true;
    emitir({
      tipo: "parado",
      processadas,
      restantes: paraVisitar.length - processadas,
    });
    return true;
  };

  /** Progresso da O.S. que vai ser lida. */
  const iniciarOs = (os) => {
    processadas += 1;
    emitir({
      tipo: "os-inicio",
      identificador: os.identificador,
      categoria: os.categoria,
      origem: os.origem,
      atual: processadas,
      total: paraVisitar.length,
    });
  };

  /**
   * Decide se a O.S. lida entra na folha e guarda as respostas.
   * tipoSistema: servico cadastrado no Field ou no HubPrime.*/
  const registrarOs = (os, linha, analise, formularios, tipoSistema) => {
    const { avaria, violacao, extravio } = analise;
    const categoria =
      os.categoria === "impedido" || !tipoSistema
        ? os.categoria
        : classificarOs({ tipo: tipoSistema, status: "" });

    if (!categoria) {
      foraDoEscopo.push({
        ...linha,
        Obs: `Tipo no ${os.origem}: ${tipoSistema} (fora do escopo)`,
      });
    } else if (categoria === "ocorrencia" && !avaria && !violacao && !extravio) {
      // Revisao, entrega/devolucao e laudo sem ocorrencia.
      foraDoEscopo.push({
        ...linha,
        Obs: [
          tipoSistema !== linha.Tipo ? `Tipo no ${os.origem}: ${tipoSistema}` : "",
          "Sem avaria, extravio ou violação",
          linha.Obs,
        ]
          .filter(Boolean)
          .join(" | "),
      });
    } else {
      resultados.push(linha);
    }

    // Aba Formularios: perguntas respondidas.
    let respondidas = 0;
    for (const formulario of formularios) {
      for (const pergunta of formulario.questions) {
        const resposta = textoDaResposta(pergunta);
        if (!resposta) continue;
        respondidas += 1;
        respostasCompletas.push({
          Identificador: os.identificador,
          Tipo: linha.Tipo,
          Formulario: formulario.name,
          "Status do formulario": formulario.status,
          // Numero como aparece no Field/HubPrime.
          Numero: pergunta.position + 1,
          Pergunta: pergunta.title,
          Resposta: resposta,
        });
      }
    }

    console.log(
      `${os.identificador} [${linha.Status}] Avaria ${linha.Avaria} ` +
        `Violação ${linha.Violação} | ` +
        `${formularios.map((f) => f.name).join(" + ") || "sem formulario"} | ${linha.Obs}`,
    );
    emitir({
      tipo: "os-ok",
      identificador: os.identificador,
      categoria: categoria || "fora do escopo",
      origem: os.origem,
      respostas: respondidas,
      exportadas: formularios.length,
    });
  };

  /** Erro numa O.S.: vai para a Obs e o fechamento continua. */
  const registrarErro = (os, erro) => {
    console.error(`Falha em ${os.identificador}: ${erro.message}`);
    resultados.push({
      ...montarLinhaBase(os),
      Obs: `ERRO: ${erro.message}`,
    });
    emitir({
      tipo: "os-erro",
      identificador: os.identificador,
      categoria: os.categoria,
      origem: os.origem,
      erro: erro.message,
    });
  };

  if (paraVisitar.length) {
    const browser = await chromium.launch({ headless: HEADLESS });

    try {
      // --- O.S. do Field -----------------------------------------------------
      if (paraField.length) {
        const page = await browser.newPage();
        const coletor = criarColetor(page);

        emitir({ tipo: "login" });
        const formulariosField = await login(page);
        emitir({ tipo: "logado" });

        // Confere os formularios do Field antes de processar as O.S.
        if (formulariosField) {
          salvarListaField(formulariosField);
          conferencia = conferirFormularios(formulariosField, registro);
          const nomes = (situacao) =>
            conferencia
              .filter((l) => l.Situacao === situacao)
              .map((l) => ({ nome: l.Formulario, detalhe: l.Observacao }));
          const diferencas = {
            novos: nomes("Novo"),
            alterados: nomes("Alterado"),
            removidos: nomes("Removido"),
            semPerguntas: nomes("Sem perguntas"),
          };
          emitir({
            tipo: "formularios",
            total: formulariosField.filter((f) => !f.archived).length,
            completos: conferencia.filter((l) => l.Situacao === "OK").length,
            ...diferencas,
          });
          for (const [rotulo, itens] of [
            ["Formulário novo no Field", diferencas.novos],
            ["Formulário alterado no Field", diferencas.alterados],
            ["Formulário removido do Field", diferencas.removidos],
            ["Formulário sem perguntas registradas", diferencas.semPerguntas],
          ]) {
            for (const item of itens) console.warn(`${rotulo}: ${item.nome}`);
          }
        } else {
          emitir({
            tipo: "aviso",
            mensagem: "Não foi possível ler a lista de formulários do Field; conferência pulada.",
          });
        }

        for (const os of paraField) {
          if (parar()) break;
          iniciarOs(os);
          try {
            coletor.limpar();
            await abrirOs(page, os.identificador);
            const dados = await lerDadosOs(page, coletor, os.identificador);

            const analise = analisarOs(dados.formularios);
            const linha = montarLinhaPagamento(os, dados, analise);

            // Formulario fora do registro: avisa na Obs.
            if (registrados.size) {
              const naoRegistrados = dados.formularios
                .map((f) => f.name)
                .filter((nome) => !registrados.has(normalizarNome(nome)));
              if (naoRegistrados.length) {
                linha.Obs = [
                  linha.Obs,
                  ...naoRegistrados.map((nome) => `Formulário não registrado: ${nome}`),
                ]
                  .filter(Boolean)
                  .join(" | ");
              }
            }

            const tipoField = dados.order.service?.description || "";
            registrarOs(os, linha, analise, dados.formularios, tipoField);
          } catch (erro) {
            registrarErro(os, erro);
          }

          await voltarParaListagem(page);
        }
        coletor.parar();
        await page.close();
      }

      // --- O.S. do HubPrime (pelo TrackerPrime) ------------------------------
      if (paraHubPrime.length && !interrompida) {
        const page = await browser.newPage();
        let listaHub = null;
        try {
          emitir({ tipo: "aviso", mensagem: "Entrando no TrackerPrime para as O.S. do HubPrime…" });
          await loginTracker(page);
          listaHub = await listarOsHubPrime(page);
          // So as O.S. da planilha sao analisadas; a lista serve para achar o link.
          emitir({
            tipo: "aviso",
            mensagem: `HubPrime: ${paraHubPrime.length} O.S. da planilha para analisar.`,
          });
        } catch (erro) {
          emitir({ tipo: "aviso", mensagem: `HubPrime: ${erro.message}` });
        }

        /** Le o relatorio da O.S. e monta a linha. */
        const lerOsHubPrime = async (aba, os) => {
          if (!listaHub) throw new Error("TrackerPrime indisponível");
          const naLista = listaHub.get(os.id);
          if (!naLista) {
            throw new Error(`ID ${os.id || "(vazio)"} não encontrado no HubPrime.`);
          }

          const relatorio = await lerRelatorio(aba, naLista.link);
          const formulario = comoFormulario(relatorio);
          const analise = analisarOs([formulario]);
          const linha = {
            ...montarLinhaBase(os),
            Status: naLista.status || os.status,
            "Data Execução": dataExcel(dataBrasilParaIso(relatorio.concluidoEm)),
            Avaria: analise.avaria ? "SIM" : "NÃO",
            Violação: analise.violacao ? "SIM" : "NÃO",
            Obs: analise.obs,
            Link: naLista.link,
          };
          return [linha, analise, [formulario], relatorio.tipo || naLista.tipo];
        };

        // Abas extras para ler varios relatorios ao mesmo tempo.
        const paginas = [page];
        if (listaHub) {
          const extras = Math.min(ABAS_HUBPRIME, paraHubPrime.length) - 1;
          for (let i = 0; i < extras; i += 1) paginas.push(await page.context().newPage());
        }

        // Le em lotes e registra na ordem da planilha.
        for (let inicio = 0; inicio < paraHubPrime.length; inicio += paginas.length) {
          if (parar()) break;
          const lote = paraHubPrime.slice(inicio, inicio + paginas.length);
          // Le o lote junto; erro numa O.S. nao derruba as outras.
          lote.forEach(iniciarOs);
          const lidos = await Promise.allSettled(
            lote.map((os, i) => lerOsHubPrime(paginas[i], os)),
          );
          lote.forEach((os, i) => {
            if (lidos[i].status === "fulfilled") registrarOs(os, ...lidos[i].value);
            else registrarErro(os, lidos[i].reason);
          });
        }
        for (const pagina of paginas) await pagina.close();
      }
    } finally {
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
    ...(conferencia.length
      ? [
          {
            nome: "Conferência de formulários",
            linhas: conferencia,
            ordem: [
              "Formulario",
              "Situacao",
              "Perguntas no Field",
              "Perguntas registradas",
              "Alterado em",
              "Observacao",
            ],
          },
        ]
      : []),
  ];

  // Planilha aberta no Excel: grava uma copia com data e hora.
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
 * Execucao pelo terminal. Sem argumento, usa a planilha mais recente.*/
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
  // Para testar as regras sem abrir o Field.
  analisarFormulario,
  analisarOs,
  montarLinhaPagamento,
  rotularEquipamentos,
  PERGUNTAS,
  SAIDA_PADRAO,
};
