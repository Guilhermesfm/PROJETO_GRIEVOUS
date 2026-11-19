/**
 * formularios.js - Conferencia dos formularios do Field.
 *
 * Compara a lista de formularios do Field com o registro
 * (formularios-registrados.json) e aponta o que e novo, alterado ou removido.
 *
 * Uso:
 *   node folha-de-pagamento/formularios.js registrar [pasta com os .fct]
 */

const fs = require("node:fs");
const path = require("node:path");

const ARQUIVO_REGISTRO = path.join(__dirname, "formularios-registrados.json");
// Copia local da ultima lista lida do Field.
const PASTA_LOCAL = path.join(__dirname, "formularios");
const ULTIMA_LISTA = path.join(PASTA_LOCAL, "lista-field.json");

// Pagina de formularios do Field.
const URL_FORMULARIOS = "https://app.fieldcontrol.com.br/painel/#/formularios";

/** Nome sem acento, minusculo e sem espacos repetidos. */
function normalizarNome(nome) {
  return String(nome || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

/** Formularios registrados ([] se o arquivo nao existe). */
function carregarRegistro() {
  if (!fs.existsSync(ARQUIVO_REGISTRO)) return [];
  return JSON.parse(fs.readFileSync(ARQUIVO_REGISTRO, "utf8"));
}

/** Nomes registrados, normalizados. */
function nomesRegistrados(registro) {
  return new Set(registro.map((r) => normalizarNome(r.nome)));
}

/** Salva a lista lida do Field na pasta local. */
function salvarListaField(itens) {
  fs.mkdirSync(PASTA_LOCAL, { recursive: true });
  fs.writeFileSync(
    ULTIMA_LISTA,
    JSON.stringify({ lidoEm: new Date().toISOString(), itens }, null, 2),
  );
}

// "2026-10-08T13:04:15.100Z" -> "08/10/2026".
const dataCurta = (iso) =>
  iso ? new Date(iso).toLocaleDateString("pt-BR", { timeZone: "America/Sao_Paulo" }) : "";

/**
 * Linhas da aba "Conferência de formulários": Novo, Alterado, Removido,
 * Sem perguntas e OK, nessa ordem. "Sem perguntas": registrado sem .fct.*/
function conferirFormularios(itensField, registro) {
  const porNome = new Map(registro.map((r) => [normalizarNome(r.nome), r]));
  const vistos = new Set();
  const linhas = [];

  for (const form of itensField.filter((f) => !f.archived)) {
    const nome = normalizarNome(form.name);
    const registrado = porNome.get(nome);
    vistos.add(nome);

    const base = {
      Formulario: form.name,
      "Perguntas no Field": form.questionsCount,
      "Perguntas registradas": registrado ? registrado.perguntas : "",
      "Alterado em": dataCurta(form.updatedAt),
    };

    if (!registrado) {
      linhas.push({ ...base, Situacao: "Novo", Observacao: "Formulário não registrado" });
    } else if (registrado.perguntas !== form.questionsCount) {
      linhas.push({
        ...base,
        Situacao: "Alterado",
        Observacao: "Quantidade de perguntas mudou",
      });
    } else if (registrado.atualizadoEm && form.updatedAt > registrado.atualizadoEm) {
      linhas.push({
        ...base,
        Situacao: "Alterado",
        Observacao: "Editado depois do registro",
      });
    } else if (!registrado.perguntasTexto) {
      linhas.push({
        ...base,
        Situacao: "Sem perguntas",
        Observacao: "Sem .fct: só nome e quantidade de perguntas",
      });
    } else {
      linhas.push({ ...base, Situacao: "OK", Observacao: "" });
    }
  }

  for (const registrado of registro) {
    if (vistos.has(normalizarNome(registrado.nome))) continue;
    linhas.push({
      Formulario: registrado.nome,
      "Perguntas no Field": "",
      "Perguntas registradas": registrado.perguntas,
      "Alterado em": "",
      Situacao: "Removido",
      Observacao: "Não está mais ativo no Field",
    });
  }

  const ordem = ["Novo", "Alterado", "Removido", "Sem perguntas", "OK"];
  return linhas.sort((a, b) => ordem.indexOf(a.Situacao) - ordem.indexOf(b.Situacao));
}

/**
 * Monta o registro com os .fct da pasta e os formularios da ultima lista
 * lida do Field. Sem .fct, entra sem o texto das perguntas.*/
function registrar(pastaFct) {
  const lista = fs.existsSync(ULTIMA_LISTA)
    ? JSON.parse(fs.readFileSync(ULTIMA_LISTA, "utf8")).itens
    : [];
  const porNome = new Map(lista.map((f) => [normalizarNome(f.name), f]));

  const comFct = fs
    .readdirSync(pastaFct)
    .filter((nome) => nome.toLowerCase().endsWith(".fct"))
    .map((arquivo) => JSON.parse(fs.readFileSync(path.join(pastaFct, arquivo), "utf8")))
    .map((form) => ({
      nome: form.name,
      perguntas: form.questions.length,
      atualizadoEm: porNome.get(normalizarNome(form.name))?.updatedAt || null,
      perguntasTexto: [...form.questions]
        .sort((a, b) => a.position - b.position)
        .map((q) => q.title.replace(/\s+/g, " ").trim()),
    }));

  const nomesFct = new Set(comFct.map((r) => normalizarNome(r.nome)));
  const soNoField = lista
    .filter((f) => !f.archived && !nomesFct.has(normalizarNome(f.name)))
    .map((f) => ({
      nome: f.name,
      perguntas: f.questionsCount,
      atualizadoEm: f.updatedAt,
      perguntasTexto: null,
    }));

  const registro = [...comFct, ...soNoField].sort((a, b) =>
    a.nome.localeCompare(b.nome),
  );

  fs.writeFileSync(ARQUIVO_REGISTRO, JSON.stringify(registro, null, 2));
  return registro;
}

if (require.main === module) {
  const [comando, pasta] = process.argv.slice(2);
  if (comando !== "registrar") {
    console.error("Uso: node folha-de-pagamento/formularios.js registrar [pasta com os .fct]");
    process.exit(1);
  }
  const registro = registrar(path.resolve(pasta || "Novo_form"));
  const semData = registro.filter((r) => !r.atualizadoEm).length;
  const semPerguntas = registro.filter((r) => !r.perguntasTexto).map((r) => r.nome);
  console.log(`${registro.length} formulario(s) registrado(s) em ${ARQUIVO_REGISTRO}`);
  if (semData) console.log(`${semData} sem data de alteracao (rode um fechamento e registre de novo).`);
  if (semPerguntas.length) console.log(`Sem .fct (so nome e quantidade): ${semPerguntas.join(", ")}`);
}

module.exports = {
  URL_FORMULARIOS,
  carregarRegistro,
  nomesRegistrados,
  normalizarNome,
  salvarListaField,
  conferirFormularios,
  registrar,
};
