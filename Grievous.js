/**
 * Grievous.js - menu principal.
 *
 * Uso:
 *   node Grievous.js              -> menu interativo
 *   node Grievous.js pagamento    -> roda o fechamento da folha de pagamento
 *   node Grievous.js placas       -> roda a automacao de placas
 *   node Grievous.js web          -> sobe a interface web
 */

const readline = require("node:readline");
const path = require("node:path");
const { carregarCredenciais, salvarCredenciais } = require("./variaveis.js");

const OPCOES = [
  {
    chave: "pagamento",
    titulo: "Fechamento (folha de pagamento)",
    descricao: "Le a planilha de folha-de-pagamento/ e extrai os formularios das O.S.",
    executar: executarPagamento,
  },
  {
    chave: "placas",
    titulo: "Automacao de placas",
    descricao: "Consulta a situacao de cada placa no Field.",
    executar: executarPlacas,
  },
  {
    chave: "web",
    titulo: "Interface web",
    descricao: "Sobe o front em http://localhost:3000.",
    executar: executarWeb,
  },
  {
    chave: "login",
    titulo: "Credenciais do Field",
    descricao: "Grava e-mail e senha no arquivo .env.",
    executar: () => pedirCredenciais(true),
  },
];

/**
 * Pede o login do Field e grava no .env.
 * Com forcar = false, so pergunta se ainda nao houver credenciais.
 */
async function pedirCredenciais(forcar = false) {
  const atual = carregarCredenciais();
  if (!forcar && atual.email && atual.senha) return true;

  console.log("\n--- Credenciais do Field ---");
  if (atual.email) console.log(`Conta atual: ${atual.email}`);

  const email = (await perguntar("E-mail: ")) || atual.email;
  const senha = await perguntar("Senha: ", true);

  if (!email || !senha) {
    console.error("E-mail e senha sao obrigatorios.");
    return false;
  }

  salvarCredenciais(email, senha);
  console.log("Credenciais gravadas no .env.\n");
  return true;
}

async function executarPagamento() {
  const {
    runPagamento,
    lerOrdensServico,
    descobrirPlanilhaEntrada,
  } = require("./folha-de-pagamento/Pagamento.js");

  const entrada = descobrirPlanilhaEntrada();
  if (!entrada) {
    console.error(
      "Nenhuma planilha encontrada em folha-de-pagamento/. Coloque o arquivo la e rode de novo.",
    );
    return;
  }

  const ordens = lerOrdensServico(entrada);
  console.log(`${ordens.length} O.S. lidas de ${path.basename(entrada)}`);
  await runPagamento(ordens);
}

async function executarPlacas() {
  const { runAutomation } = require("./placas/automation.js");
  const placas = await perguntar(
    "Placas (separadas por virgula ou espaco): ",
  ).then((resposta) =>
    resposta
      .split(/[,\s]+/)
      .map((placa) => placa.trim())
      .filter(Boolean),
  );

  if (!placas.length) {
    console.error("Nenhuma placa informada.");
    return;
  }

  await runAutomation(placas);
}

async function executarWeb() {
  require("./server.js");
}

/** Pergunta no terminal. Com oculto = true, nao ecoa o que for digitado. */
function perguntar(texto, oculto = false) {
  const io = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    terminal: true,
  });

  if (oculto) {
    // Esconde a senha substituindo a escrita no terminal.
    io._writeToOutput = (trecho) => {
      io.output.write(trecho.startsWith(texto) ? texto : "");
    };
  }

  return new Promise((resolve) => {
    io.question(texto, (resposta) => {
      io.close();
      if (oculto) process.stdout.write("\n");
      resolve(resposta.trim());
    });
  });
}

async function menu() {
  console.log("\n=== GRIEVOUS ===\n");
  OPCOES.forEach((opcao, indice) => {
    console.log(`  ${indice + 1}) ${opcao.titulo}`);
    console.log(`     ${opcao.descricao}`);
  });
  console.log("  0) Sair\n");

  const escolha = await perguntar("Escolha uma opcao: ");
  if (escolha === "0" || normalizarEscolha(escolha) === "sair") return;

  const opcao =
    OPCOES[Number(escolha) - 1] ||
    OPCOES.find((item) => item.chave === normalizarEscolha(escolha));

  if (!opcao) {
    console.error("Opcao invalida.");
    return menu();
  }

  // As automacoes precisam do login do Field; a web pede na propria tela.
  if (["pagamento", "placas"].includes(opcao.chave)) {
    if (!(await pedirCredenciais())) return;
  }

  await opcao.executar();
}

function normalizarEscolha(valor) {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .trim();
}

async function main() {
  const argumento = normalizarEscolha(process.argv[2]);
  if (!argumento) return menu();

  const opcao = OPCOES.find((item) => item.chave === argumento);
  if (!opcao) {
    console.error(
      `Opcao "${process.argv[2]}" desconhecida. Use: ${OPCOES.map((o) => o.chave).join(", ")}`,
    );
    process.exit(1);
  }

  await opcao.executar();
}

if (require.main === module) {
  main().catch((erro) => {
    console.error("Erro na execucao:", erro);
    process.exit(1);
  });
}

module.exports = { OPCOES };
