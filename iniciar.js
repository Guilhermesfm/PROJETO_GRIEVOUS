/**
 * iniciar.js - sobe o Grievous e abre o navegador na tela de login.
 *
 * E o que o executavel chama. A logica ficou aqui, em Node, e nao no .bat,
 * porque no .bat um erro do servidor sumia junto com a janela: a tela abria,
 * mas quem respondia era outra coisa (ou nada), e salvar as credenciais
 * falhava sem dizer por que.
 */

const http = require("node:http");
const net = require("node:net");
const { spawn } = require("node:child_process");

const PORTA_INICIAL = Number(process.env.PORT) || 3000;
// Quantas portas tentar antes de desistir, para rodar mais de uma copia na
// mesma maquina sem precisar configurar nada.
const PORTAS_A_TENTAR = 10;

let PORTA = PORTA_INICIAL;
let ENDERECO = `http://localhost:${PORTA}/`;

function usarPorta(porta) {
  PORTA = porta;
  ENDERECO = `http://localhost:${porta}/`;
  // O servidor le a porta daqui.
  process.env.PORT = String(porta);
}

/** A porta ja esta ocupada por alguem? */
function portaOcupada(porta) {
  return new Promise((resolve) => {
    const teste = net
      .createServer()
      .once("error", (erro) => resolve(erro.code === "EADDRINUSE"))
      .once("listening", () => teste.close(() => resolve(false)))
      .listen(porta, "127.0.0.1");
  });
}

/** Quem responde na porta e mesmo o Grievous? */
function grievousRespondendo() {
  return new Promise((resolve) => {
    const pedido = http.get(
      `${ENDERECO}api/status`,
      { timeout: 2000 },
      (resposta) => {
        let corpo = "";
        resposta.on("data", (pedaco) => (corpo += pedaco));
        resposta.on("end", () => {
          try {
            resolve(Boolean(JSON.parse(corpo).ok));
          } catch {
            resolve(false);
          }
        });
      },
    );
    pedido.on("error", () => resolve(false));
    pedido.on("timeout", () => {
      pedido.destroy();
      resolve(false);
    });
  });
}

/** Espera o servidor comecar a responder. */
async function esperarServidor(tentativas = 40) {
  for (let i = 0; i < tentativas; i += 1) {
    if (await grievousRespondendo()) return true;
    await new Promise((r) => setTimeout(r, 500));
  }
  return false;
}

function abrirNavegador(url) {
  // start e um comando interno do cmd; o primeiro "" e o titulo da janela.
  if (process.platform === "win32") {
    spawn("cmd", ["/c", "start", "", url], { detached: true, stdio: "ignore" });
  } else if (process.platform === "darwin") {
    spawn("open", [url], { detached: true, stdio: "ignore" });
  } else {
    spawn("xdg-open", [url], { detached: true, stdio: "ignore" });
  }
}

/**
 * A pasta esta dentro de um servico de sincronizacao?
 *
 * Rodar o projeto de dentro do OneDrive em mais de uma maquina e o caminho
 * mais curto para executar codigo velho: a sincronizacao tem atraso e cria
 * copias duplicadas ("arquivo (1).xlsx"). Cada maquina deve ter o seu clone.
 */
function avisarSePastaSincronizada() {
  if (!/OneDrive|Dropbox|Google Drive|iCloud/i.test(__dirname)) return;

  console.log("Atencao: o projeto esta numa pasta sincronizada.");
  console.log(`  ${__dirname}`);
  console.log("  Em mais de uma maquina isso faz rodar codigo desatualizado.");
  console.log("  Prefira um clone do git fora da pasta sincronizada e use");
  console.log("  'git pull' para atualizar cada maquina.\n");
}

async function main() {
  console.log("=== GRIEVOUS ===\n");
  avisarSePastaSincronizada();

  // Procura uma porta livre a partir da inicial. Antes o programa desistia
  // quando a 3000 estava ocupada; abrir uma segunda copia exigia configurar
  // PORT na mao.
  let encontrou = false;

  for (let tentativa = 0; tentativa < PORTAS_A_TENTAR; tentativa += 1) {
    usarPorta(PORTA_INICIAL + tentativa);

    if (!(await portaOcupada(PORTA))) {
      encontrou = true;
      break;
    }

    // Se quem esta na porta e o proprio Grievous, reaproveita em vez de subir
    // outro, que morreria com "endereco em uso" sem ninguem ver.
    if (await grievousRespondendo()) {
      console.log(`Ja havia um Grievous rodando em ${ENDERECO}`);
      console.log("Abrindo o navegador na instancia existente.\n");
      abrirNavegador(ENDERECO);
      return;
    }

    console.log(`Porta ${PORTA} ocupada por outro programa, tentando a proxima...`);
  }

  if (!encontrou) {
    console.error(
      `Nenhuma porta livre entre ${PORTA_INICIAL} e ` +
        `${PORTA_INICIAL + PORTAS_A_TENTAR - 1}.\n`,
    );
    process.exitCode = 1;
    return;
  }

  console.log("Subindo o servidor...");
  require("./web/servidor.js");

  if (!(await esperarServidor())) {
    console.error("O servidor nao respondeu a tempo. Veja o erro acima.");
    process.exitCode = 1;
    return;
  }

  console.log(`Pronto: ${ENDERECO}`);
  console.log("Feche esta janela para encerrar o Grievous.\n");
  abrirNavegador(ENDERECO);
}

main().catch((erro) => {
  console.error("Falha ao iniciar:", erro.message);
  process.exitCode = 1;
});
