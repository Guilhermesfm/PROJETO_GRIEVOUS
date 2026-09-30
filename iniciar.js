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

const PORTA = Number(process.env.PORT) || 3000;
const ENDERECO = `http://localhost:${PORTA}/`;

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

async function main() {
  console.log("=== GRIEVOUS ===\n");

  if (await portaOcupada(PORTA)) {
    // Ja tem algo na porta. Se for o proprio Grievous, reaproveita em vez de
    // subir outro que morreria com "endereco em uso" sem ninguem ver.
    if (await grievousRespondendo()) {
      console.log(`Ja havia um Grievous rodando em ${ENDERECO}`);
      console.log("Abrindo o navegador na instancia existente.\n");
      abrirNavegador(ENDERECO);
      return;
    }

    console.error(
      `A porta ${PORTA} esta ocupada por outro programa.\n` +
        `Feche-o, ou rode com outra porta:  set PORT=3001 && node iniciar.js\n`,
    );
    process.exitCode = 1;
    return;
  }

  console.log("Subindo o servidor...");
  require("./server.js");

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
