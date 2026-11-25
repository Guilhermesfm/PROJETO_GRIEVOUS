// Server simples em Node.js para servir o front do Grievous e rodar as automacoes.
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const {
  carregarCredenciais,
  salvarCredenciais,
  carregarCredenciaisTracker,
  salvarCredenciaisTracker,
} = require("../field/credenciais.js");
const { SAIDA_PADRAO } = require("../folha-de-pagamento/Pagamento.js");
const { SAIDA_PLACAS } = require("../placas/automation.js");

// Porta do front. So aceita conexoes locais: a senha do Field trafega sem HTTPS.
const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || "127.0.0.1";
const PUBLIC_DIR = path.join(__dirname, "public");
// As automacoes e as planilhas ficam na raiz do projeto.
const RAIZ = path.join(__dirname, "..");

// Mapeia a extensão do arquivo para o tipo correto do HTTP.
const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

// Lê um arquivo do diretório público e envia para o navegador.
function serveStaticFile(res, filePath) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Arquivo não encontrado.");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME_TYPES[ext] || "application/octet-stream",
    });
    res.end(content);
  });
}

// Responde JSON com o status informado.
function responderJson(res, status, corpo) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(corpo));
}

// Lê o corpo de uma requisição POST.
function lerCorpo(req) {
  return new Promise((resolve) => {
    let corpo = "";
    req.on("data", (pedaco) => {
      corpo += pedaco;
    });
    req.on("end", () => resolve(corpo));
  });
}

// Envia um arquivo para download.
function enviarDownload(res, caminho, nome) {
  if (!fs.existsSync(caminho)) {
    responderJson(res, 404, { ok: false, error: "Planilha ainda não gerada." });
    return;
  }
  res.writeHead(200, {
    "Content-Type":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": `attachment; filename="${nome}"`,
  });
  fs.createReadStream(caminho).pipe(res);
}

/**
 * Estado da execução em andamento, para o botão Parar. */
const execucao = { ativa: false, parar: false, tipo: null };

/** Marca que uma automacao ("pagamento" ou "placas") comecou a rodar. */
function iniciarExecucao(tipo) {
  execucao.ativa = true;
  execucao.parar = false;
  execucao.tipo = tipo;
}

/** Libera o servidor para a proxima execucao e zera o pedido de parada. */
function encerrarExecucao() {
  execucao.ativa = false;
  execucao.parar = false;
  execucao.tipo = null;
}

/**
 * Progresso de uma automacao, compartilhado entre as abas.*/
function criarProgresso() {
  const progresso = {
    // Eventos da execucao atual, para repassar a quem chegar depois.
    eventos: [],
    // Conexoes SSE abertas (res) acompanhando a execucao.
    ouvintes: new Set(),
    // Planilha gerada pela ultima execucao (null = a padrao).
    ultimaSaida: null,

    /** Grava o evento no historico e repassa para todas as abas. */
    emitir(evento) {
      if (evento.tipo === "fim" && evento.outputPath) {
        progresso.ultimaSaida = evento.outputPath;
      }
      progresso.eventos.push(evento);
      const linha = `data: ${JSON.stringify(evento)}\n\n`;
      for (const res of progresso.ouvintes) res.write(linha);
    },

    /** Conecta a aba ao progresso (fechar a aba so desconecta). */
    acompanhar(req, res) {
      for (const evento of progresso.eventos) {
        res.write(`data: ${JSON.stringify(evento)}\n\n`);
      }
      progresso.ouvintes.add(res);
      req.on("close", () => progresso.ouvintes.delete(res));
    },

    /** Fim da execucao: encerra as conexoes de todas as abas. */
    encerrar() {
      for (const res of progresso.ouvintes) res.end();
      progresso.ouvintes.clear();
    },
  };
  return progresso;
}

const progressoPagamento = criarProgresso();
const progressoPlacas = criarProgresso();

/** Abre a resposta como fluxo de eventos (SSE). */
function abrirFluxo(res) {
  res.writeHead(200, {
    "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
  });
}

/** Envia um unico evento e fecha o fluxo. */
function responderEvento(res, evento) {
  res.write(`data: ${JSON.stringify(evento)}\n\n`);
  res.end();
}

// Modulos compartilhados pelas duas automacoes.
const MODULOS_COMUNS = [
  "../field/filtro.js",
  "../field/popups.js",
  "../field/gaveta.js",
  "../field/link.js",
  "../field/credenciais.js",
  "../field/login.js",
  "../folha-de-pagamento/formularios.js",
  "../folha-de-pagamento/hubprime.js",
];

/**
 * Recarrega os módulos da automação do disco antes de cada execução.
 */
function recarregar(modulos, principal) {
  for (const modulo of [...modulos, principal]) {
    delete require.cache[require.resolve(modulo)];
  }
  return require(principal);
}

/** Pagamento.js recem-lido do disco, com os modulos do Field. */
function carregarAutomacaoPagamento() {
  return recarregar(MODULOS_COMUNS, "../folha-de-pagamento/Pagamento.js");
}

function carregarAutomacaoPlacas() {
  return recarregar(MODULOS_COMUNS, "../placas/automation.js");
}

const PASTA_PAGAMENTO = path.join(RAIZ, "folha-de-pagamento");
const EXTENSOES_PLANILHA = [".xlsx", ".xlsm", ".csv"];
const TAMANHO_MAXIMO = 30 * 1024 * 1024; // 30 MB

/* Limpa o nome do arquivo enviado.*/
function nomeSeguro(nomeOriginal) {
  // O front envia o nome codificado, para suportar acento e espaço no header.
  let bruto = String(nomeOriginal || "").trim();
  try {
    bruto = decodeURIComponent(bruto);
  } catch {
    // Nome que não decodifica segue como veio
  }

  const apenasNome = path.basename(bruto);
  const limpo = apenasNome.replace(/[<>:"/\\|?*\x00-\x1f]/g, "_");
  const extensao = path.extname(limpo).toLowerCase();

  if (!EXTENSOES_PLANILHA.includes(extensao)) {
    throw new Error(
      `Formato não aceito (${extensao || "sem extensão"}). Envie .xlsx, .xlsm ou .csv.`,
    );
  }

  return limpo || `planilha${extensao}`;
}

// Lê o corpo binário da requisição, com limite de tamanho.
function lerCorpoBinario(req, limite = TAMANHO_MAXIMO) {
  return new Promise((resolve, reject) => {
    const pedacos = [];
    let total = 0;

    req.on("data", (pedaco) => {
      total += pedaco.length;
      if (total > limite) {
        reject(
          new Error(
            `Arquivo maior que ${Math.round(limite / 1024 / 1024)} MB.`,
          ),
        );
        req.destroy();
        return;
      }
      pedacos.push(pedaco);
    });
    req.on("end", () => resolve(Buffer.concat(pedacos)));
    req.on("error", reject);
  });
}

// Monta o resumo da planilha de entrada com as regras atuais do Pagamento.js.
// Recebe o Pagamento ja carregado, para nao ler do disco duas vezes.
function resumirPlanilha(automacao = carregarAutomacaoPagamento()) {
  const { descobrirPlanilhaEntrada, lerOrdensServico, classificarOs } = automacao;
  const entrada = descobrirPlanilhaEntrada();
  if (!entrada) {
    return { ok: false, error: "Nenhuma planilha em folha-de-pagamento/." };
  }

  const ordens = lerOrdensServico(entrada).map((os) => ({
    ...os,
    categoria: classificarOs(os) || "fora do escopo",
  }));

  const contar = (campo) =>
    ordens.reduce((contagem, os) => {
      contagem[os[campo]] = (contagem[os[campo]] || 0) + 1;
      return contagem;
    }, {});

  return {
    ok: true,
    planilha: path.basename(entrada),
    total: ordens.length,
    porCategoria: contar("categoria"),
    // Field e HubPrime.
    porOrigem: contar("origem"),
    ordens,
  };
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  // --- Execução: pedido de parada ------------------------------------------
  if (req.method === "POST" && url.pathname === "/api/parar") {
    if (!execucao.ativa) {
      responderJson(res, 200, {
        ok: true,
        parando: false,
        motivo: "nada rodando",
      });
      return;
    }
    execucao.parar = true;
    console.log(`Parada solicitada (${execucao.tipo}).`);
    responderJson(res, 200, { ok: true, parando: true, tipo: execucao.tipo });
    return;
  }

  // --- Execução: status ----------------------------------------------------
  if (req.method === "GET" && url.pathname === "/api/status") {
    responderJson(res, 200, { ok: true, ...execucao });
    return;
  }

  // --- Credenciais: estado atual (a senha nunca sai daqui) -----------------
  if (req.method === "GET" && url.pathname === "/api/credenciais") {
    const { email, senha } = carregarCredenciais();
    const tracker = carregarCredenciaisTracker();
    responderJson(res, 200, {
      ok: true,
      configurado: Boolean(email && senha),
      email,
      trackerConfigurado: Boolean(tracker.email && tracker.senha),
      trackerEmail: tracker.email,
    });
    return;
  }

  // --- Credenciais: grava Field e/ou TrackerPrime no .env ------------------
  if (req.method === "POST" && url.pathname === "/api/credenciais") {
    try {
      const payload = JSON.parse((await lerCorpo(req)) || "{}");
      const email = String(payload.email || "").trim();
      const senha = String(payload.senha || "");
      const trackerEmail = String(payload.trackerEmail || "").trim();
      const trackerSenha = String(payload.trackerSenha || "");
      if (!(email && senha) && !(trackerEmail && trackerSenha)) {
        throw new Error("Informe e-mail e senha do Field ou do TrackerPrime.");
      }

      // Campo em branco mantem o que ja esta no .env.
      if (email && senha) {
        salvarCredenciais(email, senha);
        console.log("Credenciais do Field atualizadas no .env.");
      }
      if (trackerEmail && trackerSenha) {
        salvarCredenciaisTracker(trackerEmail, trackerSenha);
        console.log("Credenciais do TrackerPrime atualizadas no .env.");
      }
      responderJson(res, 200, {
        ok: true,
        email: carregarCredenciais().email,
        trackerEmail: carregarCredenciaisTracker().email,
      });
    } catch (erro) {
      responderJson(res, 400, { ok: false, error: erro.message });
    }
    return;
  }

  // --- Fechamento: envio da planilha ---------------------------------------
  if (req.method === "POST" && url.pathname === "/api/pagamento/planilha") {
    if (execucao.ativa) {
      responderJson(res, 409, {
        ok: false,
        error: "Há uma execução em andamento. Pare antes de trocar a planilha.",
      });
      return;
    }

    try {
      const nome = nomeSeguro(req.headers["x-nome-arquivo"]);
      const conteudo = await lerCorpoBinario(req);

      if (!conteudo.length) throw new Error("Arquivo vazio.");

      const destino = path.join(PASTA_PAGAMENTO, nome);
      fs.mkdirSync(PASTA_PAGAMENTO, { recursive: true });
      fs.writeFileSync(destino, conteudo);

      // Le de volta para validar a planilha na hora.
      const resumo = resumirPlanilha();
      if (!resumo.ok) throw new Error(resumo.error);

      console.log(`Planilha recebida: ${nome} (${conteudo.length} bytes)`);
      responderJson(res, 200, { ...resumo, enviada: nome });
    } catch (erro) {
      responderJson(res, 400, { ok: false, error: erro.message });
    }
    return;
  }

  // --- Fechamento: resumo da planilha de entrada ---------------------------
  if (req.method === "GET" && url.pathname === "/api/pagamento/planilha") {
    try {
      responderJson(res, 200, resumirPlanilha());
    } catch (erro) {
      responderJson(res, 500, { ok: false, error: erro.message });
    }
    return;
  }

  // --- Fechamento: acompanhar o que ja esta rodando (SSE) ------------------
  // Ao abrir a tela: conecta ao fechamento em andamento, se houver.
  if (req.method === "GET" && url.pathname === "/api/pagamento/acompanhar") {
    abrirFluxo(res);

    if (execucao.ativa && execucao.tipo === "pagamento") {
      progressoPagamento.acompanhar(req, res);
    } else {
      responderEvento(res, { tipo: "sem-execucao" });
    }
    return;
  }

  // --- Fechamento: execucao com progresso em tempo real (SSE) --------------
  if (req.method === "GET" && url.pathname === "/api/pagamento/run") {
    abrirFluxo(res);

    // Uma execução por vez.
    if (execucao.ativa) {
      // Ja e um fechamento: esta aba so acompanha.
      if (execucao.tipo === "pagamento") {
        progressoPagamento.acompanhar(req, res);
        return;
      }
      responderEvento(res, {
        tipo: "erro",
        erro: `Já existe uma execução em andamento (${execucao.tipo}). Pare-a antes de começar a análise do Financeiro Inoprime.`,
      });
      return;
    }

    // Marca antes de qualquer await, para nao rodar duas ao mesmo tempo.
    iniciarExecucao("pagamento");

    // Execucao nova: historico zerado, e esta aba e a primeira a acompanhar.
    progressoPagamento.eventos = [];
    progressoPagamento.acompanhar(req, res);
    const enviar = progressoPagamento.emitir;

    try {
      // Le o codigo atual do disco.
      const automacao = carregarAutomacaoPagamento();
      const resumo = resumirPlanilha(automacao);
      if (!resumo.ok) throw new Error(resumo.error);

      // Permite rodar so um pedaco da planilha (util para testar).
      const limite = Number(url.searchParams.get("limite")) || 0;
      const ordens = limite ? resumo.ordens.slice(0, limite) : resumo.ordens;

      await automacao.runPagamento(
        ordens,
        automacao.SAIDA_PADRAO,
        enviar,
        () => execucao.parar,
      );
    } catch (erro) {
      enviar({ tipo: "erro", erro: erro.message });
    } finally {
      encerrarExecucao();
      progressoPagamento.encerrar();
    }
    return;
  }

  // --- Fechamento: download da planilha gerada -----------------------------
  if (req.method === "GET" && url.pathname === "/api/pagamento/download") {
    const saida =
      progressoPagamento.ultimaSaida &&
      fs.existsSync(progressoPagamento.ultimaSaida)
        ? progressoPagamento.ultimaSaida
        : SAIDA_PADRAO;
    enviarDownload(res, saida, path.basename(saida));
    return;
  }

  // --- Placas: acompanhar a consulta em andamento (SSE) --------------------
  if (req.method === "GET" && url.pathname === "/api/placas/acompanhar") {
    abrirFluxo(res);

    if (execucao.ativa && execucao.tipo === "placas") {
      progressoPlacas.acompanhar(req, res);
    } else {
      responderEvento(res, { tipo: "sem-execucao" });
    }
    return;
  }

  // --- Placas: consulta com progresso em tempo real (SSE) ------------------
  if (req.method === "GET" && url.pathname === "/api/placas/run") {
    abrirFluxo(res);

    if (execucao.ativa) {
      // Ja e uma consulta de placas: esta aba so acompanha.
      if (execucao.tipo === "placas") {
        progressoPlacas.acompanhar(req, res);
        return;
      }
      responderEvento(res, {
        tipo: "erro",
        erro: `Já existe uma execução em andamento (${execucao.tipo}). Pare-a antes de consultar as placas.`,
      });
      return;
    }

    const placas = String(url.searchParams.get("placas") || "")
      .split(/[\n,;]+/)
      .map((placa) => placa.trim())
      .filter(Boolean);
    if (!placas.length) {
      responderEvento(res, { tipo: "erro", erro: "Informe ao menos uma placa." });
      return;
    }

    iniciarExecucao("placas");
    progressoPlacas.eventos = [];
    progressoPlacas.acompanhar(req, res);

    try {
      const automacao = carregarAutomacaoPlacas();
      await automacao.runAutomation(
        placas,
        () => execucao.parar,
        progressoPlacas.emitir,
      );
    } catch (erro) {
      progressoPlacas.emitir({ tipo: "erro", erro: erro.message });
    } finally {
      encerrarExecucao();
      progressoPlacas.encerrar();
    }
    return;
  }

  // --- Placas: download da planilha gerada ---------------------------------
  if (req.method === "GET" && url.pathname === "/api/placas/download") {
    enviarDownload(res, SAIDA_PLACAS, "situacoes_placas.xlsx");
    return;
  }

  // Rota de arquivos estáticos: serve HTML, CSS e JS do front.
  const relativePath =
    url.pathname === "/" ? "index.html" : url.pathname.replace(/^\/+/, "");
  const filePath = path.join(PUBLIC_DIR, relativePath);
  const resolvedPublicDir = path.resolve(PUBLIC_DIR);
  const resolvedFilePath = path.resolve(filePath);
  const dentroDoPublico =
    resolvedFilePath === resolvedPublicDir ||
    resolvedFilePath.startsWith(resolvedPublicDir + path.sep);

  if (dentroDoPublico) {
    serveStaticFile(res, resolvedFilePath);
    return;
  }

  res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
  res.end("Acesso negado.");
});

server.listen(PORT, HOST, () => {
  console.log(`Grievous rodando em http://localhost:${PORT}`);
});
