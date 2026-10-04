// Server simples em Node.js para servir o front do Grievous e rodar as automacoes.
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { carregarCredenciais, salvarCredenciais } = require("../field/credenciais.js");
const {
  runPagamento,
  lerOrdensServico,
  classificarOs,
  descobrirPlanilhaEntrada,
  SAIDA_PADRAO,
} = require("../folha-de-pagamento/Pagamento.js");

// Porta em que o front será exibido. Por padrão só aceita conexões da própria
// máquina, já que a tela de login trafega a senha do Field sem HTTPS.
// Para abrir na rede (HOST=0.0.0.0) seria preciso antes proteger o acesso.
const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || "127.0.0.1";
const PUBLIC_DIR = path.join(__dirname, "public");
// As automacoes e as planilhas ficam na raiz do projeto.
const RAIZ = path.join(__dirname, "..");
const SAIDA_PLACAS = path.join(RAIZ, "placas", "situacoes_placas.xlsx");

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
 * Estado da execução em andamento, para o botão Parar.
 *
 * A automação consulta `parar` entre uma O.S. (ou placa) e a próxima: ela
 * termina a que já começou e encerra, em vez de morrer no meio e perder o
 * que foi lido.
 */
const execucao = { ativa: false, parar: false, tipo: null };

function iniciarExecucao(tipo) {
  execucao.ativa = true;
  execucao.parar = false;
  execucao.tipo = tipo;
}

function encerrarExecucao() {
  execucao.ativa = false;
  execucao.parar = false;
  execucao.tipo = null;
}

/**
 * Recarrega os módulos da automação do disco antes de cada execução.
 *
 * O Node guarda o módulo em cache na primeira vez que ele é exigido; sem isso,
 * editar um seletor não teria efeito até reiniciar o servidor — e a execução
 * rodaria com o código antigo sem avisar.
 */
// Modulos compartilhados pelas duas automacoes.
const MODULOS_COMUNS = [
  "../field/filtro.js",
  "../field/popups.js",
  "../field/gaveta.js",
  "../field/link.js",
  "../field/credenciais.js",
];

function recarregar(modulos, principal) {
  for (const modulo of [...modulos, principal]) {
    delete require.cache[require.resolve(modulo)];
  }
  return require(principal);
}

function carregarAutomacaoPagamento() {
  return recarregar(MODULOS_COMUNS, "../folha-de-pagamento/Pagamento.js");
}

/**
 * A rota de placas usava o runAutomation carregado na inicializacao: o
 * servidor ficava semanas de pe e rodava codigo antigo sem avisar.
 */
function carregarAutomacaoPlacas() {
  return recarregar(MODULOS_COMUNS, "../placas/automation.js");
}

const PASTA_PAGAMENTO = path.join(RAIZ, "folha-de-pagamento");
const EXTENSOES_PLANILHA = [".xlsx", ".xlsm", ".csv"];
const TAMANHO_MAXIMO = 30 * 1024 * 1024; // 30 MB

/**
 * Limpa o nome do arquivo enviado.
 *
 * O nome vem do navegador e não pode ser usado direto: "..\\..\\algo.xlsx"
 * gravaria fora da pasta. Aqui fica só o nome, sem caminho e sem caractere
 * que o Windows recuse.
 */
function nomeSeguro(nomeOriginal) {
  // O front envia o nome codificado, para suportar acento e espaço no header.
  let bruto = String(nomeOriginal || "").trim();
  try {
    bruto = decodeURIComponent(bruto);
  } catch {
    // Nome que não decodifica segue como veio; a limpeza abaixo resolve.
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
        reject(new Error(`Arquivo maior que ${Math.round(limite / 1024 / 1024)} MB.`));
        req.destroy();
        return;
      }
      pedacos.push(pedaco);
    });
    req.on("end", () => resolve(Buffer.concat(pedacos)));
    req.on("error", reject);
  });
}

// Monta o resumo da planilha de entrada do fechamento.
function resumirPlanilha() {
  const entrada = descobrirPlanilhaEntrada();
  if (!entrada) {
    return { ok: false, error: "Nenhuma planilha em folha-de-pagamento/." };
  }

  const ordens = lerOrdensServico(entrada).map((os) => ({
    ...os,
    categoria: classificarOs(os) || "fora do escopo",
  }));

  const porCategoria = ordens.reduce((contagem, os) => {
    contagem[os.categoria] = (contagem[os.categoria] || 0) + 1;
    return contagem;
  }, {});

  return {
    ok: true,
    planilha: path.basename(entrada),
    total: ordens.length,
    porCategoria,
    ordens,
  };
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  // --- Execução: pedido de parada ------------------------------------------
  if (req.method === "POST" && url.pathname === "/api/parar") {
    if (!execucao.ativa) {
      responderJson(res, 200, { ok: true, parando: false, motivo: "nada rodando" });
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
    responderJson(res, 200, {
      ok: true,
      configurado: Boolean(email && senha),
      email,
    });
    return;
  }

  // --- Credenciais: grava EMAIL e PASSWORD no .env -------------------------
  if (req.method === "POST" && url.pathname === "/api/credenciais") {
    try {
      const payload = JSON.parse((await lerCorpo(req)) || "{}");
      salvarCredenciais(
        String(payload.email || "").trim(),
        String(payload.senha || ""),
      );
      console.log("Credenciais do Field atualizadas no .env.");
      responderJson(res, 200, { ok: true, email: String(payload.email).trim() });
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

      // Lê de volta para avisar na hora se a planilha não serve, em vez de
      // deixar o erro aparecer só quando a execução começar.
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

  // --- Fechamento: execucao com progresso em tempo real (SSE) --------------
  if (req.method === "GET" && url.pathname === "/api/pagamento/run") {
    res.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache",
      Connection: "keep-alive",
    });

    const enviar = (evento) => res.write(`data: ${JSON.stringify(evento)}\n\n`);

    // Uma execução por vez.
    //
    // Sem esta trava, duas rodavam em paralelo com dois navegadores no mesmo
    // Field: basta recarregar a página no meio (o EventSource do navegador
    // reconecta sozinho quando o fluxo cai e o servidor começaria tudo de
    // novo) ou abrir a interface em duas abas.
    if (execucao.ativa) {
      enviar({
        tipo: "erro",
        erro:
          "Já existe um fechamento em andamento. Acompanhe a janela que o " +
          "iniciou, ou use o botão Parar antes de começar outro.",
      });
      res.end();
      return;
    }

    // Marca antes de qualquer espera, para não abrir brecha entre a checagem
    // e o início de fato.
    iniciarExecucao("pagamento");

    try {
      const resumo = resumirPlanilha();
      if (!resumo.ok) throw new Error(resumo.error);

      // Permite rodar so um pedaco da planilha (util para testar).
      const limite = Number(url.searchParams.get("limite")) || 0;
      const ordens = limite ? resumo.ordens.slice(0, limite) : resumo.ordens;

      // Pega a versão atual do código, não a que estava em memória.
      const automacao = carregarAutomacaoPagamento();
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
      res.end();
    }
    return;
  }

  // --- Fechamento: download da planilha gerada -----------------------------
  if (req.method === "GET" && url.pathname === "/api/pagamento/download") {
    enviarDownload(res, SAIDA_PADRAO, "pagamento_formularios.xlsx");
    return;
  }

  // --- Placas: roda a automacao -------------------------------------------
  if (req.method === "POST" && url.pathname === "/api/run") {
    if (execucao.ativa) {
      responderJson(res, 409, {
        ok: false,
        error: `Já existe uma execução em andamento (${execucao.tipo}).`,
      });
      return;
    }

    try {
      const payload = JSON.parse((await lerCorpo(req)) || "{}");
      const placas = Array.isArray(payload.placas)
        ? payload.placas.map((placa) => String(placa).trim()).filter(Boolean)
        : [];

      iniciarExecucao("placas");
      const automacao = carregarAutomacaoPlacas();
      const resultado = await automacao.runAutomation(placas, () => execucao.parar);
      responderJson(res, 200, {
        ok: true,
        outputPath: resultado.outputPath,
        resultados: resultado.resultados,
        interrompida: resultado.interrompida,
      });
    } catch (erro) {
      responderJson(res, 500, { ok: false, error: erro.message });
    } finally {
      encerrarExecucao();
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
