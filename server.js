// Server simples em Node.js para servir o front e receber as placas via API.
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { runAutomation } = require("./files/automation.js");

// Porta em que o front será exibido.
const PORT = 3000;
const PUBLIC_DIR = path.join(__dirname, "public");

// Mapeia a extensão do arquivo para o tipo correto do HTTP.
const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
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

// Cria o servidor HTTP e trata rota do front + rota da API.
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  // Rota da API: recebe as placas do front e roda a automação.
  if (req.method === "POST" && url.pathname === "/api/run") {
    let body = "";

    req.on("data", (chunk) => {
      body += chunk;
    });

    req.on("end", async () => {
      try {
        // Converte as placas enviadas pelo front em uma lista limpa.
        const payload = JSON.parse(body || "{}");
        const placas = Array.isArray(payload.placas)
          ? payload.placas.map((placa) => String(placa).trim()).filter(Boolean)
          : [];

        // Executa a automação com a lista recebida.
        const resultado = await runAutomation(placas);

        // Responde ao front com o caminho do arquivo gerado e os resultados.
        res.writeHead(200, {
          "Content-Type": "application/json; charset=utf-8",
        });
        res.end(
          JSON.stringify({
            ok: true,
            outputPath: resultado.outputPath,
            resultados: resultado.resultados,
          }),
        );
      } catch (error) {
        // Se der erro, devolve mensagem clara para o front exibir.
        res.writeHead(500, {
          "Content-Type": "application/json; charset=utf-8",
        });
        res.end(JSON.stringify({ ok: false, error: error.message }));
      }
    });
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

// Inicia o servidor.
server.listen(PORT, () => {
  console.log(`Front rodando em http://localhost:${PORT}`);
});
