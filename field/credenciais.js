// Leitura e gravacao das credenciais do Field e do TrackerPrime no .env.
const fs = require("node:fs");
const path = require("node:path");
const dotenv = require("dotenv");

// .env na raiz do projeto.
const CAMINHO_ENV = path.join(__dirname, "..", ".env");

/** Le EMAIL e PASSWORD do .env (relido a cada chamada). */
function carregarCredenciais() {
  dotenv.config({ path: CAMINHO_ENV, override: true, quiet: true });
  return {
    email: process.env.EMAIL || "",
    senha: process.env.PASSWORD || "",
  };
}

/** Valores com espaco ou aspas precisam ser escritos entre aspas. */
function formatarValor(valor) {
  const texto = String(valor ?? "");
  return /[\s"'#]/.test(texto)
    ? `"${texto.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`
    : texto;
}

/** Grava EMAIL e PASSWORD (Field) no .env preservando as demais chaves. */
function salvarCredenciais(email, senha) {
  gravarPar("EMAIL", "PASSWORD", email, senha);
  return { email };
}

/** Le TRACKER_EMAIL e TRACKER_PASSWORD (TrackerPrime/HubPrime) do .env. */
function carregarCredenciaisTracker() {
  dotenv.config({ path: CAMINHO_ENV, override: true, quiet: true });
  return {
    email: process.env.TRACKER_EMAIL || "",
    senha: process.env.TRACKER_PASSWORD || "",
  };
}

/** Grava TRACKER_EMAIL e TRACKER_PASSWORD no .env. */
function salvarCredenciaisTracker(email, senha) {
  gravarPar("TRACKER_EMAIL", "TRACKER_PASSWORD", email, senha);
  return { email };
}

/** Grava um par e-mail/senha no .env preservando as demais chaves. */
function gravarPar(chaveEmail, chaveSenha, email, senha) {
  if (!email || !senha) {
    throw new Error("E-mail e senha sao obrigatorios.");
  }
  if (/[\r\n]/.test(email) || /[\r\n]/.test(senha)) {
    throw new Error("E-mail e senha nao podem conter quebras de linha.");
  }

  const linhasExistentes = fs.existsSync(CAMINHO_ENV)
    ? fs.readFileSync(CAMINHO_ENV, "utf8").split(/\r?\n/)
    : [];

  const novosValores = {
    [chaveEmail]: formatarValor(email),
    [chaveSenha]: formatarValor(senha),
  };
  const jaEscritas = new Set();

  const linhas = linhasExistentes.map((linha) => {
    const chave = (linha.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=/) || [])[1];
    if (chave && novosValores[chave] !== undefined) {
      jaEscritas.add(chave);
      return `${chave}=${novosValores[chave]}`;
    }
    return linha;
  });

  for (const [chave, valor] of Object.entries(novosValores)) {
    if (!jaEscritas.has(chave)) linhas.push(`${chave}=${valor}`);
  }

  const conteudo = linhas.join("\n").replace(/\n{3,}$/, "\n");
  // 0o600: so o dono le o arquivo (ignorado no Windows).
  fs.writeFileSync(CAMINHO_ENV, conteudo.endsWith("\n") ? conteudo : `${conteudo}\n`, {
    encoding: "utf8",
    mode: 0o600,
  });

  // Atualiza o processo em memoria para a proxima execucao.
  process.env[chaveEmail] = email;
  process.env[chaveSenha] = senha;
}

// Enderecos padrao do Field (podem ser trocados no .env).
const URLS_PADRAO = {
  login:
    "https://app.fieldcontrol.com.br/autenticador-v2/#/login?to=:hash:%2Fatividades",
  listagem: "https://app.fieldcontrol.com.br/#/atividades",
};

/** Le as URLs do Field do .env, caindo nos padroes quando nao definidas. */
function carregarUrlsField() {
  dotenv.config({ path: CAMINHO_ENV, override: true, quiet: true });
  return {
    login: process.env.FIELD_URL_LOGIN || URLS_PADRAO.login,
    listagem: process.env.FIELD_URL_LISTAGEM || URLS_PADRAO.listagem,
  };
}

module.exports = {
  carregarCredenciais,
  salvarCredenciais,
  carregarCredenciaisTracker,
  salvarCredenciaisTracker,
  carregarUrlsField,
};
