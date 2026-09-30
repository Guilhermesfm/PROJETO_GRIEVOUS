// Leitura e gravacao das credenciais do Field no arquivo .env.
const fs = require("node:fs");
const path = require("node:path");
const dotenv = require("dotenv");

const CAMINHO_ENV = path.join(__dirname, ".env");

/**
 * Le o .env do disco a cada chamada, para que a tela de login altere as
 * credenciais sem precisar reiniciar o servidor.
 */
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

/**
 * Grava EMAIL e PASSWORD no .env preservando as demais chaves do arquivo.
 */
function salvarCredenciais(email, senha) {
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
    EMAIL: formatarValor(email),
    PASSWORD: formatarValor(senha),
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
  // mode 0o600: so o dono le o arquivo (ignorado no Windows, util no Linux).
  fs.writeFileSync(CAMINHO_ENV, conteudo.endsWith("\n") ? conteudo : `${conteudo}\n`, {
    encoding: "utf8",
    mode: 0o600,
  });

  // Atualiza o processo em memoria para a proxima execucao.
  process.env.EMAIL = email;
  process.env.PASSWORD = senha;

  return { email };
}

// Enderecos do Field. Ficam no .env porque o Field troca de tela sem aviso e
// nao da para depender de editar o codigo para acompanhar.
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
  carregarUrlsField,
  URLS_PADRAO,
  CAMINHO_ENV,
};
