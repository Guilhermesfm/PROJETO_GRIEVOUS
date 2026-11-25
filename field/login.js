// Login no Field: e-mail, Continuar, senha, Continuar.

const { carregarCredenciais } = require("./credenciais.js");

/** Abre a url de login e preenche o e-mail e a senha do .env. */
async function entrarNoField(page, url) {
  const { email, senha } = carregarCredenciais();
  if (!email || !senha) {
    throw new Error(
      "Credenciais nao configuradas. Informe o login do Field na tela inicial.",
    );
  }

  await page.goto(url);

  await page.click('input[name="email"]');
  await page.fill('input[name="email"]', email);
  await page.click('palantir-button:has-text("Continuar")');

  await page.click('input[aria-label="password"]');
  await page.fill('input[aria-label="password"]', senha);
  await page.click('palantir-button:has-text("Continuar")');
}

module.exports = { entrarNoField };
