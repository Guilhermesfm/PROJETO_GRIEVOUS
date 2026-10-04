// Captura o link publico da O.S. a partir da gaveta aberta.
//
// O link nao esta no DOM: so aparece depois de acionar "Compartilhar", que
// fica na barra da gaveta ou dentro de "Mais opcoes". Dependendo da versao
// do Field, o resultado vem num campo do dialogo ou direto na area de
// transferencia — as duas formas sao cobertas aqui.

const PADRAO_LINK = /https?:\/\/[^\s"'<>]*fieldcontrol\.com\.br[^\s"'<>]*/i;

// Links da propria aplicacao (scripts, estilos) nao sao link de O.S.
const LINK_DE_RECURSO = /\.(js|css|png|jpe?g|svg|woff2?|mp4|ico)(\?|$)/i;

const SELETORES = {
  compartilhar:
    'palantir-button[icon="share-filled"], ' +
    'button:has-text("Compartilhar"), ' +
    'palantir-button:has-text("Compartilhar")',
  maisOpcoes:
    'palantir-button[icon="more-horizontal-outline"], ' +
    'palantir-button:has-text("Mais opções")',
};

/**
 * Clica de verdade, com plano B.
 *
 * O botao Compartilhar fica no topo da gaveta e o Playwright recusava o
 * clique com "Element is outside of the viewport" — nem o force resolve,
 * porque ele ainda precisa da posicao na tela. O click pelo DOM funciona
 * independente disso.
 */
async function clicar(alvo) {
  await alvo.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => {});

  const tentativas = [
    () => alvo.click({ timeout: 4000 }),
    () => alvo.click({ force: true, timeout: 4000 }),
    () => alvo.evaluate((elemento) => elemento.click()),
  ];

  for (const tentar of tentativas) {
    try {
      await tentar();
      return true;
    } catch {
      // Tenta a proxima forma.
    }
  }

  return false;
}

/** O texto parece um link de O.S. do Field? */
function ehLinkDeOs(texto) {
  const achado = String(texto || "").match(PADRAO_LINK);
  if (!achado) return "";
  const url = achado[0];
  return LINK_DE_RECURSO.test(url) ? "" : url;
}

/** Procura o link no dialogo aberto: campo preenchido ou texto na tela. */
async function procurarLinkNaTela(page) {
  const nosCampos = await page
    .evaluate(() => {
      const campos = document.querySelectorAll("input, textarea");
      return Array.from(campos, (campo) => campo.value || "");
    })
    .catch(() => []);

  for (const valor of nosCampos) {
    const link = ehLinkDeOs(valor);
    if (link) return link;
  }

  // Alguns dialogos mostram o link como texto, nao como campo.
  const noTexto = await page
    .evaluate(() => {
      const overlay = document.querySelector(".cdk-overlay-container");
      return overlay ? overlay.innerText || "" : "";
    })
    .catch(() => "");

  return ehLinkDeOs(noTexto);
}

/** Le a area de transferencia, quando o Field copia o link direto. */
async function lerAreaDeTransferencia(page) {
  const texto = await page
    .evaluate(() => navigator.clipboard.readText())
    .catch(() => "");
  return ehLinkDeOs(texto);
}

/**
 * Captura o link da O.S. com a gaveta ja aberta.
 * Devolve "" quando nao consegue — nunca derruba a execucao por causa disso.
 */
async function capturarLinkOs(page) {
  // A permissao evita o navegador recusar a leitura da area de transferencia.
  await page
    .context()
    .grantPermissions(["clipboard-read", "clipboard-write"])
    .catch(() => {});

  // A gaveta termina de montar depois de aparecer: a barra com o
  // Compartilhar chega alguns segundos apos as abas. Checar na hora fazia a
  // captura desistir antes de o botao existir, e o link saia sempre vazio.
  let botao = page.locator(SELETORES.compartilhar).first();
  let visivel = await botao
    .waitFor({ state: "visible", timeout: 15000 })
    .then(() => true)
    .catch(() => false);

  if (!visivel) {
    // Em telas mais estreitas o Compartilhar fica dentro de "Mais opcoes".
    const mais = page.locator(SELETORES.maisOpcoes).first();
    if (await mais.isVisible().catch(() => false)) {
      await clicar(mais);
      botao = page.locator(SELETORES.compartilhar).first();
      visivel = await botao
        .waitFor({ state: "visible", timeout: 5000 })
        .then(() => true)
        .catch(() => false);
    }
  }

  if (!visivel) return "";

  if (!(await clicar(botao))) return "";

  // O dialogo (ou a copia) leva um instante; tenta algumas vezes.
  let link = "";
  const limite = Date.now() + 6000;
  while (!link && Date.now() < limite) {
    await page.waitForTimeout(400);
    link = (await procurarLinkNaTela(page)) || (await lerAreaDeTransferencia(page));
  }

  // Fecha o dialogo para nao atrapalhar a proxima O.S.
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(300);

  return link;
}

module.exports = { capturarLinkOs, ehLinkDeOs, SELETORES, PADRAO_LINK };
