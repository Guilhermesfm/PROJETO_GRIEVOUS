// Fecha os avisos e modais que o Field abre por cima da tela.
//
// Aparecem em momentos imprevisiveis (logo apos o login, mas tambem no meio da
// navegacao) e bloqueiam qualquer clique, porque o overlay cobre a pagina.
// Por isso a limpeza roda antes das acoes, e nao so uma vez apos entrar.
//
// O Field mistura duas tecnologias: modais antigos em AngularJS (<md-dialog>,
// classe .fc-modal) e componentes novos em CDK (.cdk-overlay-container).
// Os dois precisam ser cobertos.

// Textos dos botoes que fecham esses avisos.
const TEXTOS_FECHAR =
  /^(fechar|mais tarde|agora n[aã]o|entendi|ok|pular|depois|n[aã]o,? obrigad[oa])$/i;

// Onde os avisos vivem. Restrito a dialogos: nao queremos varrer a tela toda
// e sair clicando em botao de acao da listagem.
const CONTAINERS = [
  "md-dialog",
  ".fc-modal",
  ".md-dialog-container",
  ".cdk-overlay-container",
];

// Botoes com texto — o alvo principal ("Fechar", "Mais tarde"...).
const BOTOES = CONTAINERS.map(
  (container) => `${container} button, ${container} palantir-button`,
).join(", ");

// O "X" do canto, que nao tem texto nenhum.
const ICONES_FECHAR = [
  ...CONTAINERS.map((c) => `${c} md-icon[md-svg-src*="close"][role="button"]`),
  ...CONTAINERS.map((c) => `${c} palantir-button[icon="close-outline"]`),
  ...CONTAINERS.map((c) => `${c} [aria-label="Fechar" i]`),
  ...CONTAINERS.map((c) => `${c} [aria-label="close" i]`),
].join(", ");

// Marca os botoes ja acionados nesta passagem, para nao insistir no mesmo
// quando o clique nao faz o aviso sumir — foi o que travava a limpeza antes
// de chegar no segundo modal da fila.
const MARCA = "data-grievous-fechado";

async function limparMarcas(page) {
  await page
    .evaluate((marca) => {
      for (const elemento of document.querySelectorAll(`[${marca}]`)) {
        elemento.removeAttribute(marca);
      }
    }, MARCA)
    .catch(() => {});
}

async function marcar(alvo) {
  await alvo
    .evaluate((elemento, marca) => elemento.setAttribute(marca, "1"), MARCA)
    .catch(() => {});
}

/**
 * Acha o primeiro botao de fechamento visivel ainda nao acionado.
 *
 * A comparacao do texto e feita aqui, e nao no filtro do Playwright, porque o
 * texto real vem com quebra de linha em volta ("\nFechar") e um regex
 * ancorado nunca casaria.
 */
async function acharBotaoPorTexto(page) {
  const botoes = page.locator(BOTOES);
  const total = await botoes.count().catch(() => 0);

  for (let indice = 0; indice < total; indice += 1) {
    const botao = botoes.nth(indice);
    if (await botao.getAttribute(MARCA).catch(() => null)) continue;
    if (!(await botao.isVisible().catch(() => false))) continue;

    const texto = ((await botao.innerText().catch(() => "")) || "").trim();
    if (TEXTOS_FECHAR.test(texto)) return botao;
  }

  return null;
}

/** Mesmo criterio, para o "X" do canto, que nao tem texto. */
async function acharIconeFechar(page) {
  const icones = page.locator(ICONES_FECHAR);
  const total = await icones.count().catch(() => 0);

  for (let indice = 0; indice < total; indice += 1) {
    const icone = icones.nth(indice);
    if (await icone.getAttribute(MARCA).catch(() => null)) continue;
    if (await icone.isVisible().catch(() => false)) return icone;
  }

  return null;
}

/**
 * Clica de verdade, com plano B.
 *
 * O clique normal respeita a acionabilidade; o forcado ignora a checagem mas
 * ainda precisa da posicao na tela; o ultimo recurso dispara o click pelo DOM,
 * que funciona mesmo se o elemento estiver fora da viewport ou coberto.
 * Retorna se algum deles deu certo.
 */
async function clicar(alvo) {
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

/**
 * Fecha os popups visiveis. Retorna quantos foram realmente fechados.
 * Nunca lanca erro: fechar popup e melhor-esforco.
 */
async function fecharPopups(page, tentativas = 6) {
  let fechados = 0;

  // Marcas de uma chamada anterior nao valem mais: o aviso pode ter voltado.
  await limparMarcas(page);

  for (let volta = 0; volta < tentativas; volta += 1) {
    // Botao com texto primeiro; o "X" so quando nao houver.
    const alvo = (await acharBotaoPorTexto(page)) || (await acharIconeFechar(page));
    if (!alvo) break;

    // Marca antes de clicar: se o aviso nao sumir, a proxima volta tenta
    // outro botao em vez de insistir neste.
    await marcar(alvo);

    if (await clicar(alvo)) {
      fechados += 1;
      // Os avisos vem em fila e com animacao; da tempo do proximo aparecer.
      await page.waitForTimeout(1200);
    }
  }

  return fechados;
}

/** Sobrou algum overlay bloqueando a tela? */
async function temOverlayBloqueando(page) {
  return page
    .locator(
      "md-dialog:visible, .md-dialog-container:visible, " +
        ".cdk-overlay-backdrop:visible, md-backdrop:visible",
    )
    .first()
    .isVisible()
    .catch(() => false);
}

module.exports = {
  fecharPopups,
  temOverlayBloqueando,
  TEXTOS_FECHAR,
  BOTOES,
  ICONES_FECHAR,
};
