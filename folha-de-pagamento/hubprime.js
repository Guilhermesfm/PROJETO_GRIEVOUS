/**
 * hubprime.js - O.S. do HubPrime, lidas pelo TrackerPrime.
 *
 * Faz login no TrackerPrime, lista as O.S. de origem HubPrime em
 * "Serviços internos" e le o "Relatório de Atendimento" de cada uma (link do
 * olho). As respostas viram um formulario no mesmo formato do Field, para
 * passar pelas mesmas regras do Pagamento.js.
 */

const { carregarCredenciaisTracker } = require("../field/credenciais.js");

const URL_SERVICOS = "https://trackerprime.inoprime.com.br/servicos-internos";

/** Faz login no TrackerPrime e espera a pagina de servicos internos. */
async function loginTracker(page) {
  const { email, senha } = carregarCredenciaisTracker();
  if (!email || !senha) {
    throw new Error(
      "Credenciais do TrackerPrime nao configuradas. Informe na tela de login do Grievous.",
    );
  }

  await page.goto(URL_SERVICOS, { waitUntil: "domcontentloaded" });
  if (page.url().includes("/login")) {
    await page.fill('input[type="email"], input[name*="mail" i]', email);
    await page.fill('input[type="password"]', senha);
    await page.press('input[type="password"]', "Enter");
    await page.waitForURL(/servicos-internos|dashboard|home/, { timeout: 30000 }).catch(() => {});
    if (page.url().includes("/login")) {
      throw new Error("Login no TrackerPrime nao concluido. Confira o e-mail e a senha.");
    }
    if (!page.url().includes("servicos-internos")) {
      await page.goto(URL_SERVICOS, { waitUntil: "domcontentloaded" });
    }
  }
  await page.waitForSelector("#pesquisar-ordens", { timeout: 30000 });
}

/**
 * O.S. de origem HubPrime, por ID: { id, identificador, tipo, status, link }.*/
async function listarOsHubPrime(page) {
  // Origem = HubPrime (select2 por cima do <select id="origem">).
  const selecionou = await page.evaluate(() => {
    const select = document.querySelector("select#origem");
    const opcao = select && [...select.options].find((o) => /hubprime/i.test(o.textContent));
    if (!opcao) return false;
    if (window.jQuery) window.jQuery(select).val(opcao.value).trigger("change");
    else {
      select.value = opcao.value;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
    return true;
  });
  if (!selecionou) {
    throw new Error('Origem "HubPrime" nao encontrada no filtro do TrackerPrime.');
  }
  await page.click("#pesquisar-ordens");
  await esperarTabela(page);

  // A tabela mostra 10 por pagina: pede 1000.
  await page.selectOption('select[name="ordens_length"]', "1000");
  await esperarTabela(page);

  const linhas = await page.locator("#ordens tbody tr").evaluateAll((trs) =>
    trs.map((tr) => {
      const celulas = [...tr.querySelectorAll("td")].map((td) => td.innerText.trim());
      const olho = tr.querySelector("i.fa-eye");
      const link = olho && olho.closest("a") ? olho.closest("a").href : "";
      return { celulas, link };
    }),
  );

  const porId = new Map();
  for (const { celulas, link } of linhas) {
    if (!link) continue;
    porId.set(String(celulas[1]), {
      id: String(celulas[1]),
      identificador: celulas[2],
      tipo: celulas[4],
      status: celulas[7],
      link,
    });
  }
  return porId;
}

/** Espera a tabela parar de mudar (ate 20s). */
async function esperarTabela(page) {
  let anterior = -1;
  const limite = Date.now() + 20000;
  while (Date.now() < limite) {
    await page.waitForTimeout(1000);
    const atual = await page.locator("#ordens tbody tr").count();
    const processando = await page
      .locator("#ordens_processing")
      .isVisible()
      .catch(() => false);
    if (atual === anterior && !processando) return;
    anterior = atual;
  }
}

/**
 * Le o relatorio de atendimento da O.S.
 * Retorna { tipo, concluidoEm ("dd/mm/aaaa hh:mm"), perguntas }.*/
async function lerRelatorio(page, link) {
  await page.goto(link, { waitUntil: "domcontentloaded" });
  await page
    .getByText(/Relat[oó]rio de Atendimento/i)
    .first()
    .waitFor({ timeout: 20000 })
    .catch(() => {});
  const texto = await page.locator("body").innerText();
  if (!/RESPOSTAS DO FORMUL/i.test(texto)) {
    throw new Error("Relatório do HubPrime sem respostas do formulário.");
  }

  const tipo = (/OS\s+OS-\S+\s+·\s+(.+)/.exec(texto) || [])[1] || "";
  const concluidoEm =
    (/CONCLU[IÍ]DO EM\s*\n\s*(\d{2}\/\d{2}\/\d{4} \d{2}:\d{2})/i.exec(texto) || [])[1] || "";
  return { tipo: tipo.trim(), concluidoEm, perguntas: lerRespostas(texto) };
}

/** Perguntas numeradas da secao "RESPOSTAS DO FORMULÁRIO". */
function lerRespostas(texto) {
  const linhas = texto.split("\n").map((l) => l.trim()).filter(Boolean);
  const inicio = linhas.findIndex((l) => /^RESPOSTAS DO FORMUL/i.test(l));
  const fim = linhas.findIndex((l, i) => i > inicio && /^CONFIRMA/i.test(l));
  const perguntas = [];

  for (const linha of linhas.slice(inicio + 1, fim > inicio ? fim : undefined)) {
    const numerada = /^(\d+)\.\s+(.*)$/.exec(linha);
    if (numerada) {
      perguntas.push({ numero: +numerada[1], titulo: numerada[2], partes: [] });
      continue;
    }
    // Icone de foto, data da foto e avisos da pagina nao sao resposta.
    if (!perguntas.length) continue;
    if (/^⤢$|^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$|^Toque em uma foto|v[ií]deo\(s\) registrado/i.test(linha)) {
      continue;
    }
    perguntas[perguntas.length - 1].partes.push(linha);
  }

  return perguntas.map(({ numero, titulo, partes }) => ({
    numero,
    titulo,
    resposta: partes.join(" / "),
  }));
}

/** Relatorio no formato de formulario do Field (name, status, questions). */
function comoFormulario(relatorio) {
  return {
    name: `HubPrime - ${relatorio.tipo || "O.S."}`,
    status: "done",
    questions: relatorio.perguntas.map((p) => ({
      title: p.titulo,
      answer: p.resposta || null,
      type: "text",
      position: p.numero - 1,
      conditions: [],
    })),
  };
}

module.exports = {
  loginTracker,
  listarOsHubPrime,
  lerRelatorio,
  comoFormulario,
};
