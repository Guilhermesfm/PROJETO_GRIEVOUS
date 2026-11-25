// Formatacao das planilhas de saida: ordem das colunas, largura, filtro e formatos.

const XLSX = require("xlsx");

// Texto usado quando o Field nao devolveu nada (celula vazia seria ambigua).
const SEM_RETORNO = "sem retorno";

/** Largura da coluna pelo conteudo, com minimo e maximo razoaveis. */
function larguraDaColuna(titulo, valores) {
  const maior = valores.reduce(
    (maximo, valor) => Math.max(maximo, String(valor ?? "").length),
    String(titulo).length,
  );
  return { wch: Math.min(Math.max(maior + 2, 10), 52) };
}

/**
 * Monta a aba a partir das linhas: colunas de `ordem` primeiro, o resto depois.
 */
function montarAba(linhas, ordem = [], formatos = {}) {
  if (!linhas.length) return XLSX.utils.json_to_sheet([]);

  const presentes = new Set();
  for (const linha of linhas) {
    for (const chave of Object.keys(linha)) presentes.add(chave);
  }

  const cabecalho = [
    ...ordem.filter((c) => presentes.has(c)),
    ...[...presentes].filter((c) => !ordem.includes(c)),
  ];

  // Normaliza: toda linha com todas as colunas, para nenhuma celula escorregar.
  const normalizadas = linhas.map((linha) =>
    Object.fromEntries(cabecalho.map((c) => [c, linha[c] ?? ""])),
  );

  const aba = XLSX.utils.json_to_sheet(normalizadas, { header: cabecalho });

  aba["!cols"] = cabecalho.map((titulo) =>
    larguraDaColuna(
      titulo,
      normalizadas.map((linha) => linha[titulo]),
    ),
  );

  // Formato por coluna ({ "Data": "dd/mm/yyyy" }).
  // Datas vao como serial do Excel: um Date sairia deslocado pelo fuso.
  for (const [titulo, formato] of Object.entries(formatos)) {
    const coluna = cabecalho.indexOf(titulo);
    if (coluna === -1) continue;
    for (let r = 1; r <= normalizadas.length; r += 1) {
      const celula = aba[XLSX.utils.encode_cell({ r, c: coluna })];
      if (celula && celula.t === "n") celula.z = formato;
    }
  }

  // Filtro em cada coluna.
  aba["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: normalizadas.length, c: cabecalho.length - 1 },
    }),
  };

  return aba;
}

/** Grava o arquivo com uma ou mais abas ja formatadas. */
function gravarPlanilha(caminho, abas) {
  const workbook = XLSX.utils.book_new();
  for (const { nome, linhas, ordem, formatos } of abas) {
    XLSX.utils.book_append_sheet(workbook, montarAba(linhas, ordem, formatos), nome);
  }
  XLSX.writeFile(workbook, caminho);
  return caminho;
}

module.exports = { gravarPlanilha, SEM_RETORNO };
