// Formatacao das planilhas de saida.
//
// O json_to_sheet entrega a aba crua: colunas todas com a mesma largura, sem
// filtro e na ordem em que as chaves aparecem. Com centenas de linhas e
// colunas de texto longo, isso fica ilegivel.

const XLSX = require("xlsx");

// Texto unico para "o Field nao devolveu nada". Celula vazia e ambigua: nao
// da para saber se a pergunta nao foi respondida ou se a leitura falhou.
const SEM_RETORNO = "sem retorno";
const NAO_EXISTE = "nao existe neste formulario";

/** Largura da coluna pelo conteudo, com minimo e maximo razoaveis. */
function larguraDaColuna(titulo, valores) {
  const maior = valores.reduce(
    (maximo, valor) => Math.max(maximo, String(valor ?? "").length),
    String(titulo).length,
  );
  return { wch: Math.min(Math.max(maior + 2, 10), 52) };
}

/**
 * Monta a aba a partir das linhas, com as colunas na ordem pedida.
 *
 * Colunas listadas em `ordem` vem primeiro, na ordem dada; o resto vem
 * depois, na ordem em que aparecer. Sem isso a ordem segue a primeira linha
 * e muda conforme a O.S. que for lida primeiro.
 */
function montarAba(linhas, ordem = []) {
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

  // Filtro em cada coluna. Congelar o cabecalho nao entra aqui: o SheetJS
  // desta versao nao grava painel congelado no arquivo, entao o codigo so
  // daria a impressao de ter configurado.
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
  for (const { nome, linhas, ordem } of abas) {
    XLSX.utils.book_append_sheet(workbook, montarAba(linhas, ordem), nome);
  }
  XLSX.writeFile(workbook, caminho);
  return caminho;
}

module.exports = {
  montarAba,
  gravarPlanilha,
  larguraDaColuna,
  SEM_RETORNO,
  NAO_EXISTE,
};
