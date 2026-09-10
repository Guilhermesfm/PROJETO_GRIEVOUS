const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'ia.txt');
const html = fs.readFileSync(filePath, 'utf8');

function stripHtml(value) {
  return value
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function rowsFromHtml(rawHtml) {
  const rowRegex = /<tr\b[^>]*role="row"[^>]*>(.*?)<\/tr>/gs;
  const rows = [];
  let match;
  while ((match = rowRegex.exec(rawHtml)) !== null) {
    const rowHtml = match[1];
    const cellRegex = /<(?:td|th)\b[^>]*>(.*?)<\/(?:td|th)>/gs;
    const cells = [];
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
      cells.push(stripHtml(cellMatch[1]));
    }

    if (cells.length) {
      rows.push(cells);
    }
  }
  return rows;
}

function normalizeCell(value) {
  return value
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function toCsv(rows) {
  const header = ['identificador', 'cliente', 'situacao', 'tipo', 'colaborador'];
  const finalRows = rows
    .map((cells) => cells.map(normalizeCell))
    .filter((cells) => cells.length >= 5)
    .map((cells) => {
      const [identificador, cliente, situacao, tipo, colaborador] = cells;
      return [identificador, cliente, situacao, tipo, colaborador].join(',');
    });

  return [header.join(','), ...finalRows].join('\n');
}

const rows = rowsFromHtml(html);
const csv = toCsv(rows);
console.log(csv);
