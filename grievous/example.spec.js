import { test, expect } from '@playwright/test';
import fs from 'fs';
import path from 'path';
// Importa as variáveis seguras de login do outro arquivo
import { email_Login, password_Login } from '../variaveis.js';

test('get started link', async ({ page }) => {
  // Abre o site na página de login
  await page.goto('https://app.fieldcontrol.com.br/autenticador-v2/#/login?to=:hash:%2Fatividades');

  // Digita o e-mail
  await page.fill('input[type="email"]', email_Login);

  // Clica no botão de avançar (a classe .min-w-full é do botão)
  await page.click('.min-w-full');

  // Digita a senha
  await page.fill('input[type="password"]', password_Login);

  // Clica no botão de "Next" (Entrar)
  await page.click('span:has-text("Next")');

  // Espera a tabela do sistema carregar e aparecer na tela
  await page.waitForSelector('table[role="table"]');

  // Espera a rede acalmar (garante que os dados pararam de carregar)
  await page.waitForLoadState('networkidle');

  // Fecha o aviso "Mais tarde" que pode aparecer na frente
  await page.click('button:has-text("Mais tarde")')

  // Abre o menu de filtros da tabela
  await page.click('[data-cy="filter-button"]');

  // Clica para filtrar os itens que estão com o status "Completed"
  await page.click(
    'palantir-content:has-text("By status") button:has-text("Completed")'
  )

  // Executa Javascript dentro do navegador para extrair as linhas da tabela (tbody tr)
  const rows = await page.locator('table[role="table"] tbody tr').evaluateAll((elements) => {
    return Array.from(elements)
      .map((row) => {
        // Pega as colunas <td>, limpa espaços vazios e transforma em texto
        const cells = Array.from(row.querySelectorAll('td')).map((cell) =>
          cell.innerText.replace(/\s+/g, ' ').trim()
        );

        // Retorna um objeto só com os índices HTML exatos das 6 colunas desejadas
        return {
          identificador: cells[1] || '',
          cliente: cells[3] || '',
          localizacao: cells[4] || '',
          situacao: cells[5] || '',
          tipo: cells[9] || '',
          colaborador: cells[10] || ''
        };
      })
      // Ignora as linhas vazias ou de cabeçalho filtrando só quem tem identificador
      .filter((row) => row.identificador);
  });

  // Cria o formato de texto CSV juntando tudo com vírgulas
  const csv = [
    // Linha de cabeçalho
    ['Identificador', 'Cliente', 'Localização', 'Situação', 'Tipo', 'Colaborador'].join(','),
    // Dados de cada linha coletada
    ...rows.map((row) =>
      [
        row.identificador,
        row.cliente,
        row.localizacao,
        row.situacao,
        row.tipo,
        row.colaborador
      ].join(',')
    )
  ].join('\n'); // Quebra de linha no final

  // Define a pasta onde o CSV será guardado
  const pasta = path.join(process.cwd(), 'thisisnotthedroidyouarelookingfor');

  // Cria a pasta automaticamente se ela não existir
  if (!fs.existsSync(pasta)) {
    fs.mkdirSync(pasta, { recursive: true });
  }

  // Define o caminho final do arquivo de saída
  const filePath = path.join(pasta, 'dados.csv');

  // Escreve os dados (salva) dentro do arquivo CSV no seu PC
  fs.writeFileSync(filePath, csv, 'utf8');

  // Avisa no terminal que deu tudo certo
  console.log('CSV salvo em:', filePath);
  console.log(csv);

  // Testa se conseguiu pegar no mínimo 1 linha de dado
  expect(rows.length).toBeGreaterThan(0);
});