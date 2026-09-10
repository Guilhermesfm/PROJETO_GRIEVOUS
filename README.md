# PROJETO_GRIEVOUS

Automação de testes em Playwright para acessar o sistema FieldControl, filtrar registros por status e exportar os dados em CSV.

## Visão geral

Este projeto automatiza o login em um sistema web, navega até a listagem de atividades, aplica filtros e salva as linhas relevantes em um arquivo CSV localizado em:

- `thisisnotthedroidyouarelookingfor/dados.csv`

A automação foi desenvolvida em JavaScript usando:

- Playwright
- Node.js
- dotenv

## Estrutura do projeto

- `grievous/example.spec.js` — fluxo principal do teste automatizado
- `variaveis.js` — leitura das credenciais via variáveis de ambiente
- `extract_table_csv.js` — utilitário para leitura e transformação de HTML em CSV
- `thisisnotthedroidyouarelookingfor/` — pasta de saída do CSV gerado
- `.github/workflows/playwright.yml` — pipeline de CI para execução dos testes no GitHub Actions

## Pré-requisitos

- Node.js 18+
- npm
- Conta com acesso ao sistema de destino

## Instalação

1. Clone o repositório
2. Instale as dependências:

```bash
npm install
```

3. Crie um arquivo `.env` na raiz do projeto com as seguintes variáveis:

```env
EMAIL=seu_email@exemplo.com
PASSWORD=sua_senha
```

> O arquivo `.env` está ignorado pelo Git por segurança.

## Execução

Para rodar o teste automatizado:

```bash
npx playwright test
```

Ao final, o resultado será salvo em:

```bash
thisisnotthedroidyouarelookingfor/dados.csv
```

## Observações importantes

- As credenciais não devem ser enviadas para o repositório.
- O arquivo `.env` deve ficar localmente em sua máquina.
- O projeto usa o GitHub Actions para rodar os testes automaticamente em cada push e pull request.

## Licença

Este projeto é destinado a uso interno e automação operacional do ambiente informado.
