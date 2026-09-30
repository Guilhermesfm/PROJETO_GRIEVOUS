# Grievous

Automações do Field Control:

- **Fechamento da folha** — lê a planilha de O.S. exportada do Field, abre cada
  O.S. e extrai as respostas dos formulários.
- **Placas** — consulta a situação e a data de criação de cada placa.

## Instalar em uma máquina nova

Pré-requisito: [Node.js](https://nodejs.org) 18 ou superior.

```bash
git clone <url-do-repositorio>
cd GRIEVOUS
npm run setup
```

O `npm run setup` instala as dependências **e** baixa o Chromium usado pelo
Playwright. Sem esse segundo passo a automação não abre o navegador.

## Rodar

No Windows, o jeito mais simples e o atalho:

```powershell
.\criar-atalho.ps1
```

Isso cria um **Grievous** na Area de Trabalho, com o icone do General. Um
duplo clique sobe o servidor e abre o navegador ja na tela de login. Na
primeira vez ele instala as dependencias sozinho.

Pelo terminal:

```bash
npm start
```

Abre o menu no terminal:

```
1) Fechamento (folha de pagamento)
2) Automação de placas
3) Interface web
4) Credenciais do Field
```

Ou direto:

```bash
npm run web         # interface em http://localhost:3000
npm run pagamento   # fechamento pelo terminal
npm run placas      # consulta de placas pelo terminal
npm run login       # grava/troca as credenciais
```

## Credenciais

Na primeira execução o Grievous pede o e-mail e a senha do Field e grava em um
arquivo `.env` na raiz do projeto. Pela interface web isso acontece na tela de
login; pelo terminal, na opção 4.

O `.env` está no `.gitignore`, então **não** vai junto no clone — cada máquina
precisa fazer esse cadastro uma vez.

O servidor web escuta apenas em `127.0.0.1`, ou seja, só é acessível pelo
navegador da própria máquina. Isso é proposital: a tela de login trafega a
senha do Field sem HTTPS. Para acessar de outro computador da rede seria
preciso antes colocar uma senha de acesso no próprio Grievous.

## Fechamento da folha

1. Exporte a planilha de O.S. do Field (colunas `Identificador`, `Tipo`,
   `Status`, `Cliente`, `Técnico`, `Avaria`).
2. Envie o arquivo pelo botão **Enviar planilha** da interface (ou arraste-o
   sobre o painel). Ele é gravado em `folha-de-pagamento/` e conferido na
   hora — se a planilha não servir, o erro aparece antes de começar.
   Copiar o arquivo direto para a pasta também funciona.
3. Rode o fechamento. O Grievous pega a planilha mais recente da pasta.

O resultado sai em `folha-de-pagamento/pagamento_formularios.xlsx`, com duas
abas:

- **Pagamento** — uma linha por O.S., com as perguntas mapeadas do tipo.
- **Formularios** — todas as respostas lidas, para conferência manual.

### Quais formulários entram

| Tipo da O.S.   | Perguntas                     |
| -------------- | ----------------------------- |
| Instalação     | 11, 15, 34, 45                |
| Remoção        | 13, 31, 34, 37, 39, 40, 42    |
| Manutenção     | 13, 31, 34, 37, 39, 40, 42    |
| Remanejamento  | união de instalação + remoção |

O.S. com status **Impedido** e tipos fora do escopo (preparação de materiais,
revisão preventiva, laudo) não são abertas — entram na planilha só como
registro. O mapeamento fica em `MAPA_PERGUNTAS`, no topo de
`folha-de-pagamento/Pagamento.js`.

## Estrutura

```
Grievous.js              menu principal
server.js                servidor da interface web
variaveis.js             leitura e gravação das credenciais no .env
filtro-field.js          filtro "Identificador" da listagem (usado pelas duas)
public/index.html        interface web
assets/grievous.ico      icone do atalho do Windows
public/grievous.svg      mascara usada no front e no favicon
Grievous.bat             sobe o servidor e abre o navegador
criar-atalho.ps1         cria o atalho na Area de Trabalho
placas/                  automação de placas (código e planilha gerada)
folha-de-pagamento/      fechamento (código, planilha de entrada e saída)
```

## Depuração

Para acompanhar o navegador durante a execução:

```bash
HEADLESS=false npm run pagamento
```

O HTML de cada O.S. visitada fica em `folha-de-pagamento/debug-formularios/`,
útil quando alguma resposta vem vazia.
