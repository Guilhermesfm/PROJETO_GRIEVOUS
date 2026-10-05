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

No Windows basta abrir o **Grievous.exe** que ja vem no repositorio: ele sobe
o servidor e abre o navegador na tela de login.

Para deixar na Area de Trabalho, crie um atalho: botao direito no
`Grievous.exe` > Enviar para > Area de trabalho. Assim ele continua ao lado
do `iniciar.js` e funciona sem mais nada.

Se preferir mover o proprio arquivo para fora da pasta, rode antes:

```powershell
npm run exe
```

Isso regera o executavel e grava um `Grievous.txt` com o caminho do projeto;
leve os dois juntos. O `Grievous.txt` nao e versionado porque guarda um
caminho absoluto, que muda de maquina para maquina.

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

## Em mais de uma maquina

Cada maquina precisa do seu proprio clone, **fora** de pastas sincronizadas
(OneDrive, Dropbox, Google Drive):

```bash
git clone <url-do-repositorio>
cd GRIEVOUS
npm run setup
```

Para atualizar depois de uma mudanca: `git pull` em cada maquina, e feche a
janela do Grievous antes de abrir de novo.

Rodar o projeto de dentro do OneDrive em duas maquinas e o caminho mais curto
para executar codigo velho: a sincronizacao tem atraso e cria copias
duplicadas ("arquivo (1).xlsx"), inclusive das planilhas de entrada. O
Grievous avisa no inicio quando detecta isso.

O servidor escuta so em `127.0.0.1`, entao cada maquina tem o seu, sem
disputa de porta entre elas. Na mesma maquina, uma segunda copia sobe
sozinha na porta seguinte (3001, 3002...).

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
Grievous.js              menu do terminal
iniciar.js               sobe o servidor e abre o navegador
Grievous.exe             executavel do Windows (versionado)
hello-there.ps1          recompila o Grievous.exe
Grievous.bat             alternativa ao .exe, pelo terminal

field/                   conversa com o Field (usado pelas duas automacoes)
  credenciais.js           le e grava EMAIL/PASSWORD e as URLs no .env
  filtro.js                filtro "Identificador" da listagem
  popups.js                fecha os avisos que bloqueiam a tela
  gaveta.js                abre e fecha a O.S. pelo botao de edicao
  link.js                  pega o link da O.S. pelo Compartilhar

web/                     interface
  servidor.js              rotas e execucao das automacoes
  public/index.html        a tela
  public/grievous.svg      mascara, usada no front e no favicon

placas/                  consulta de placas (codigo e planilha gerada)
folha-de-pagamento/      fechamento (codigo, planilha de entrada e saida)
assets/                  icone do executavel
ferramentas/             utilitarios avulsos, fora do fluxo principal
```

## Depuração

Para acompanhar o navegador durante a execução:

```bash
HEADLESS=false npm run pagamento
```

O HTML de cada O.S. visitada fica em `folha-de-pagamento/debug-formularios/`,
útil quando alguma resposta vem vazia.
