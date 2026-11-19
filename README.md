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
   `Status`, `Cliente`).
2. Envie o arquivo pelo botão **Enviar planilha** da interface (ou arraste-o
   sobre o painel). Ele é gravado em `folha-de-pagamento/` e conferido na
   hora — se a planilha não servir, o erro aparece antes de começar.
   Copiar o arquivo direto para a pasta também funciona.
3. Rode o fechamento. O Grievous pega a planilha mais recente da pasta.

Folhas já fechadas vão para `folha-de-pagamento/fechamentos/`: deixadas na
pasta principal, podem ser lidas como entrada do próximo fechamento.

O resultado sai em `folha-de-pagamento/pagamento_formularios.xlsx`:

- **Pagamento** — a folha, no formato usado com as empresas: Cliente,
  Identificador, Tipo, Status, Data Execução, Avaria, Violação, Obs, Link.
- **Formularios** — todas as respostas lidas, para conferência manual.
- **Fora do escopo** — tipos sem formulário de pagamento (preparação de
  materiais, preparação V2, laudo de equipamentos, revisão preventiva,
  validação de frotas, check list, extração de vídeo), que não são abertos,
  e as revisões/devoluções/laudos sem nenhuma ocorrência.

Revisão Rouxinol, Entrega e Devolução e Laudo Técnico InoPrime V2 não são
serviço pago: a O.S. é aberta e só vai para a folha se tiver avaria nos
equipamentos, extravio ou violação (no laudo, foto em "Tire fotos das
avarias").

A coluna Tipo da folha é a da planilha enviada. Mas, ao abrir a O.S., vale o
serviço cadastrado no Field: uma O.S. que a planilha chama de "Manutenção" e
o Field de "Extração de Videos" vai para "Fora do escopo", com o tipo do
Field na Obs.

As respostas não são lidas da tela: o programa abre a O.S. e usa os dados que
o próprio Field carrega (`formAnswers`), com todos os formulários da O.S.

### O que entra na folha

As perguntas são localizadas pelo **texto**, não pelo número — o número muda
entre formulários (o extravio é a 37 na "REMOÇÃO - GERAL" e a 34 na
"REMANEJAMENTO - REMOÇÃO").

| Formulário                          | Marca SIM em                     | Obs                                         |
| ----------------------------------- | -------------------------------- | ------------------------------------------- |
| Remoção / Remanejamento - Remoção   | Avaria: dano nos equipamentos    | equipamentos removidos, dano, extravio      |
| Instalação / Remanejamento - Instal.| —                                | equipamentos instalados                     |
| Manutenção / Revisão Rouxinol       | Violação: violação na instalação | descrição da violação, extravio             |
| Remanejamento InoPrime V2           | Avaria: periférico avariado      | "Periférico InoPrime avariado"              |
| Entrega e Devolução                 | Avaria: material devolvido       | relato da avaria, o que foi extraviado      |
| Laudo Técnico InoPrime V2           | Avaria: foto de avaria anexada   | aparelho, o que foi constatado, nº de fotos |
| Impedimento de visita               | —                                | descrição do motivo + motivo sem assinatura |

Avaria no veículo, antes ou depois do serviço, não entra na folha.

Tipos com outro nome também são reconhecidos: "RETIRADA" conta como remoção
e "MANUTEÇÃO" (assim, sem o N, no Field da Terrabel) como manutenção.

Nomes dos equipamentos na Obs:

| Marcado no formulário                  | Na folha                                |
| -------------------------------------- | --------------------------------------- |
| GalileoSky / Galileo 7X / JR12 / J-R12 | telemetria                              |
| Rastreador básico                      | telemetria basica                       |
| GalileoSky (ou JR) + rastreador básico | telemetria basica e telemetria avançada |
| Câmeras / V5 / NT407 / Momenta         | cam                                     |
| Outros (sistema de fadiga, redundante) | o próprio nome                          |

Formulário enviado só com fotos, sem as respostas que a folha usa
(equipamentos, dano, extravio, violação), recebe "Formulário incompleto" na
Obs. Formulário aberto e não finalizado é ignorado, com aviso na Obs. As
perguntas ficam em `PERGUNTAS`, no topo de `folha-de-pagamento/Pagamento.js`.

### Conferência de formulários

A cada fechamento, logo depois do login e antes de processar as O.S., o
Grievous lê a lista de formulários do Field e compara com o registro
(`folha-de-pagamento/formularios-registrados.json`, hoje com os 32 do
Field). A tela mostra quantos estão completos (com as perguntas) e lista,
pelo nome, na faixa "Formulários a registrar" e na aba **Conferência de
formulários**:

- **Novo** — está no Field e não no registro.
- **Alterado** — mudou a quantidade de perguntas ou foi editado depois do
  registro.
- **Removido** — está no registro e não está mais ativo no Field.
- **Sem perguntas** — registrado só com nome e quantidade de perguntas,
  sem o `.fct`. Some quando o `.fct` for registrado.

O.S. com formulário fora do registro recebe "Formulário não registrado: NOME"
na Obs.

O registro só tira o aviso: para a folha ler as respostas de um formulário
novo, as perguntas dele precisam estar em `PERGUNTAS` (Pagamento.js).

Para registrar formulários novos ou alterados: exporte o `.fct` do Field,
coloque na pasta `Novo_form` e rode (formulário sem `.fct` entra só com nome
e quantidade de perguntas, a partir da última lista lida do Field)

```bash
npm run formularios
```

A última lista lida do Field fica em `folha-de-pagamento/formularios/`
(local, fora do git).

## Consulta de placas

Na aba **Placas**, informe as placas (uma por linha ou separadas por
vírgula) e clique em Consultar. A barra abaixo dos botões mostra o
andamento placa a placa; outra aba aberta acompanha a mesma consulta. O
resultado sai em
`placas/situacoes_placas.xlsx`: Placa, Situacao, Data de criacao, Dias desde
a criacao (contado no dia da consulta), Link e Observacao.

O.S. não concluída nem cancelada criada há mais de 30 dias aparece em
alerta na tela: faixa "Distúrbio na Força" no topo e selo com os dias na
linha.

## Estrutura

```
Grievous.js              menu do terminal
iniciar.js               sobe o servidor e abre o navegador
Grievous.exe             executavel do Windows (versionado)
hello-there.ps1          recompila o Grievous.exe
Grievous.bat             alternativa ao .exe, pelo terminal
planilha.js              grava as planilhas (usado pelas duas automacoes)

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
  fechamentos/             folhas ja fechadas; o programa nao le esta pasta
assets/                  icone do executavel
```

## Ver o navegador

Para acompanhar o navegador durante a execução:

```bash
HEADLESS=false npm run pagamento
```

Quando uma O.S. não abre ou o Field não devolve os formulários, o motivo
aparece na Obs da linha, com "ERRO:".
