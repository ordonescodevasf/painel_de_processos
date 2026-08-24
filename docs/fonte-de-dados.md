# De onde o painel lê os dados

Duas fontes, nesta ordem:

1. **Google Sheets** — a planilha viva, em `CONFIG.googleSheetId` (`js/app.js`).
   Editar a planilha reflete no painel no próximo carregamento, sem
   republicar arquivo nenhum.
2. **`data/painel-processos-dados.xlsx`** — a cópia local, usada se o Google
   não responder.

Se as duas falharem, o painel mostra a tela de erro. É deliberado: melhor
dizer "não carregou" do que exibir um retrato antigo dos dados como se fosse
o atual.

O rótulo de fonte no cabeçalho do painel diz qual das duas venceu — confira
ali para saber se a edição da planilha chegou.

## Onde se configura

Em **`index.html`**, no bloco `window.PAINEL_CONFIG` do topo. Ele é aplicado
sobre os padrões de `js/app.js`, então **o que estiver ali vence** — e uma
chave deixada vazia ali anula o valor de `app.js`. É o único lugar a editar
para trocar de planilha.

```js
window.PAINEL_CONFIG = {
  googleSheetId: '1DzZJpqni8if_nT08ZDcDzw0TZDGLQtwi6wqjPbWfXic',
  arquivoXlsx: 'data/painel-processos-dados.xlsx'
};
```

## Google Sheets: o que precisa estar certo

- **Compartilhamento**: "qualquer pessoa com o link pode ver". O painel lê via
  `gviz`, que é JSONP e dispensa CORS, mas não dispensa a planilha ser
  legível sem login.
- **Nomes das abas** iguais aos de `CONFIG.abas` em `js/app.js`. Uma aba
  renomeada vem vazia e o console diz qual — exceto `Macroprocessos` e
  `Processos`, sem as quais não há painel: aí a fonte inteira falha e a cópia
  local assume.
- **Fórmulas não sobrevivem à importação.** As colunas calculadas do `.xlsx`
  chegam ao Google Sheets como `#ERROR!`: `Trilha` nas cinco abas da
  hierarquia, `Percentual` em Processos, `Vinculo_Nivel` em Documentos,
  Riscos, Indicadores e Papeis, e `Nivel_PxI` e `Classificacao` em Riscos. O
  painel recalcula todas a partir de dado puro da mesma linha — código da
  hierarquia, contagem dos dez marcos, formato do `Vinculo_Codigo`,
  probabilidade × impacto —, então não é preciso colar como valores. Se
  preferir corrigir na origem, cole as colunas como valores; o painel aceita
  as duas formas. O console informa quantos registros foram recalculados.
- **Ao trocar a planilha de lugar**, troque o ID no `index.html` — não em
  `js/app.js`.
- **Cabeçalhos das colunas** iguais aos da planilha local. O painel casa
  coluna por nome, não por posição, então mover coluna de lugar é seguro;
  renomear não.

## Por que não deu pelo Sharepoint

Ficou registrado porque a pergunta vai voltar quando a planilha mudar para o
OneDrive da Codevasf.

- **Link de pasta não serve.** Não existe API anônima que liste os arquivos de
  uma pasta compartilhada do OneDrive for Business. Quem faz isso é o
  Microsoft Graph, com token OAuth.
- **Link de arquivo também não.** O SharePoint não devolve
  `Access-Control-Allow-Origin`, então `fetch()` de outra origem é bloqueado —
  mesmo com o arquivo público e o link correto. O arquivo abre numa aba do
  navegador, mas o JavaScript do painel não consegue lê-lo.

Duas situações em que o SharePoint volta a funcionar:

- **Painel publicado na mesma origem da planilha** (no próprio SharePoint da
  Codevasf): a leitura é same-origin e CORS não se aplica.
- **Sincronismo** por Power Automate ou tarefa agendada, copiando o arquivo
  mais recente da pasta do OneDrive para `data/` no servidor do painel. Quem
  atualiza continua só subindo o arquivo na pasta.

## Consequência de ter só duas fontes

O painel **não abre mais com dois cliques no `index.html`** (protocolo
`file://`): sem servidor, o navegador bloqueia a leitura do `.xlsx` local, e
não há mais fonte embutida para cobrir. Para uso local, sirva a pasta por
HTTP:

```
python -m http.server 8000
```

O arquivo `js/dados.js` continua no projeto como retrato dos dados e insumo
dos scripts, mas não é mais carregado pelo painel. Para reativá-lo como
terceira fonte, basta voltar a incluí-lo no `index.html` e reintroduzir o elo
na cadeia de `carregarDados()`.
