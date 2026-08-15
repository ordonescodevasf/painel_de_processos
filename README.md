# Repositório de Processos — Codevasf

Painel de consulta pública do portfólio de processos da Codevasf: cadeia de
valor, fichas de macroprocesso a tarefa, marcos do mapeamento (M1–M10),
documentos publicados, riscos, indicadores, repositório de materiais, glossário
e perguntas frequentes.

Site estático — HTML, CSS e JavaScript clássico, sem etapa de build. Abre direto
no navegador e roda no GitHub Pages sem nenhuma configuração além de ligar a
publicação.

Segue o **Padrão Digital de Governo (GOV.BR DS), versão 4**. O registro de cada
decisão de design está em [DESIGN-SYSTEM.md](DESIGN-SYSTEM.md).

## Publicar no GitHub Pages

1. Suba o conteúdo desta pasta na **raiz** do repositório (o `index.html` precisa
   ficar na raiz, não dentro de uma subpasta).
2. No repositório: **Settings → Pages**.
3. Em *Source*, escolha **Deploy from a branch**; em *Branch*, `main` e pasta
   `/ (root)`. Salve.
4. Em um ou dois minutos o painel fica disponível em
   `https://<usuario>.github.io/<repositorio>/`.

O arquivo `.nojekyll` na raiz desliga o processamento Jekyll do GitHub Pages —
sem ele, arquivos e pastas que comecem com `_` seriam ignorados na publicação.

Todos os caminhos são relativos, então o painel funciona igual na raiz do
domínio, em subpasta de repositório ou aberto do disco (`file://`).

## Estrutura

```
index.html                     única página; as telas são rotas por #hash
css/govbr-ds.css               tokens e componentes do GOV.BR DS v4
css/painel.css                 o que é específico deste painel
js/app.js                      dados, rotas e telas
js/govbr-ui.js                 comportamento dos componentes do DS
js/govbr-datetimepicker.js     seletor de data e período
js/dados.js                    dados embutidos (gerado da planilha)
data/painel-processos-dados.xlsx   a planilha que alimenta o painel
img/                           logo, ilustrações oficiais e arte do loading
scripts/                       geradores em Python (planilha e diagramas)
```

## Como o painel obtém os dados

Três fontes, na ordem em que são tentadas:

1. **Google Sheets** — se `googleSheetId` estiver preenchido no `index.html`.
   Importe a planilha no Google Sheets, compartilhe como "qualquer pessoa com o
   link pode ver" e cole o ID (o trecho entre `/d/` e `/edit`). É o caminho
   recomendado para atualizar sem publicar de novo.
2. **`data/painel-processos-dados.xlsx`** — lido no navegador pelo SheetJS.
   Funciona no GitHub Pages; não funciona em `file://`.
3. **`js/dados.js`** — cópia embutida da planilha, gerada por
   `scripts/planilha_para_js.py`. É o que faz o painel abrir mesmo do disco.

## Alimentar com dados reais

A planilha em `data/` já traz a estrutura completa e dados de demonstração para
substituir. A aba **LEIA-ME** explica as regras de preenchimento, os códigos que
ligam as abas e as colunas em cinza, que se calculam sozinhas.

Depois de preencher, regenere a cópia embutida:

```
python scripts/planilha_para_js.py
```

Scripts de manutenção da planilha:

| Situação | Comando |
| --- | --- |
| Só aplicar a formatação do painel na planilha atual | `python scripts/formatar_planilha.py` |
| Acrescentar colunas e linhas novas que o painel passou a ler | `python scripts/atualizar_planilha.py` |
| Recriar a planilha de exemplo do zero (**apaga os dados**) | `python scripts/gerar_planilha.py` |

Requisito dos scripts: Python 3 com `openpyxl` (`pip install openpyxl`).

## Textos editáveis sem mexer no código

A aba **Parametros** da planilha controla os textos e contatos que aparecem no
painel — título e subtítulo da tela inicial, contato institucional, e-mail e
telefone da unidade, links do rodapé. Alterar ali muda o painel sem tocar em
nenhum arquivo.

## Conferência opcional do código

Nenhuma é necessária para o painel funcionar:

```
npx @biomejs/biome check js      # formato e lint do JavaScript
npx cspell --config cspell.json  # ortografia pt-BR
```

## Licenças

O painel reimplementa componentes, tokens e utilitários do Padrão Digital de
Governo e distribui ilustrações do pacote oficial (CC0 1.0 Universal e MIT).
Os créditos completos estão em [LICENSES-TERCEIROS.md](LICENSES-TERCEIROS.md).
