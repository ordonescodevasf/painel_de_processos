# Painel de Gestão de Processos — Codevasf (AE/GPE/UNP)

Site estático (HTML + CSS + JavaScript puros, padrão **gov.br DS v4**) que publica a
cadeia de valor, o repositório de processos e a rastreabilidade do trabalho de
mapeamento da Codevasf. Substitui o painel construído no Lovable (React + Supabase):
não há build, backend nem consumo de créditos — basta hospedar no **GitHub Pages** e
alimentar por **planilha**.

> **ATENÇÃO — DADOS FICTÍCIOS:** todos os nomes, números, normativos internos,
> resultados e registros desta versão são fictícios, criados apenas para demonstrar
> o painel. Substitua pelo conteúdo real antes de divulgar.

O cabeçalho, o rodapé, o menu lateral, a seção de equipe e o VLibras são herdados do
**Painel de Transformação Digital** da Codevasf (mesmo CSS, extraído do arquivo
original), garantindo identidade visual idêntica entre os dois painéis.

## Estrutura do repositório

```
painel-processos/
├── index.html                     ← página única do painel (SPA com rotas #/)
├── .nojekyll                      ← evita processamento Jekyll no GitHub Pages
├── css/
│   ├── govbr-ds.css               ← DS gov.br v4 (extraído do painel do PTD)
│   └── painel.css                 ← estilos das seções deste painel
├── js/
│   ├── app.js                     ← dados, rotas e telas
│   ├── govbr-ui.js                ← comportamentos do cabeçalho/menu/rodapé
│   └── dados.js                   ← dados embutidos (GERADO — fallback offline)
├── data/
│   └── painel-processos-dados.xlsx← A PLANILHA (fonte dos dados)
├── img/diagramas/*.svg            ← diagramas BPMN (troque pelos do Bizagi)
└── scripts/
    ├── gerar_planilha.py          ← gera a planilha de exemplo (dados fictícios)
    ├── planilha_para_js.py        ← planilha → js/dados.js (fallback)
    ├── gerar_diagramas.py         ← gera os SVGs ilustrativos
    └── montar_index.py            ← remonta o index a partir do painel do PTD
```

## Como o painel é alimentado

O `js/app.js` tenta as fontes **nesta ordem** e usa a primeira que funcionar
(o chip "Tempo real" no menu *Atalhos gov.br* mostra qual fonte está ativa):

1. **Google Sheets (recomendado)** — mesmo esquema do painel do PTD, com
   atualização sem commit:
   1. Acesse [sheets.google.com](https://sheets.google.com) → **Arquivo →
      Importar** → envie `data/painel-processos-dados.xlsx` (mantenha os nomes
      das abas).
   2. **Compartilhar → Qualquer pessoa com o link → Leitor**.
   3. Copie o ID da planilha (o trecho entre `/d/` e `/edit` na URL).
   4. Cole no topo do `index.html`, em `window.PAINEL_CONFIG`:
      ```js
      googleSheetId: 'COLE_O_ID_AQUI',
      ```
   5. Pronto: edite a planilha no Google e o painel reflete no próximo
      carregamento (F5). O item "Abrir planilha de origem" do cabeçalho passa a
      apontar para ela.

2. **Planilha no repositório** — se `googleSheetId` estiver vazio, o site baixa
   `data/painel-processos-dados.xlsx` e a lê no navegador (SheetJS). Fluxo:
   editar o arquivo → *commit* → GitHub Pages atualiza.

3. **Dados embutidos (`js/dados.js`)** — usado quando as opções acima falham
   (ex.: abrir o `index.html` com dois cliques, sem internet/servidor). Depois de
   alterar a planilha, regenere com:
   ```bash
   python scripts/planilha_para_js.py     # requer: pip install openpyxl
   ```

A planilha tem **17 abas**. A hierarquia segue o CBOK 4.0 em cinco níveis: **Macroprocesso → Processo de negócio (aba Processos) → Processo de trabalho (aba Subprocessos) → Atividade → Tarefa (aba Tarefas)** — os nomes técnicos das abas foram mantidos por compatibilidade, e o site exibe a nomenclatura correta.  além da hierarquia (Macroprocessos → Atividades) e dos
vínculos (Documentos, Riscos, Indicadores, Diario_Mapeamento), as abas **Jornada**
e **Repositorio** alimentam a aba *Repositório* do site (jornada de mapeamento,
metodologia/guia RES 031/2025, instrumentos, modelos e ferramentas), **NUGEP**
cadastra os integrantes do núcleo, **Glossario** e **FAQ** alimentam as abas
homônimas, e **Parametros** guarda contato institucional e os links da metodologia
e do guia. Nada de conteúdo fica fixo no site: para editar, adicionar ou remover
qualquer informação, mexa na planilha.

### Regras de preenchimento (resumo — detalhes na aba LEIA-ME)

Os vínculos usam os **códigos**: `Processos.Macroprocesso` → `MP-xx`,
`Subprocessos.Processo` → `P-xx.yy`, `Atividades.Subprocesso` → `SP-xx.yy.zz`,
`Tarefas.Atividade` → `A-xx.yy.zz.ww`;
Documentos, Riscos e Indicadores usam `Vinculo_Nivel` + `Vinculo_Codigo`; o
Diário usa o código do Processo. Listas dentro de uma célula são separadas por
`;`. Evidências do diário seguem `Nome|URL` (várias separadas por `;`). Datas em
`dd/mm/aaaa`; percentuais de 0% a 100%. Nas colunas cinza (calculadas), copie a
fórmula da linha de cima ao inserir linhas.

## Publicar no GitHub Pages

1. Crie um repositório e envie **todo o conteúdo desta pasta para a raiz**
   (incluindo o `.nojekyll`).
2. **Settings → Pages → Build and deployment → Deploy from a branch** →
   branch `main`, pasta `/ (root)` → **Save**.
3. O painel fica em `https://SEU-USUARIO.github.io/NOME-DO-REPO/`.

## Testar no computador

```bash
cd painel-processos
python -m http.server 8080     # abra http://localhost:8080
```
Abrir o `index.html` direto (file://) também funciona, pelos dados embutidos.

## Diagramas do Bizagi (recomendado: por link)

Exporte o diagrama do Bizagi como imagem e **publique-a on-line** (Google Drive
público, intranet acessível, repositório de imagens etc.); depois cole a **URL na
coluna `Imagem_Bizagi`** da planilha. O painel exibe a imagem e, ao clicar nela
(ou no botão "Abrir diagrama no link publicado"), a pessoa é levada ao link
original em nova aba. Links de compartilhamento do Google Drive são convertidos
automaticamente para exibição; se a pré-visualização não carregar (link privado,
bloqueio de intranet), o painel mostra o botão de acesso mesmo assim.
Alternativa: salvar o arquivo em `img/diagramas/` e usar o caminho relativo.

## Ajustes que você provavelmente vai querer fazer

Quase tudo se edita **na planilha** (inclusive equipe do NUGEP, glossário, FAQ e
repositório). No `index.html` ficam apenas: a linha institucional do rodapé, o
botão flutuante **"Sugerir melhoria de processo"** (aponta para
`ae.gpe.unp@codevasf.gov.br`) e os textos dos modais "Como usar" e
"Acessibilidade". A seção de Ouvidoria/Fala.BR do painel anterior foi removida.
A logo do Governo Federal no rodapé aponta para uma URL externa (repositorio.ifms.edu.br); para não depender de site de terceiro, baixe o PNG para `img/` e ajuste o `src` no `index.html`.

## Metodologia

Estruturado no **BPM CBOK 4.0** (hierarquia macroprocesso → processo →
subprocesso → atividade; tipos finalístico/suporte/gerencial; ciclo de vida BPM
em 5 fases; SIPOC; dono do processo) e no **PMBOK** (cada mapeamento tratado
como projeto: termo de abertura, marcos M1–M9, entregáveis, riscos e lições
aprendidas). A aba **Metodologia** do painel documenta tudo isso para o público.
