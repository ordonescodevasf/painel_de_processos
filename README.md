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
└── scripts/
    ├── gerar_planilha.py          ← gera a planilha de exemplo (dados fictícios)
    ├── planilha_para_js.py        ← planilha → js/dados.js (fallback)
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

## Diagramas do Bizagi (somente por link)

O repositório **não tem pasta de imagens**: exporte o diagrama do Bizagi e
**publique a imagem on-line** (Google Drive público, intranet acessível,
repositório institucional de imagens etc.), depois cole a **URL na coluna
`Imagem_Bizagi`** — presente em *todos* os níveis: Macroprocessos, Processos,
Subprocessos, Atividades e Tarefas (preencha onde for pertinente). O painel
exibe a imagem e, ao clicar nela ou no botão "Abrir diagrama no link publicado",
leva ao endereço original em nova aba. Links de compartilhamento do Google Drive
são convertidos automaticamente para exibição; se a prévia não carregar (link
privado, bloqueio de rede), o painel mostra o botão de acesso mesmo assim. As
URLs atuais (`placehold.co`) são apenas exemplos.

## Navegação: 6 abas fixas + "Mais"

A faixa de seções mostra sempre **Início, Catálogo, Dashboard, Documentos,
Riscos e Indicadores** em uma única linha, sem quebrar e sem precisar arrastar
a barra para o lado. As demais seções (Diário, Repositório, NUGEP, Glossário,
FAQ) ficam num botão **"Mais"**, que abre uma lista (mesmo componente List/Item
do menu de atalhos do cabeçalho) — quando uma dessas seções está aberta, o
próprio botão "Mais" mostra o nome dela. Em telas menores que ~992px, as abas
fixas viram só ícone (com texto acessível para leitor de tela), garantindo que
sempre caibam numa linha só, em qualquer tamanho de tela. O painel "Mais" é
posicionado por JavaScript (`position:fixed`, calculado a partir do botão),
o que o deixa imune a cortes de overflow ou disputas de empilhamento com o
resto da página — sempre visível e clicável por cima de tudo.

## Responsividade

Os cartões, grades e a cadeia de valor usam `minmax(min(Npx, 100%), 1fr)` —
nunca vazam para o lado, mesmo em telas bem estreitas (a partir de 320px). Os
valores institucionais da cadeia de valor ficam numa única linha quando cabem
e quebram para duas ou mais linhas automaticamente em telas menores, sem
precisar de rolagem. Tabelas longas (Riscos, Processos etc.) rolam apenas
dentro de si mesmas quando não cabem, sem mover a página inteira.

## Botões flutuantes, LGPD e tutorial

Adaptados diretamente do painel do PTD (mesmos componentes gov.br DS), com os
textos ajustados para este painel:

- **Voltar ao topo** (canto inferior direito): discreto, só aparece depois de
  rolar a página.
- **Reportar um erro** (canto inferior esquerdo): abre um modal que tenta
  capturar a tela (Screen Capture API — o navegador pede permissão; não existe
  captura silenciosa) e oferece o print para download. Como um link `mailto:`
  **não aceita anexos** (limitação de todos os navegadores/clientes de
  e-mail, sem contorno possível), o e-mail abre pronto com título, descrição
  e instrução para anexar manualmente o print baixado antes de enviar.
- **Aviso de recursos de terceiros (LGPD)**: barra fixa no primeiro acesso,
  listando honestamente o que o painel carrega (fontes, ícones, SheetJS,
  VLibras e, se configurado, Google Sheets) — sem alegar cookies que não
  existem no código. Só fecha com "Entendi" (fica salvo no navegador); o botão
  de cookies no menu de atalhos reabre a versão detalhada a qualquer momento.
- **Tutorial (onboarding)**: carrossel de 5 passos, abre sozinho no primeiro
  acesso (meio segundo depois, para não competir com o carregamento dos
  dados) e pode ser reaberto por "Como usar este painel" (rodapé ou menu de
  atalhos) ou pelo switch "Abrir tutorial automaticamente" nesse mesmo menu.
- **Equipe de Desenvolvimento**: link discreto no rodapé ("Sobre o Painel"),
  distinto do NUGEP — credita quem desenvolve e mantém tecnicamente o painel
  (por padrão, Antonio Ordones Neto; ajuste o nome/e-mail em `index.html`,
  busque por `modalEquipeDev`).

## LGPD, Ouvidoria e crédito da Equipe de Desenvolvimento

- **Equipe de Desenvolvimento** (rodapé → "Sobre o Painel"): agora com foto,
  e-mail e telefone do Antonio Ordones, no mesmo formato do card "Quem é Quem"
  do painel do PTD.
- **Rodapé → Atendimento**: link direto para a Ouvidoria da Codevasf
  (Plataforma Fala.BR).
- **Rodapé (linha institucional)**: endereço completo e CNPJ da Codevasf,
  conferidos na fonte oficial (codevasf.gov.br).
- **Política de Privacidade e Termos de Uso** (rodapé, dois links): página
  completa e verificada — controlador (CNPJ/endereço), Encarregado/DPO
  oficial da Codevasf com canal Fala.BR, quais dados o painel efetivamente
  trata (praticamente nenhum: é estático, sem servidor próprio), base legal,
  recursos de terceiros/transferência internacional, retenção, direitos do
  titular (art. 18 da LGPD) e termos de uso. Não substitui aconselhamento
  jurídico — para a política corporativa completa, a página linka a fonte
  oficial da Codevasf.

## Cores nos campos da ficha

Com muitos campos preenchidos, o texto das fichas (Macroprocesso, Processo,
Subprocesso, Atividade, Tarefa) ficava difícil de escanear rápido. Os campos
agora têm uma barra lateral + rótulo tingido por categoria — só com cores já
usadas no painel: **azul** = descritivo (Descrição, Objetivo), **navy** =
responsabilidade (Dono, Unidade Responsável, Interlocutor), **verde** = valor
entregue (Entregas, Beneficiários, Entradas/Saídas), **petróleo** =
técnico/normativo (Sistemas, Normativos, Base Normativa). Sem legenda — o
nome do campo já diz o que é; a cor só ajuda a agrupar visualmente. Some
sozinho no modo alto contraste.

## Metodologia: marcos do mapeamento (M1–M10)

Os marcos passaram de 9 para **10 etapas**, começando pela reunião de
contextualização (antes era o 3º passo):

`M1` Conhecer o processo · `M2` Processos modelados · `M3` Subprocessos
modelados · `M4` AS-IS modelado · `M5` AS-IS validado · `M6` Procedimento
aprovado · `M7` Processo publicado · `M8` TO-BE elaborado · `M9` TO-BE
aprovado · `M10` Processo transformado.

M2 e M3 representam as diversas oficinas de modelagem entre a contextualização
e o AS-IS consolidado — desenhando macroprocesso, processo e subprocessos
(inclusive descobrindo subprocessos ainda não mapeados). M6/M7 tratam do PRO
(Procedimento, CBOK 4.0): primeiro validado tecnicamente
pela equipe, depois aprovado pela autoridade competente. Cada marco tem um
tooltip explicando o que significa — passe o cursor sobre ele na ficha do
processo ou na aba Repositório. Na planilha, as colunas viraram
`M1_Reuniao_Contextualizacao` … `M10_Publicado_Repositorio` (10 colunas
Sim/Não na aba Processos).

## Subprocesso dentro de subprocesso

O CBOK 4.0 não fixa a profundidade da decomposição de processos ("Levels Vary
in Number and Name") — um subprocesso pode conter outro subprocesso, quantos
níveis o processo exigir, até chegar à atividade. Na aba **Subprocessos** da
planilha, a coluna `Vinculo_Pai` aceita tanto o código de um **Processo**
(`P-...`) quanto de **outro Subprocesso** (`SP-...`); o painel detecta
automaticamente pelo prefixo do código e monta o breadcrumb e a ficha
corretamente em qualquer profundidade. Veja o exemplo já incluído:
`SP-06.01.03.01` (Tratamento de Preços Inexequíveis ou Excessivos) é filho de
`SP-06.01.03` (Pesquisa de Preços), que por sua vez é filho do processo
`P-06.01`.

## Ajustes desta rodada (resumo)

- **Marcos**: M2 "Processos modelados"; M6 "Procedimento elaborado" (tooltip
  menciona o encaminhamento para validação do dono do processo); M7
  "Procedimento aprovado" pela **Diretoria Executiva (DEX)**; M9 "TO-BE
  aprovado" (validado pelo dono, aprovado no DEX).
- **Diário de Mapeamento removido** do painel inteiro (menu, ficha do
  processo, busca, planilha) — evidências agora ficam no processo
  e-Codevasf/SEI.
- **Repositório**: 18 itens removidos (formulários, templates e ferramentas
  hoje cobertos pelo SEI/e-Codevasf ou fora do escopo atual); seções "Ciclo
  de vida BPM" e "Como este painel é alimentado" removidas.
- **Alertas (sino) removidos** do cabeçalho — não é mais necessário.
- **Catálogo → Processos**: aba renomeada.
- **Início**: tooltips em todos os KPIs; card de documentos/diário
  removido; novo card de contagem de **Processos** como 2º card.
- **Ficha do Macroprocesso**: campo "Normativos aplicáveis" removido; botão
  explícito de abrir diagrama removido (a imagem já é clicável).
- **Ficha do Processo**: "Gerência responsável" (era Área); "Responsável no
  NUGEP" no lugar de "Interlocutor do mapeamento" (mesmo dado); Prioridade e
  Complexidade em campos separados; Descrição e Objetivo em cards
  separados; chip de fase do ciclo BPM removido; o card "Cronograma do
  projeto" virou **"Subprocessos vinculados"** no aside — navegação por
  drill-down igual à do macroprocesso.
- **Ficha da Atividade**: sem diagrama (não existe nesse nível); sem
  indicadores de desempenho.
- **Ficha da Tarefa**: sem diagrama, sem riscos/indicadores/documentos.
- **NUGEP**: sem cargo — só sigla do setor, e-mail e telefone — mais os
  **processos sob responsabilidade** de cada pessoa (derivado
  automaticamente da coluna `Interlocutor` da aba Processos).
- **Dashboard**: "Processos publicados" (era Cobertura da carteira); cards
  de Indicadores na meta e Tarefas mapeadas removidos; 3 novos cards
  (Subprocessos, Atividades, Tarefas); gráficos maiores e interativos (a
  barra "Avanço por macroprocesso" agora navega para a ficha ao clicar);
  gráfico de marcos corrigido para M1→M10; gráfico "Documentos por
  situação" removido; texto de introdução removido.
- **Tradução**: termos em inglês soltos traduzidos ("process owner", os
  parênteses do SIPOC etc.) — mantidas as siglas técnicas usadas assim
  mesmo na redação oficial brasileira de BPM (SIPOC, BPMN, CBOK, PMBOK,
  AS-IS, TO-BE, KPI, PRO).

## Diagrama via iframe (Bizagi Web Publish) — modo único

O painel incorpora o diagrama **ao vivo por iframe** — o mesmo padrão que a
própria Codevasf já usa internamente (Base de Conhecimento/Wiki.js, AA/GTI)
para o "Gerenciamento de Incidentes de TI", carregando a publicação web do
Bizagi Modeler direto de um servidor. Disponível em **Macroprocessos,
Processos e Subprocessos** (os 3 níveis com diagrama). Não há mais modo de
imagem estática — foi removido a pedido, junto com o antigo botão "Abrir
diagrama no link publicado".

Como usar: no Bizagi Modeler, **Publish → Web**, suba o pacote gerado num
servidor e cole a URL da publicação em `Imagem_Bizagi`, na planilha. Deixe
em branco enquanto não houver publicação — o painel mostra um aviso de
"ainda não publicado" em vez de quebrar. `P-06.01` já traz um exemplo real
(`https://fluxosti.codevasf.gov.br/incidentes/`) só para visualização — troque
pela publicação real do processo quando ela existir.

**Antes de apontar para um servidor real, confirme com quem o hospeda:**
como o painel roda no GitHub Pages (origem diferente de qualquer servidor
interno Codevasf), o cabeçalho `Content-Security-Policy: frame-ancestors`
do servidor da publicação precisa incluir o domínio do painel — sem isso, o
quadro aparece em branco (o painel sempre mostra também um botão "Abrir em
nova aba" como alternativa, que funciona independentemente disso). Se o
servidor for de rede interna, quem acessa o painel fora da VPN não verá o
conteúdo do iframe.

## Ajustes desta rodada (resumo)

- **Diagrama: só iframe agora** (Bizagi Web Publish) — removida a opção de
  imagem estática e o antigo botão "Abrir diagrama no link publicado", como
  combinado. `P-06.01` já vem com a URL real de exemplo
  (`https://fluxosti.codevasf.gov.br/incidentes/`, o mesmo fluxo que a AA/GTI
  usa na Base de Conhecimento) — **não é a URL final**, é só para visualizar
  como fica; troque quando o processo real tiver sua própria publicação. Os
  outros 20 itens (macroprocessos/processos/subprocessos) tiveram a imagem
  de exemplo removida e mostram honestamente "diagrama ainda não publicado".
- **"Normativos relacionados" removido** da ficha do Processo — já coberto
  por "Normativos e documentos vinculados".
- **Checklist**: novo tipo de documento, com um exemplo fictício (DOC-015,
  em P-06.01) só para você ver como fica — ainda não é padrão para todos os
  processos, mas o texto já sinaliza isso.
- **Subprocesso ganhou Entradas (insumos)**, além de Saídas (produtos) —
  antes só existia esse segundo campo, chamado "Entregas".
- **Nenhum campo/seção some mais quando vazio**: corrigido para sempre
  mostrar o título com "nada cadastrado" — isso incluía o card
  "Subprocessos deste subprocesso" (sumia inteiro quando o subprocesso não
  tinha filhos aninhados), o card "Objetivo" do Processo, e os cards de pai
  na Atividade e na Tarefa.
- **Novo card "Macroprocesso"** na ficha do Subprocesso, ao lado do card de
  pai direto (Processo ou outro Subprocesso) — acesso rápido ao topo da
  cadeia mesmo quando o subprocesso está vários níveis abaixo.

## Compartilhar e imprimir (em toda página)

Dois botões discretos aparecem no topo do conteúdo, em toda página do
painel (não é preciso trocar de tela para vê-los — ficam fora das views
individuais, então funcionam igual em qualquer rota, inclusive fichas).

- **Compartilhar**: usa o menu nativo de compartilhamento do
  sistema/navegador quando disponível (Web Share API); sem suporte, copia o
  link da página atual para a área de transferência (com uma confirmação
  visual); sem nenhum dos dois, mostra o link num diálogo para copiar à
  mão. Nunca falha silenciosamente.
- **Imprimir**: abre a caixa de impressão do navegador. O painel já tem uma
  folha de estilo própria para impressão — some com menu, rodapé, cookiebar
  e botões flutuantes, mostrando só o conteúdo da página.

## Novo material no Repositório

Adicionada a especificação oficial **BPMN 2.0 (Business Process Model and
Notation)** da OMG, usada por baixo dos panos em qualquer diagrama feito no
Bizagi: <https://www.omg.org/spec/BPMN/2.0/>.

## Correção: compartilhar/imprimir sem reagir ao clique

A navegação principal usa `position:sticky` para ficar fixa no topo ao
rolar a página. A barra de compartilhar/imprimir tinha sido colocada logo
depois dela, fora desse contexto sticky — ao rolar mesmo que um pouco, a
navegação fixa passava por cima da barra, tornando os botões inclicáveis
(o clique caía na navegação, não nos botões). Um bug real de sobreposição
visual, do tipo que testes automatizados não pegam por não fazerem
renderização de layout de verdade. Corrigido movendo a barra para **dentro**
de `#navigation`, no mesmo contexto sticky — agora rola junto com a
navegação e nunca fica coberta.

## Dois gráficos novos no Dashboard

A partir do catálogo de possibilidades do amCharts que você mandou, dois
tipos que ainda não usávamos e que encaixam bem com os dados que já temos:

- **Gauge (medidor)** — Avanço médio geral da carteira, com faixas de cor
  (vermelho/amber/verde) para leitura rápida do "estado geral" do
  portfólio, sem precisar comparar vários números.
- **Bubble (dispersão)** — Prioriza processos por avanço × riscos abertos:
  cada bolha é um processo, o tamanho reflete a quantidade de atividades
  mapeadas (contagem recursiva, considerando subprocessos aninhados),
  clicar leva direto à ficha. Processos com avanço baixo e algum risco
  aberto aparecem em vermelho — é o quadrante que merece atenção primeiro.

**Outros candidatos fortes do catálogo, para uma próxima rodada se
interessar:** um **Treemap** (categoria Hierarquia) mostrando a proporção
do portfólio por quantidade de tarefas em cada macroprocesso/processo; e um
**Gantt** (categoria Gantt) com o cronograma de mapeamento de cada
processo (início/prazo/conclusão), sinalizando os atrasados — os dados de
data já existem na planilha, só não estão visualizados como linha do tempo.

## Navegação hierárquica — um card por ficha, todos os ancestrais

Cada ficha (Processo, Subprocesso, Atividade, Tarefa) agora tem um único
card **"Navegar para"** no aside, reunindo o acesso direto a **todos** os
níveis acima do atual — não só o pai imediato. Por exemplo, na ficha de
uma tarefa: Macroprocesso, Processo, Subprocesso (a cadeia toda, se houver
subprocesso dentro de subprocesso) e Atividade, cada um numa linha
compacta e clicável, com ícone e cor por tipo. Substituiu os cards de
"pai" avulsos que existiam antes (um card por nível, repetindo "suba um
nível") — mesma informação, mas consolidada, mais rápida de escanear e sem
precisar subir um nível de cada vez para chegar ao topo da cadeia.

## Sistema de cores por camada

Cada camada da hierarquia recebe uma **família de matiz diferente** da paleta
oficial do gov.br DS — antes eram cinco tons do mesmo azul, indistinguíveis
lado a lado. Todas passam em AA (>=4,5:1) com texto branco por cima.

| Camada | Sigla | Família oficial | Cor | Gradiente da ficha |
| --- | --- | --- | --- | --- |
| Macroprocesso gerencial | MG | Blue Warm Vivid 80 | `#0C326F` azul | `#071D41 → #0040B4` |
| Macroprocesso finalístico | MF | Green Cool Vivid 50 | `#168821` verde | `#154C21 → #168821` |
| Macroprocesso de suporte | MS | Orange Vivid 50 | `#C05600` laranja | `#8C471C → #C05600` |
| Processo | PP | Cyan Vivid 50 | `#0081A1` turquesa | `#00687D → #0081A1` |
| Subprocesso | SP | Indigo Vivid 60 | `#4A50C4` índigo | `#373C93 → #4A50C4` |
| Atividade | AT | Yellow Vivid 60 | `#776017` bronze | `#5C4809 → #776017` |
| Tarefa | TR | Gray 80 | `#333333` grafite | `#1B1B1B → #555555` |

Os gradientes foram mantidos: cada um percorre **dois degraus da mesma
família**, o escuro no início e o cheio no fim — a matiz não muda no meio do
caminho, então a profundidade não custa a identidade da camada. Uma única
parada não é degrau publicado, `#373C93` (o Indigo Vivid 60 escurecido 25%),
porque a família índigo do DS não publica degrau abaixo do 60.

A mesma paleta vale para a cadeia de valor, o card "Navegar para", as tags e
os gráficos do dashboard: a série categórica passou a ser
`#0C326F · #168821 · #C05600 · #0081A1 · #4A50C4 · #776017 · #555555 · #B50909`,
matizes diferentes entre si em vez de quatro azuis.

**Siglas** — todas com duas letras, para que o nível se leia no próprio
código: `MG`/`MF`/`MS` (macroprocessos, por tipo), `PP` (processo), `SP`
(subprocesso), `AT` (atividade), `TR` (tarefa). O código da planilha
(`MP-`, `P-`, `A-`, `T-`) continua sendo a chave de vínculo — só a exibição
muda, então nenhuma planilha precisa ser regerada.

## Paginação (br-pagination)

Toda lista com mais de um item ganhou o componente Pagination do gov.br, com
os seis módulos — os cinco opcionais além das setas obrigatórias:

| ID | Módulo | Referência |
| --- | --- | --- |
| 1 | Setas de Navegação | Button (circular) |
| 2 | Identificadores de Páginas | Button |
| 3 | Botão Reticências | Button |
| 4 | Módulo de Exibição | Select |
| 5 | Módulo de Informação | Tipografia |
| 6 | Módulo de Atalho | Select |

Aplicado ao catálogo de processos, documentos, riscos, indicadores,
repositório, glossário e FAQ.

## Campos das fichas

- `Unidade responsável` e `Gerência responsável` viraram **Unidade Orgânica
  responsável** em todas as fichas.
- Novo campo **Unidades orgânicas corresponsáveis** (coluna
  `Unidades_Corresponsaveis`) em macroprocesso, processo, subprocesso,
  atividade e tarefa.
- Novo campo **Executor** na atividade (coluna `Executor`), preenchido com um
  cargo — ex.: "Analista em Desenvolvimento Regional".
- Saíram os campos `Dono do processo`, `Dono do macroprocesso` e `Dono`, e o
  `Responsável` da ficha da tarefa.
- `Procedimento Operacional Padrão (POP)` passou a **Procedimento (PRO)** em
  toda a base.

## Revisão da paleta de cores (gov.br DS)

Auditoria de **todas** as cores dos arquivos `css/govbr-ds.css`, `css/painel.css`
e `js/app.js` contra a Paleta do Design System gov.br (Fundamentos Visuais >
Cores) e contra o mínimo de contraste AA da WCAG 2.1 (4,5:1 para texto normal,
3:1 para texto grande e elemento gráfico), que é o nível de conformidade
adotado pelo DS.

**Tokens do DS corrigidos**

| Token | Antes | Depois | Motivo |
| --- | --- | --- | --- |
| `--gray-70` | `#4a4a4a` | `#555555` | valor publicado na família Gray |
| `--gray-30` | *não existia* | `#adadad` | era usado em 4 regras sem estar declarado |
| `--blue-warm-20` | literal `#c5d4eb` ×6 | token | Cor Interativa oficial para fundo escuro |
| `.br-switch --on` | `#538ff2` | `--blue-warm-vivid-70` | hex fora da paleta |
| `.br-switch.inverted` | `#cfe0fb` | `--blue-warm-vivid-10` | hex fora da paleta |
| `--focus` | *não existia* | alias de `--focus-color` | referenciado em 3 regras |
| `--noto-mono` | *não existia* | alias de `--font-mono` | referenciado em 3 regras |

Acrescentados os degraus oficiais que faltavam (`--gray-1/3/4`, `--pure-100`,
`--cyan-vivid-60`, `--orange-warm-vivid-50`) e as **cores de estado** do DS
(`--interactive`, `--danger`, `--warning`, `--success`, `--info`) — `--danger`
e `--success` já eram usados no CSS sem existir.

**Reprovações de contraste corrigidas**

| Onde | Antes | Depois |
| --- | --- | --- |
| Fase "Evoluir" da jornada | `#c45a00` — 4,37:1 | `--orange-warm-vivid-50` `#cf4900` — 4,56:1 |
| Fatia "Não iniciado" do donut | `#9e9e9e` — 2,68:1 | `--gray-50` `#757575` — 4,61:1 |
| Série ciano dos gráficos | `#74c9ea` — 1,86:1 | Cyan Vivid 60 `#00687d` — 6,41:1 |
| Série lima dos gráficos | `#89bd2b` — 2,24:1 | Green Warm Vivid 50 `#6a7d00` — 4,62:1 |
| Acento de hover da cadeia (finalístico) | `--pp-cyan` — 1,86:1 | `--cv-blue` — 6,78:1 |
| Acento de hover da cadeia (suporte) | `--pp-lime` — 2,24:1 | `--cv-green` — 5,19:1 |

**Mantido de propósito**

- **Marca Codevasf** (`--cv-blue`, `--cv-green`, `--cv-navy`, `--cv-cyan`,
  `--cv-lime`) e a assinatura gov.br do rodapé — identidade institucional, fora
  do escopo da paleta do DS. `--pp-cyan` e `--pp-lime` seguem declarados, agora
  com nota de que servem só a uso decorativo.
- **`--blue-warm-vivid-70: #0040b4`** — antecipação deliberada da v4.0 do DS
  (o valor em produção hoje é `#1351b4`).
- **Gradientes** — todos já usavam exclusivamente degraus oficiais da mesma
  família (`blue-warm-vivid` 90→70, `green-cool-vivid` 70→50, `gray-2`), então
  foram preservados: são o principal recurso de hierarquia visual das fichas e
  da hero, e nenhum deles reprova em contraste.
- **Tinturas próprias que passam em AA** — `#8a6d00`, `#137436`, `#c5170b`,
  `#0a6b80`, `#fff5d2`, `#e3f5e1`, `#fdeceb`, `.bg-lime`, `.bg-orange`, células
  da matriz de risco e `#31406f`. Não estão na paleta oficial, mas todas
  cumprem o contraste exigido na posição em que aparecem. Podem ser migradas
  para os degraus oficiais mais próximos numa próxima rodada, se quiser
  aderência total.

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
como projeto: termo de abertura, marcos M1–M10, entregáveis, riscos e lições
aprendidas). A aba **Metodologia** do painel documenta tudo isso para o público.

## Processo sem subprocessos: atividades e tarefas direto no processo

O CBOK 4.0 não obriga todos os níveis da decomposição ("Levels Vary in Number
and Name"). Há processo que não tem subprocesso e se decompõe direto em
**atividades** — e essas atividades têm **tarefas**, como sempre. O painel
agora trata esse caso como um caminho normal, não como exceção:

- Na aba **Atividades** da planilha, a coluna `Subprocesso` passou a se chamar
  `Vinculo_Pai` e aceita o código de um **Subprocesso** (`SP-...`) **ou** de um
  **Processo** (`P-...`). O painel identifica o nível pelo prefixo do código —
  mesma regra que a aba Subprocessos já usava. O nome antigo da coluna
  continua sendo lido, então planilhas ainda não atualizadas não quebram.
- A **ficha do Processo** ganhou o card *"Atividades ligadas direto ao
  processo"*, com a mesma tabela da ficha do Subprocesso (responsável,
  entradas, saídas, prazo e a contagem de tarefas de cada atividade). Quando o
  processo tem subprocessos e nenhuma atividade direta, o card explica onde as
  atividades estão, em vez de sumir.
- O card **"Subprocessos vinculados"** informa, quando está vazio e existem
  atividades diretas, que o processo se decompõe direto em atividades.
- **Breadcrumb** e card **"Navegar para"** (Atividade e Tarefa) montam a cadeia
  sem o nível que não existe: Macroprocesso › Processo › Atividade › Tarefa.
- A faixa da ficha do Processo mostra a contagem de subprocessos, atividades e
  tarefas — atividades e tarefas contadas **recursivamente**, somando as que
  penduram direto no processo e as de todos os subprocessos, em qualquer
  profundidade. O gráfico de bolhas do Dashboard usa a mesma contagem.
- **Marco M3 ("Subprocessos modelados")** aceita agora o valor
  **"Não se aplica"** na planilha (lista `Sim_Nao`), com ícone neutro e
  tooltip próprio na ficha. Sem isso, um processo sem subprocessos ficaria com
  um marco eternamente pendente.

### Como atualizar a planilha

A planilha entregue tem colunas acrescentadas depois da geração inicial (ex.:
`Objetivo` em Atividades e Tarefas). Por isso **não** rode
`gerar_planilha.py` sobre ela — esse script recria a planilha de exemplo do
zero. Use o script novo, que altera o arquivo existente no lugar, preservando
dados, colunas extras, formatação, fórmulas e validações:

```bash
python scripts/atualizar_planilha.py --exemplo   # estrutura + exemplo P-06.03
python scripts/planilha_para_js.py               # regenera js/dados.js
```

O `--exemplo` insere o caso demonstrativo: **P-06.03 (Gestão e Fiscalização
Contratual)**, um processo sem subprocessos, com duas atividades ligadas direto
a ele (`A-06.03.01`, `A-06.03.02`) e três tarefas da segunda. Esse exemplo já
está no `js/dados.js` desta versão, então o modo offline (abrir o
`index.html` com dois cliques) mostra o caso mesmo antes de rodar os scripts.

## Texto da Cadeia de Valor conforme o tamanho da tela

O parágrafo de apresentação da Cadeia de Valor tem três versões — curta, média
e completa — e o navegador mostra a que cabe na largura atual, usando os
breakpoints do gov.br DS (xs até 575,98px · sm/md 576–991,98px · lg a partir de
992px):

| Largura | Texto exibido |
| --- | --- |
| < 576px | Consulte a hierarquia completa, do macroprocesso à atividade. |
| 576–991,98px | …com fichas, diagramas BPMN, documentos, riscos e indicadores. |
| ≥ 992px | Texto completo, com "(Bizagi)" e o registro rastreável de cada mapeamento. |

A troca é feita em CSS (`.pp-tx-resp`), não em JavaScript: acompanha o
redimensionamento e a rotação do aparelho na hora, sem recarregar a tela. As
versões não exibidas saem do fluxo com `display:none`, então o leitor de tela
lê o texto uma única vez; na impressão vale sempre a versão completa. O helper
`txResp(curto, medio, completo)` do `js/app.js` está disponível para
qualquer outro texto que precise do mesmo comportamento.

## Alinhamento ao gov.br DS nesta rodada

O Fundamento de Superfície do DS gov.br v4 pede fundos **sólidos** no chrome de
interface — imagem, ilustração e trama ficam reservadas a conteúdo editorial,
estados vazios e telas de erro. Os gradientes decorativos do
`css/painel.css` foram trocados por cor plana da paleta institucional:

- `.pp-hero`: gradiente de três paradas → `--cv-navy`; a trama de pontos
  (`radial-gradient` repetido) foi removida.
- `.cadeia`: gradiente claro → `--gray-2`.
- `.cv-aside`, `.cv-valores` e os três títulos de bloco da cadeia
  (`.cv-bloco.cat-*`): gradiente → cor da categoria (navy, azul, verde).

Nada mais mudou no visual: tipografia (Noto Sans/Noto Sans Mono), escala de
espaçamento, raios, elevação, foco laranja, ícones Font Awesome 6 e as cores
semânticas (`#168821` sucesso · `#B38C00` alerta · `#E52207` erro ·
`#155BCB` informação) já seguiam os tokens da v4.

## Ajustes desta rodada (tooltips, marcos, tutorial e NUGEP)

- **Tutorial de abertura**: no lugar do link "Pular tutorial", dois botões
  lado a lado no rodapé do tour — **Pular** (secundário) e **Iniciar**
  (primário) —, o mesmo par da barra de navegação do Wizard da jornada. O
  primário acompanha o passo: *Iniciar* → *Avançar* → *Concluir*; "Pular"
  sai de cena na última página. Os botões que ficavam soltos dentro da
  primeira e da última página saíram (faziam a mesma coisa).
- **Jornada de mapeamento**: removida a faixa "Descobrir → Definir →
  Desenvolver → Entregar → Evoluir" entre o título e o Wizard — a jornada
  já se explica no próprio Wizard.
- **Marcos do mapeamento (M1–M10)**: os dez marcos ficam **todos à vista**
  na ficha do processo, sem barra horizontal para arrastar (o Step perdeu o
  `data-scroll`); abaixo de 992px a trilha assume a orientação **vertical**
  do componente, em vez de espremer dez colunas. Na bolinha do marco
  concluído aparece agora o **número do marco**, não o check — a bolinha
  continua pintada de azul.
- **Tooltip padrão gov.br em todo o painel**: o que antes só aparecia no
  `title` nativo do navegador (KPIs da tela inicial e do Dashboard, marcos,
  legendas dos gráficos, pinos do radar de riscos, botões flutuantes) virou
  o Componente Tooltip, acionado por um ícone **"i"** discreto ao lado do
  rótulo. Ganho de acessibilidade real: o `title` nativo não abre no foco
  do teclado nem é lido de forma confiável por leitor de tela. O balão
  agora também se desloca sozinho para não sair da tela e usa
  `width:max-content` (antes herdava a largura do gatilho e descia em
  coluna de uma palavra por linha).
- **NUGEP — Contato institucional**: o bloco passou a mostrar a
  **hierarquia da unidade** — Gerente-Executivo (AE), Gerente (AE/GPE) e a
  equipe da AE/GPE/UNP —, cada pessoa com **avatar (foto), nome, papel,
  e-mail e telefone**. As chefias aparecem só nesse bloco; a lista de
  integrantes do núcleo continua como estava.
- **Foto na planilha**: a aba NUGEP ganhou as colunas **Foto** (URL de
  imagem pública, que substitui as iniciais do avatar) e **Hierarquia**
  (1 = Gerente-Executivo, 2 = Gerente, 3 = equipe da Unidade, 0 =
  interlocutor de outra área). `scripts/gerar_planilha.py` já gera as duas
  colunas; enquanto a planilha em uso não as tiver, o painel aproveita as
  fotos de exemplo de `js/dados.js`, casando pelo e-mail. As URLs atuais
  são retratos de exemplo (randomuser.me) só para teste — troque pelas
  fotos oficiais.
- **Botão "Reportar um erro" volta a flutuar**: quem fica fixo agora é o
  contêiner (`.report-error-wrap`). O botão é também gatilho de tooltip e
  precisa de `position:relative`; a regra `.br-button.circle` reimpunha
  esse relative e derrubava o botão para o fim do corpo da página, onde
  virava uma faixa clara abaixo do rodapé em vez de flutuar.
- **Faixa de seções**: as ações de página (alertas, compartilhar,
  imprimir) perderam o `max-width`/`margin:0 auto` herdados de quando
  viviam numa linha própria — era isso que jogava o bloco por cima do botão
  **"Mais"** e do nome da seção aberta. Saiu junto o `border-top`, que
  desenhava o **divisor extra** acima do sino. O recorte das abas passou a
  reservar 8px de folga, para o botão "Mais" nunca ficar cortado.
- **Sistema de Grid**: a faixa de seções e o rodapé estavam em containers
  de 1440px enquanto cabeçalho e conteúdo usavam 1200px (1560px em TV).
  Agora todos usam o mesmo container — pelo Fundamento Grid, o conteúdo
  vive dentro do mesmo container em toda a página.
- **Estados**: os novos controles (ícone "i", botões do tutorial, etapas da
  trilha de marcos) seguem o Fundamento Estados — cor no interativo,
  overlay no hover, contorno de foco com os tokens `--focus-*` e alvo de
  toque de 24px mesmo com ícone de 12px.
- **Últimos hover nativos**: os botões só-ícone de densidade da tabela,
  exportar seleção e limpar seleção também trocaram o `title` pelo Tooltip —
  não sobrou nenhum `title` nas telas do painel.
- **Logo do gov.br no rodapé**: a URL externa passou a responder 404; a marca
  agora vem de `img/govbr-negativa.svg`, dentro do próprio repositório.

## Ajustes desta rodada (período, ilustração, utilitários e movimento)

- **Filtro "Prazo previsto entre" (Portfólio)**: o calendário fechava sozinho
  depois da primeira data — o clique redesenha a grade e, quando o evento
  chegava ao `document`, o botão clicado já não estava no DOM, então o
  "clique fora fecha" disparava. Corrigido. Além disso, o modo intervalo
  ganhou **dica de passo** ("Passo 1 de 2: escolha a data inicial" →
  "Passo 2 de 2: escolha a data final"), **prévia do intervalo** enquanto o
  cursor percorre os dias e **extremos arredondados** no trecho selecionado.
  O rótulo do campo ganhou o "i" explicando que são duas datas.
- **Ilustração (Fundamento)**: as ilustrações oficiais do pacote gov.br
  foram copiadas para `img/ilustracoes/` e aplicadas nos três cenários que o
  fundamento indica — **tutorial de abertura** (5 páginas), **estados vazios**
  (portfólio, documentos, repositório, NUGEP, glossário e busca) e **tela de
  erro** de carregamento dos dados.
- **Utilitários dos fundamentos**: publicados como classes reais, mapeadas
  nos tokens já existentes — Superfície (`border-solid-*`, `border-dashed-*`,
  `rounder-*`, `opacity-*`, `overflow-*`, `d-*` e flexbox), Elevação
  (`shadow-*`, `shadow-*-inset`, `layer-0..4`), Movimento (`ease`, `ease-in`,
  `ease-out`, `ease-in-out`, `linear` + `very-fast`…`very-slow`) e Tipografia
  (`text-up-*`/`text-down-*`, `text-weight-*`, transformação, alinhamento e
  quebra). Junto vieram os tokens com os nomes oficiais
  (`--surface-*`, `--duration-*`, `--animation-ease-*`, `--font-size-scale-*`,
  `--font-weight-*`, `--font-line-height-*`) apontando para os que já existiam.
- **Movimento**: durações e atenuações soltas passaram a usar os tokens, e o
  painel inteiro passou a respeitar `prefers-reduced-motion` — a preferência
  de menos movimento declarada no sistema operacional desliga transições e
  animações, como pede o fundamento (distúrbios vestibulares, epilepsia,
  déficit de atenção).
- **Tipografia na escala oficial**: os tokens de tamanho passaram a usar os
  degraus da escala **Minor Third (1,2)** sobre a base de 14px do DS —
  `--fs-sm` 11,67px (down-01), `--fs-base` 14px, `--fs-md` 16,8px (up-01),
  `--fs-lg` 20,16px (up-02), `--fs-xl` 24,19px (up-03) e `--fs-xxl` 29,03px
  (up-04). Antes o painel usava uma escala própria (12/16/20/24/32), próxima
  mas fora da progressão. Como tudo se apoia nesses tokens, a mudança é sutil
  (menos de 1px na maioria dos casos; o maior efeito é no título da capa) e
  alinha o painel inteiro à mesma escala.

## Conferência contra o código-fonte oficial (govbr-ds v4.0.0-next.32)

Os utilitários publicados na rodada anterior tinham sido escritos a partir da
documentação da **v3.7**. Com o repositório oficial em mãos
(`projects/tokens/src/figma/*.json` e `projects/design/src/utilities/*.scss`),
os valores foram conferidos e corrigidos:

- **A v4 renomeou boa parte das classes.** O painel agora aceita os **dois**
  nomes — o oficial da v4 e o da v3.7 que já estava em uso:

  | v3.7 | v4 (oficial) |
  | --- | --- |
  | `border-solid-sm/md/lg` | `border-solid` + `border-width-sm/md/lg` (e `border-dotted`) |
  | `rounder-*` | `rounded-none/sm/md/lg/full` |
  | `opacity-xs…xl` | `opacity-10…90/full` |
  | `shadow-*` + `layer-*` | `elevation-0…4` (sombra e camada na mesma classe) |
  | `ease-in`, `fast` | `easing-ease-in`, `duration-fast-1…4` |
  | `d-flex`, `d-block` | `display-flex`, `display-block` |
  | `text-up-*`, `text-weight-*` | `font-size-up-*`, `font-weight-*` |

- **Elevação estava errada.** Cada nível do DS soma **duas** sombras — "cast"
  (projetada, com deslocamento vertical) e "occlusion" (de contato, sem
  deslocamento e mais difusa) —, ambas em preto a 20% no tema claro. Os valores
  agora saem dos tokens: nível 1 `0 1px 1px + 0 0 4px`; nível 2 `0 4px 4px +
  0 0 8px`; nível 3 `0 8px 8px + 0 0 12px`; nível 4 `0 12px 12px + 0 0 16px`.
  A classe `elevation-N` aplica sombra **e** `z-index` (0, 1000, 2000…).
- **Movimento**: a escala real tem quatro degraus por faixa — `fast` 50/100/150/200ms,
  `moderate` 250/300/350/400ms, `slow` 450…600ms, `very-slow` 700…1000ms —, mais as
  atenuações `bounce-out` e `elastic-out`, que faltavam.
- **Superfície**: `rounder-lg` era 16px por estimativa; o token oficial é **12px**.
  Também entraram `position-*`, `overflow-visible` e `display-grid`.
- **Tokens** passaram a existir com o nome oficial (`--br-surface-*`,
  `--br-elevation-*`, `--br-motion-*`, `--br-typography-*`), com os nomes antigos
  mantidos como apelido.

### Padrão Dados Ausentes aplicado aos estados vazios

A diretriz de **Dados Ausentes** define a anatomia do estado vazio: Apoio
Visual (ilustração, opcional), **Título** (obrigatório quando o vazio ocupa a
tela ou um bloco grande), **Mensagem** (obrigatória) e **Suporte para Ações**
(opcional, ao final do conteúdo). Os estados vazios do painel tinham só
ilustração e uma frase; agora têm título e mensagem, e o portfólio e a busca
ganharam ação — "Limpar filtros" e "Ver o portfólio", em ênfase secundária,
como manda a hierarquia do componente Button. Os textos seguem o tom que a
diretriz pede: neutro, específico quanto ao critério usado e orientando o
passo seguinte.

### Identidade visual: o painel segue na v3.7

A v4 muda a tipografia (Rawline → **Noto Sans**), o azul institucional
(`#1351b4` → `#0040B4`) e a escala de tamanhos (Minor Third sobre 14px →
Major Second 1,125 sobre 16px). **Não migramos** — e por um motivo prático: as
páginas gov.br no ar ainda são v3.7 (o próprio site do DS serve `theme-color:
#1351b4`), e a v4 está em pré-lançamento (`next.32`). Um painel institucional
que destoasse das outras páginas da Codevasf e do gov.br pagaria caro por
antecipar o salto. Os valores da v4 ficam registrados nos tokens, prontos para
quando a migração for decidida.

## Revisão das diretrizes v4 — fundamentos e padrões (1ª leva)

Leitura arquivo por arquivo de `projects/diretrizes`, com o que foi corrigido:

- **Estado**: o estado **Pressionado** deixou de ser mudança de cor e passou a
  ser **sombra interna** (`state.pressed.shadow`: deslocamento 1px, suavidade
  8px, preto a 50%) — é a mudança que a v4 fez no fundamento. O **Foco** ganhou
  o espaçamento de segurança correto (4px = `spacing.adjust.1`; estava 2px), e
  o **Desabilitado** passou a remover a sombra, como a diretriz manda.
- **Iconografia**: os nomes dos ícones foram padronizados pelos tokens
  `br.iconography.iconName.*` — o painel misturava nomes canônicos do FA6 com
  apelidos (`fa-triangle-exclamation` x `fa-exclamation-triangle`,
  `fa-up-right-from-square` x `fa-external-link-alt`, `fa-circle-info` x
  `fa-info-circle`). 13 ocorrências alinhadas; visualmente idênticas, agora
  rastreáveis até o token.
- **Tooltip (acessibilidade)**: o gatilho passou a ser associado ao balão por
  `aria-describedby` (o balão complementa, não nomeia); o balão ganhou a
  carência de 400ms antes de fechar e continua aberto enquanto o cursor
  estiver sobre ele — exigência do WCAG 2.2 · 1.4.13, que o fechamento
  imediato violava. A sombra passou a ser a da camada 4, como pede
  `br.tooltip.shadowStyle`.
- **Content Overflow**: o texto truncado do card "Navegar para" passou a dar
  acesso ao conteúdo completo (a diretriz proíbe truncar sem saída). As
  tabelas largas ganharam a **sombra de rolagem** (técnica 1.3, *scroll
  affordance*): a sombra aparece na borda enquanto há conteúdo além da área
  visível e some ao chegar ao fim.
- **Espaçamento**: escala conferida contra os tokens — layout 8…80 bate; a
  escala de ajuste oficial vai só até 36 (o painel tinha 44 e 52 a mais, agora
  documentados como extensão). Faltava o degrau 80, incluído.
- **Densidade**: o menu das tabelas usava "Densidade alta/média/baixa"; a
  classificação do DS é **Compacta / Regular / Espaçada**. Renomeado nas duas
  tabelas, e o balão do gatilho, que dizia "Ver mais opções", passou a dizer o
  que o menu faz.
- **Writing**: nos *buttons*, a diretriz pede iniciais maiúsculas em todas as
  palavras (exceto artigos e preposições) — "Limpar Filtros", "Ver o
  Portfólio", "Abrir em Nova Aba". Rótulos de campo, itens de menu e links
  seguem em sentença, como a diretriz também pede.
- **Dropdown**: conferido, sem correção — o painel já usa elevação a partir da
  camada 2, inverte a direção quando falta espaço, mantém o estado no
  acionador enquanto aberto e dispensa o identificador nos gatilhos cujo
  ícone já carrega a semântica (uso opcional previsto na diretriz).

## Revisão das diretrizes v4 — 2ª leva (identificadores e conferências)

- **Collapse e Dropdown — identificadores trocados**: o painel usava `angle-down` /
  `angle-up` nos dois padrões, e esse nome não existe em nenhuma das duas tabelas
  do fundamento Iconografia. A diretriz reserva o **chevron** ao identificador de
  *collapse* (accordion, grupos do cookiebar, expansão de linha de tabela, blocos
  do rodapé) e o **caret** ao identificador de *dropdown* (o Select). Trocado nos
  13 pontos, no HTML inicial e na alternância em JS — o ícone continua apontando
  para a direção em que o conteúdo abre, como a diretriz exige.
- **Densidade**: nomes alinhados à classificação do DS (Compacta / Regular /
  Espaçada).
- **Formulário**: conferido, sem correção. Rótulos em sentença, sem dois-pontos e
  acima do campo; par de botões com a ênfase primária à direita da secundária; sem
  botão "Redefinir"/"Limpar formulário" em campos de entrada (o "Limpar Filtros" é
  de painel de filtros, uso que a diretriz de Dados Ausentes recomenda).
- **Writing (fechamento)**: a titulação dos *buttons* ficou completa — "Limpar
  Filtros" (no rodapé de filtros e no estado vazio, a mesma ação com a mesma
  grafia), "Página Inicial", "Definir Cookies", "Baixar Print", "Tentar Capturar
  de Novo", "Abrir E-mail". Ficam em sentença apenas os rótulos que embutem o
  nome de outro elemento ("Abrir Notícias") e os nomes próprios ("Atalhos
  gov.br").
- **Button (acessibilidade)**: auditado ao vivo — 100% dos botões visíveis têm
  nome acessível e nenhum fica abaixo do alvo mínimo de 24px (WCAG 2.5.8).

### Ainda por revisar

Padrões *skeleton-screen*, *tema*, *sign-in*, *slot-de-composição*, *design
conversacional* e *internacionalização*; e os componentes Input, Select, Avatar,
Header, Footer, List, Message e Loading.

## Revisão das diretrizes v4 — componentes

Auditoria feita no painel em execução, contra as listas de verificação de
acessibilidade de cada componente:

- **Message — corrigido**: a diretriz exige que o leitor de tela anuncie *qual* é
  o tipo da mensagem (erro, aviso, sucesso, informação — WCAG 4.1.2). O ícone que
  carrega essa informação é `aria-hidden`, então o leitor lia só o corpo do texto.
  Um `BRMessageInit` passa a inserir o rótulo invisível ("Aviso: ", "Sucesso: ")
  em toda mensagem, presente e futura; os `role="alert"`/`role="status"` já
  cobriam o anúncio dinâmico.
- **List — corrigido**: a lista de classes de cookies era uma `div.br-list` sem
  `role` — não era anunciada como lista nem informava a contagem de itens. Ganhou
  `role="list"` e nome acessível.
- **Button**: 100% dos botões visíveis têm nome acessível e nenhum fica abaixo do
  alvo mínimo de 24px (WCAG 2.5.8).
- **Input**: 53 campos, nenhum sem rótulo associado — nenhum depende de
  *placeholder* como rótulo.
- **Select**: `role="combobox"` com `aria-expanded` e `aria-controls` apontando para
  o `role="listbox"` correspondente.
- **Avatar**: decorativo (`aria-hidden`) com o nome da pessoa visível no card —
  a foto nunca é o único portador da informação.
- **Header / Footer**: um `banner`, um `main` e um `contentinfo` por página; as
  navegações têm nome próprio, e as que usam `role="none"` (trilhas de Step)
  colocam o rótulo no `listbox` interno, como o componente prevê.
- **Loading**: `role="progressbar"` com rótulo, acompanhado de texto em
  `role="status" aria-live="polite"`.

## Ferramentas herdadas do repositório do DS

O monorepo do Padrão Digital de Governo traz um conjunto de ferramentas de
qualidade. Foram avaliadas uma a uma; entraram as que funcionam num projeto
estático, sem etapa de build:

| Arquivo | Decisão |
| --- | --- |
| `biome.json` | **Adotado**, adaptado. Formata e linta o JS do painel. Rode `npx @biomejs/biome check js` (ou `--write` para corrigir). O escopo virou `js/**/*.js`, o `$schema` passou a apontar para a URL pública (não havia `node_modules`) e as regras de import/`any`/complexidade saíram: o painel é ES5 clássico, sem módulos, e as funções de render são longas por natureza. |
| `cspell.json` + `cspell-wordlist.txt` | **Adotados**, adaptados. Corretor ortográfico com o vocabulário do DS (cookiebar, skiplink, rawline, focável, rolável…), acrescido dos termos do painel (macroprocesso, NUGEP, CBOK, Bizagi…). Saíram os dicionários `pt-br` e `lorem-ipsum`: são pacotes npm separados que o `cspell` não embute, e mantê-los faria o comando falhar na resolução antes de checar qualquer palavra — o vocabulário pt-BR que interessa já está na lista local. Rode `npx cspell --config cspell.json`. |
| `LICENSE` | **Adotado** como `LICENSES-TERCEIROS.md` — o painel deriva CSS, JS e ilustrações do DS, e o crédito MIT/CC0 precisa acompanhar o código. |
| `package.json`, `nx.json`, `postcss.config.cjs`, `release.config.js`, `commitlint.config.js`, `cz.config.cjs`, `lint-staged.config.mjs` | **Não adotados.** São a infraestrutura de um monorepo pnpm/Nx com semantic-release: build de pacotes, versionamento automático e padrão de mensagem de commit. O painel é servido como arquivo estático e abre sem instalar nada — trazer isso custaria Node, pnpm e ~30 dependências de desenvolvimento para não mudar uma linha do que o usuário vê. |
| `AGENTS.md`, `AI-CONTEXT.md`, `CLAUDE.md`, `GEMINI.md` | **Aproveitados** no `CLAUDE.md` da raiz do projeto, com as regras que valem aqui: tokens como fonte da verdade, preservar nomes públicos de classe, WCAG 2.2 AA como linha de base, menor diff possível. |
| `CONTRIBUTING.md`, `CODE_OF_CONDUCT.md`, `SECURITY.md` | **Não adotados.** Governam a contribuição ao repositório do DS, não a este painel. |

As duas ferramentas adotadas são opcionais: o painel continua abrindo e
funcionando sem instalar nada. Elas só entram em cena se alguém quiser
conferir o código antes de publicar.

## Revisão das diretrizes v4 — padrão Tema

- **O modo de alto contraste é um tema, e está conforme.** A diretriz permite que
  um tema altere **apenas cores** — espaçamento, movimento e grid ficam de fora.
  Conferido: `body.contraste-alto` não toca em nenhuma propriedade que não seja cor.
- **Cores de estado derivadas do Interativo.** O padrão define cada cor de estado
  a partir da cor Interativo (aqui Blue Warm Vivid 70). Agora estão declaradas com
  nome próprio, para nenhum componente ter de redescobrir a regra: Visitado
  (mesma família e grau, círculo padrão em vez do Vivid), Desabilitado (cinza no
  mesmo grau), Ativo (±1 grau), Selecionado (±2), Ligado (±3) e Foco (família
  complementar, grau 40 — o laranja que o painel já usava).
- **Hover no modelo novo.** A v3 usava uma opacidade única sobre a cor
  interativa; a v4 usa **dois overlays** — claro sobre elemento escuro, escuro
  sobre elemento claro —, escolhendo o que gera mais contraste com o primeiro
  plano. O Accordion, que ainda estava no modelo antigo, passou ao novo (overlay
  escuro na versão positiva, claro na negativa).

## Revisão das diretrizes v4 — padrões restantes

- **Skeleton Screen — aplicado.** A tela inicial mostrava um giro sem contexto
  enquanto a planilha carregava. Agora desenha a **forma** do que vem a seguir —
  título, faixa de indicadores e cartões —, com os valores dos tokens
  `br.skeleton`: fundo `background.main.secondary` com canto `rounder.sm`, brilho
  `background.main.primary` com canto `rounder.full`, altura de linha igual ao
  `fontSize.default` e movimento `linear` em `600ms` (`duration.slow.4`). O usuário
  entende o que está carregando e a troca pelo conteúdo real não desloca a
  página. O esqueleto é decorativo (`aria-hidden`); quem anuncia o carregamento
  continua sendo o texto em `role="status"`.
- **Internacionalização — um ajuste.** O painel já cumpria o essencial:
  `lang="pt-BR"` no `<html>`, UTF-8 e datas por `toLocaleString('pt-BR')`. Faltava o
  tamanho de arquivo do Upload, que saía de `toFixed(2)` e portanto sempre com
  ponto decimal ("1.50 MB"); passou a usar `Intl.NumberFormat('pt-BR')` e agora
  respeita a vírgula ("1,50 MB").
- **Sign-in — não se aplica.** O painel é de consulta pública, sem autenticação.
  O componente existe no CSS para o caso de vir a ter login.
- **Slot de Composição — não se aplica.** É um padrão de API de componente
  (como um componente aceita conteúdo de fora), pensado para bibliotecas em
  framework. Aqui o HTML é montado direto em JS.
- **Design Conversacional — não se aplica.** Trata de chatbots e assistentes;
  o painel não tem interface de conversa.

## Componente Loading — arte oficial incorporada

O anel de carregamento era desenhado à mão em CSS (uma borda com um lado
transparente). Os arquivos oficiais do componente entraram no lugar, em
`img/loading/` — `loading-global-{small,medium,large}-{main,contrast}.svg`, seis
no total. Eles já são desenhados para receber a cor por custom property
(`--br-loading-color-background-*` e `--br-loading-color-foreground-*`), então o
mesmo traçado serve ao tema claro e ao alto contraste; o painel só precisa
apontar as quatro variáveis para os seus tokens.

Detalhes da incorporação:

- Os tamanhos passaram a ser os dos próprios arquivos: **19 · 40 · 72px**
  (antes 24 e 44, estimados).
- O caminho do SVG é embutido em `js/govbr-ui.js` e injetado por
  `BRLoadingInit` — e não buscado por `fetch` — porque o painel também roda
  aberto direto do arquivo (`file://`), onde `fetch` é bloqueado. Os `.svg`
  ficam no repositório como origem do traçado.
- O filtro de sombra interna que vinha do Figma foi descartado: seus ids
  colidiriam entre vários Loadings na mesma página e o efeito é
  imperceptível no tamanho em que o componente roda.
- O anel em CSS continua no arquivo como reserva, ativado por `:empty` —
  vale enquanto o script não rodou e no modo determinado (percentual).
- As variantes **conversational** não entraram: são do padrão Design
  Conversacional, que o painel não usa.
- Os fills usam nomes **neutros** (`--br-loading-color-background` /
  `--foreground`), e quem troca o tema é o CSS. Decidir a variante no
  momento da injeção não funcionava: o alto contraste do painel é um
  toggle de classe em runtime e o Loading nasce no render da rota, então a
  arte clara ficava congelada. No alto contraste o arco passa ao dourado
  (`--gold-vivid-20`): o azul interativo sobre preto dá 2,39:1, abaixo do
  mínimo de 3:1 para elemento não textual (WCAG 1.4.11); o dourado dá
  12,66:1. A troca cobre a arte oficial e o anel de reserva de uma vez.
- O Loading do Upload nasce depois do init da tela, então recebe a arte
  com uma chamada própria a `BRLoadingInit`.

## Correções pedidas nesta rodada

- **Tutorial — texto "na lateral"**: na página 2 o parágrafo aparecia à direita,
  fora do modal. A regra do componente (`.br-carousel .carousel-content`) tem a
  mesma especificidade da regra do tutorial e vencia por ordem, reimpondo
  `flex-wrap:wrap` — num contêiner de altura fixa, o que não cabia transbordava
  para uma **segunda coluna**. Corrigido qualificando a regra pelo carrossel.
- **Tutorial — tamanho fixo**: altura única de 452px, dimensionada pela página
  mais alta, para o modal não mudar de tamanho a cada passo.
- **"Não mostrar novamente"** saiu do rodapé do tutorial; fechar o tour já vale
  como visto (o tour continua acessível pelo menu de ajuda).
- **Cookiebar**: faixa mais baixa e texto de ponta a ponta — o `max-width` em `ch`
  vinha da medida de leitura de texto corrido e deixava metade da linha vazia
  numa faixa que atravessa a tela.
- **Links do cookiebar que não respondiam**: a área de conteúdo não encolhia
  para caber na faixa, então crescia com o conteúdo e empurrava os três links
  (LGPD, política do gov.br, falar com a UNP) para **fora da tela**. Com
  `flex:1 1 auto` e `min-height:0` o conteúdo rola e os links voltam ao alcance.
- **Botão Fechar (X) dos modais** aparecia embaixo do texto:
  `.br-button.circle{position:relative}` (regra do alvo de toque, no fim do
  arquivo) tem a mesma especificidade de `.br-modal-header .close` e vencia por
  ordem — o botão voltava ao fluxo. Mesmo tropeço do botão flutuante de
  reportar erro; aqui a correção foi qualificar a regra pelo modal.
- **Logo gov.br do rodapé em preto**: os caminhos do SVG só tinham
  `class="st0"`, sem nenhuma regra que a definisse, e o navegador caía no preto
  padrão. Como o arquivo entra por `<img>`, `currentColor` não chega — a cor foi
  para dentro do próprio SVG (branco, que é a definição da versão negativa).
- **Rodapé**: a linha "Unidade de Gestão Normativa e de Processos — AE/GPE/UNP"
  saiu (também em `scripts/montar_index.py`, senão voltaria na próxima geração).
- **Textos de apoio removidos**: o tooltip do filtro "Prazo previsto entre" (o
  campo já se explica sozinho depois da correção do calendário), "Vale para os
  cartões e para a paginação abaixo" e "Selecione um ou mais status…".
- **Tooltip do botão de densidade**: o balão ficava aberto por cima do menu que
  o próprio botão abre, tapando a primeira opção. Agora sai de cena no clique.
- **Colunas das tabelas**: `.col-curta` encolhe a coluna ao conteúdo
  (`width:1%`), o que espremia códigos e rótulos curtos em duas ou três linhas.
  As colunas ganharam largura mínima — ID de riscos 92px, ID de indicadores
  100px, Nível 104px, Status 124px, Código de tarefas e subprocessos 132px,
  "Vinculado a" 200px.
- **"Vinculado a"** perdeu o rótulo entre parênteses ("(Macroprocesso)",
  "(Processo)"): a sigla de duas letras no início do link já diz o nível.
- **Fundo do painel**: era cinza claro com cards brancos, o que fazia cada card
  e cada collapse aparecerem como um retângulo mais claro — o degrau que se via
  no Glossário e no FAQ. Agora o fundo é branco (`--background` do DS) em toda
  parte, e a separação fica por conta da borda e da sombra, que é como o DS
  separa superfícies.
- **Filtros do Glossário e do FAQ** passaram ao Componente Tag (interação
  persistente): mesma altura de 24px, canto arredondado, cor interativa e foco
  visível — eram dois controles desenhados à mão, cada um com um raio e uma
  altura.
- **Avatar do NUGEP** no tamanho grande do componente (64px no card, 56px no
  bloco institucional).
- **Check dos marcos** desalinhado: o `margin-left:auto` do ícone não tinha folga
  para consumir — o rótulo ocupava a linha inteira e o check parava logo depois
  do texto, em posição diferente a cada marco. Com `flex:1` no rótulo, os dez
  checks voltam à mesma coluna.
- **Marcos M1–M10 em duas linhas de cinco** (a partir de 992px): dez colunas
  numa linha só reduziam cada etapa a uma faixa estreita, com o rótulo
  quebrando em três ou quatro linhas.

## Conferência contra os tokens Base14 e o Dicionário

- **Escala tipográfica conferida contra o modo Base14 oficial**
  (`tokens/Core/Base14 → typography.json`): 14 · up-1 15,75 · up-2 17,72 ·
  up-3 19,93 · up-4 22,43 · up-5 25,23 · up-6 28,38 · up-7 31,93 · up-8 35,92;
  down-1 12,44 · down-2 11,06 · down-3 9,83 · down-4 8,74. O painel tinha 31,94
  em `--fs-xxl` (arredondamento próprio) — corrigido; entraram os degraus up-8 e
  down-4, que faltavam. O manifesto dos tokens confirma que **Base14 é um modo
  oficial** da v4, ao lado de Base16 — manter a base 14 nesta interface de
  trabalho está dentro do padrão.
- **Tema declarado no elemento raiz**, como manda o Guia de uso da v4:
  `data-theme="light"` no `<html>`. O Modo Contraste passou a marcar também
  `data-contrast="enabled"` ali — o contrato público da v4 —, mantida a classe
  `.contraste-alto`, que é o que as regras do painel usam.
- **Dicionário de Vocabulário Controlado**: em fluxo de *onboarding*, onde o
  usuário lê de forma passiva, o termo é **"Próximo"** — "Avançar" pertence a
  fluxos de tarefa, com o usuário ativo; o botão do tutorial foi corrigido. E
  **"Pesquisar"** substituiu "Buscar" nos 11 pontos de consulta (rótulo de
  campo, placeholder e nome acessível do botão), porque "Buscar" não está no
  dicionário.
- **Font Awesome**: a v4 pina a versão 7.0.1; o painel segue na 6.5.2. Decisão
  deliberada — todos os ícones em uso renderizam na 6.5.2 (conferido), e trocar
  de versão maior arrisca apagar ícones por renomeação, sem ganho visível.

## Ajustes de tabela, avatar e planilha

- **Foto do avatar preenchendo o círculo**: a imagem entrava no tamanho natural,
  encostada no canto superior esquerdo, e deixava as iniciais à mostra em volta.
  Agora ocupa 100% do círculo com `object-fit:cover` — recorta pelo centro e cobre
  em qualquer proporção de foto.
- **Coluna "Risco"**: a descrição do risco em negrito, o tratamento logo abaixo em
  texto de apoio — dá para separar os dois de relance.
- **Coluna ID de Documentos publicados**: largura mínima de 108px (a tabela não
  tem `colgroup`, então a medida foi para o `th`).
- **Coluna Nível** da tabela de riscos: de 104px para 128px, para "Extremo" e
  "Moderado" não quebrarem em duas linhas.
- **Balão dos três pontinhos**: o texto saía por fora do fundo azul. Os gatilhos
  das tabelas vivem dentro de `.actions-trigger.text-nowrap`, e o `nowrap` descia
  para dentro do balão — o texto seguia numa linha só, passando dos 240px da
  caixa. O balão é superfície própria e agora quebra linha sozinho; a largura
  máxima subiu para 280px e o texto virou título + explicação em duas linhas.
- **"Próxima ação" no Componente Message**: era um parágrafo solto. Agora é uma
  mensagem informativa com ícone e `role="status"` (e um texto próprio quando não
  há ação programada, em vez de um travessão); a **Pendência** virou mensagem de
  aviso com `role="alert"`. As duas herdam o rótulo de tipo para leitor de tela.

### Planilha pronta para receber os dados reais

A planilha já vinha com cabeçalho fixo, filtro automático, listas suspensas,
colunas de fórmula em cinza e largura por coluna. O que mudou para o
preenchimento render mais:

- **Listas suspensas que avisam**: a validação estava com `showErrorMessage=False`
  — um valor digitado fora da lista passava em silêncio e só quebrava depois, no
  painel. Agora a célula recusa o valor com uma mensagem que diz o que fazer
  ("Escolha uma das opções da seta ao lado da célula"), no tom do Padrão Writing:
  sem culpar quem preencheu. Ao entrar na célula, uma dica anuncia a lista.
- **Listra zebrada** nas linhas de dados: com abas de até 37 colunas, é o que
  impede de trocar de linha no meio do preenchimento.
- **Cor de guia por grupo**: azul para a hierarquia (Macroprocessos → Tarefas),
  verde para os anexos (Documentos, Riscos, Indicadores), azul claro para o
  conteúdo editorial (Jornada, Repositório, NUGEP, Glossário, FAQ) e cinza para
  as abas de apoio (Parâmetros, Listas). A barra de guias é a única navegação da
  planilha, e são 16 abas.
- **Zoom em 90%**: as abas largas cabem sem rolagem horizontal constante.
- **Azul do cabeçalho** atualizado para o `#0040B4` da v4 (era o `#1351B4` da v3).
- A fonte segue **Arial**, e não Noto Sans: a planilha é preenchida no
  Excel/LibreOffice de cada pessoa, onde a Noto Sans não está instalada — fonte
  ausente vira substituição imprevisível.

### Qual script rodar

| Situação | Script |
| --- | --- |
| Quero **só a formatação** na planilha que já tenho | `python scripts/formatar_planilha.py` |
| Quero **acrescentar** linhas/colunas novas do painel | `python scripts/atualizar_planilha.py` (já chama o formatador ao final) |
| Quero **recriar** a planilha de exemplo do zero | `python scripts/gerar_planilha.py` — **apaga os dados reais** |

`scripts/formatar_planilha.py` existe porque os dois caminhos anteriores não
serviam a quem já tem conteúdo real: o gerador recria o arquivo do zero e o
atualizador, de propósito, herda o estilo da linha vizinha e não aplica
formatação nova. O formatador abre o `.xlsx` atual, repinta cabeçalho, listras,
cor de guia, zoom e as mensagens das listas suspensas, e salva — **sem ler nem
alterar uma célula de dado**. Célula pintada de propósito (coluna de fórmula em
cinza, destaque manual de quem preencheu) é preservada. Rode com `--conferir`
para ver o relatório sem salvar nada.

## Planilha alinhada ao painel

`scripts/esquema_planilha.py` declara, numa lista só, o que o painel espera de
cada aba: nome de coluna, largura e um valor de demonstração para quando a
célula está vazia. É a fonte de verdade que faltava — antes o gerador tinha os
cabeçalhos e o atualizador tinha uma lista própria e incompleta, então colunas
novas (Foto e Hierarquia do NUGEP, por exemplo) chegavam à planilha de exemplo
e nunca à planilha em uso.

`python scripts/atualizar_planilha.py` agora, além do que já fazia:

1. **acrescenta as colunas que faltam** em cada aba, à direita, herdando o
   estilo do cabeçalho vizinho e com a largura do esquema;
2. **preenche só as células vazias** com dado fictício — quem já tem conteúdo
   real não perde nada, e o painel abre sem buraco enquanto os dados verdadeiros
   não chegam;
3. **acrescenta as linhas que o painel espera** e talvez não existam — as duas
   chefias (Gerente-Executivo da AE e Gerente da AE/GPE) que o bloco "Contato
   institucional" precisa para mostrar a hierarquia;
4. **completa as listas suspensas** com os itens que o painel usa;
5. **formata tudo ao final** (cabeçalho v4, listras, cor de guia, zoom,
   mensagens de validação), chamando `formatar_planilha.py`.

Nada é recriado e nenhum dado é apagado. Depois de rodar, regenere o fallback
offline com `python scripts/planilha_para_js.py`.

### UI kits v3.5.0

A pasta `govbr-ds-design-uikits-v3.5.0` traz só os arquivos de design (`.fig` e
`.xd`) dos kits web, wireframe, iOS e Material, mais o material de repositório
(licença, contribuição, lint). Não há nada a incorporar: são binários de
ferramenta de design, na **v3.5** — anteriores à v4 que o painel já segue —, e o
kit wireframe é de baixa fidelidade (símbolos `.wf-*` em tons de cinza), sem
valor visual exato. Os valores que valeriam já vieram dos tokens da v4, que é a
implementação versionada.

## Planilha atualizada no arquivo (não só nos scripts)

`data/painel-processos-dados.xlsx` foi **regravado** com o esquema e a formatação
do painel. Nenhum dado foi apagado; o que havia foi lido, realinhado e devolvido.
O que mudou no arquivo:

- **NUGEP**: colunas **Foto** e **Hierarquia** criadas e preenchidas (retratos de
  exemplo, para trocar pelas fotos oficiais), e as duas chefias que o bloco
  "Contato institucional" precisa — Gerente-Executivo da AE e Gerente da AE/GPE —
  entraram como linhas, com a coluna Ordem renumerada.
- **Atividades**: a coluna antiga `Subprocesso` virou `Vinculo_Pai` (13 linhas
  migradas), que é o nome que o painel lê. A coluna `Objetivo`, que existia fora
  do esquema, foi preservada ao final.
- **Processos**: coluna `Ultima_Atualizacao` criada.
- **Duas abas novas**: **LEIA-ME**, com as regras de preenchimento e o que cada
  aba alimenta no painel, e **Listas**, com as 21 listas suspensas.
- **Formatação**: cabeçalho no azul `#0040B4` da v4 com texto branco e altura de
  30px, listra zebrada nas linhas de dados, borda fina cinza, cabeçalho
  congelado por aba, filtro automático, largura por coluna, zoom em 90% e cor de
  guia por grupo (azul para a hierarquia, verde para os anexos, azul claro para o
  conteúdo editorial, cinza para as abas de apoio).
- **Listas suspensas com mensagem**: valor fora da lista é recusado com "Escolha
  uma das opções da seta ao lado da célula", e ao entrar na célula aparece a dica
  de uso. Antes a validação era silenciosa e o erro só aparecia no painel.

Os scripts continuam valendo para as próximas rodadas (`atualizar_planilha.py`,
`formatar_planilha.py`, `esquema_planilha.py`) — a diferença é que o arquivo já
está pronto para baixar e começar a preencher.

## Revisão pelos guias de código do DS

Passagem pelos guias de **Boas práticas de HTML**, **Acessibilidade no Código
HTML (eMAG)**, **Boas práticas de CSS**, **Codificação JavaScript**, **Codificação
Sass**, **Uso do WAI-ARIA** e **Navegadores suportados**:

| Guia | Situação |
| --- | --- |
| HTML — esqueleto | `<!doctype html>`, `lang="pt-BR"`, UTF-8, viewport e `title` ✓. Faltava a meta **`X-UA-Compatible`** — incluída. |
| HTML — formatação | Tags e atributos em minúsculo, aspas duplas, atributos booleanos sem valor (`hidden`, `disabled`, `active`), `type` omitido em `script`/`link` ✓. |
| eMAG — um H1 por página | **Corrigido**: as telas internas abriam em `h2` e nenhuma tinha `h1`. Cada tela é uma página do painel e agora traz um `h1` com o próprio nome. |
| eMAG — hierarquia sem salto | **Corrigido**: os cards do NUGEP e do bloco institucional saltavam de `h2` para `h4`; passaram a `h3`. |
| eMAG — alternativa em texto | Nenhuma `<img>` sem `alt`; as decorativas usam `alt=""` + `aria-hidden` ✓. |
| eMAG — etiquetas nos campos | 53 campos, todos com `label` associado ✓ (auditado ao vivo). |
| eMAG — tabelas só para dados | Nenhuma tabela de diagramação; todas com `th`, `caption` e `scope` ✓. |
| CSS — arquitetura e comentários | Comentário de bloco por componente explicando finalidade e limite da escolha ✓. |
| CSS — mobile first | Os utilitários e o grid usam `min-width`; as exceções de tela pequena estão documentadas no ponto de uso. |
| CSS — construir pensando em temas | Sem valor hard-coded nos componentes: tudo sai de token (`--br-*` e apelidos), que é o que permite o Modo Contraste ✓. |
| JavaScript — `no-console` | Nenhum `console.log`; só `console.error` no tratamento de falha de carga ✓. |
| JavaScript — classe e método privado | O painel não usa classes ES6 (é ES5 clássico, para abrir por `file://` sem transpilação); a convenção equivalente — prefixo `_` e função por componente — é seguida. Documentado como simplificação intencional. |
| Sass | Não se aplica: o painel serve CSS direto, sem etapa de build. Os blocos seguem a mesma divisão que o guia pede em mixins (tokens → default → variações). |
| WAI-ARIA | `role` + `aria-*` conforme a taxonomia em todos os componentes; `role="none"` só onde o componente prevê, com o rótulo no `listbox` interno ✓. |
| Navegadores | Chrome 49+, Edge 14+, Firefox 67+ e Safari 11+ — o painel usa apenas recursos suportados nessa faixa. Sem Internet Explorer, como o guia recomenda. |

### Formato de data e de percentual na planilha

O LEIA-ME gravado na primeira versão contradizia o dado: mandava escrever
`DD/MM/AAAA` enquanto todas as datas estavam em ISO, e "percentual de 0 a 100"
enquanto a coluna guardava fração. Quem seguisse o texto quebrava a coluna.
Corrigido no arquivo:

- **Datas em `AAAA-MM-DD`, como texto** — é o que os dados usam e o que o painel
  lê. O LEIA-ME ganhou também o aviso de formatar a coluna como Texto antes de
  digitar: o Excel converte `29/05/2026` em número de série.
- **Percentual em fração**: 0,45 = 45%, 1 = 100%.

E `fmtData` deixou de aceitar só ISO. Agora entende as três formas que aparecem
na prática — ISO, `DD/MM/AAAA` e o número de série do Excel — para uma digitação
fora do padrão não apagar a coluna inteira da tela.

## Fichas, tags e utilitários JS

### Fichas (macroprocesso → processo → subprocesso → atividade → tarefa)

- **Um h1 por ficha (eMAG)**: as cinco fichas abriam o cabeçalho em `h2` e não
  tinham `h1` nenhum — o leitor de tela não achava o título da ficha. O elemento
  visual é o mesmo, só o nível mudou.
- **Rótulo de nível uniforme**: o "Subprocesso · Categoria" dependia de o
  macroprocesso ser alcançado pelo processo pai; num subprocesso aninhado isso
  podia falhar e o rótulo saía sem a categoria. Agora uma função só
  (`categoriaDe`) sobe a hierarquia a partir de qualquer nível, e os cinco heróis
  usam a mesma fórmula — inclusive a **Tarefa**, que não mostrava categoria.
- **Coluna "Documento"** (Documentos publicados) com largura mínima de 420px.
- Variável morta (`paiDireto`) removida da ficha do subprocesso.

### Tags no Componente Tag

As pílulas de **status do mapeamento** (Não iniciado, Em andamento, Concluído,
Suspenso) e de **nível de risco** (Baixo, Moderado, Alto, Extremo) eram
desenhadas à mão. Passaram a ser `.br-tag`, e o CSS local ficou só com a cor de
cada situação — altura, canto e espaçamento vêm do componente.

A diretriz de acessibilidade da Tag é explícita: **a cor não pode ser o único
meio de diferenciação**. Cada tag ganhou um ícone próprio (o nível de risco em
escada: check → exclamação → triângulo → radiação), e o texto vive num `<span>`
referenciado por `aria-describedby`, como o componente prevê.

### Utilitários JavaScript

| Utilitário | Situação |
| --- | --- |
| **Checkgroup** | **Corrigido**: no estado intermediário o checkbox pai leva `checked` **e** `indeterminate`. O painel marcava só `indeterminate`, então o pai nascia visualmente vazio quando havia filhos selecionados, e "selecionar todos" virava "desselecionar" no primeiro clique. |
| **Collapse** | Acionadores com `data-toggle` + `data-target`; `aria-controls`, `aria-expanded` e `hidden` gerados pelo JS ✓. Ícone chevron, corrigido na rodada anterior. |
| **Dropdown** | `data-toggle`, `data-target` e `id` no acionador; alvo com `id` ✓. Ícone caret. |
| **Accordion** | O painel não usa o utilitário `data-toggle="accordion"` — os blocos que se comportariam como accordion (Filtros, grupos do menu) são `collapse` independentes, sem exclusão mútua, que é a diferença entre os dois. O componente `br-accordion` (FAQ) é outro caminho e segue sua própria anatomia. |
| **Scrim** | `data-scrim="true"` no scrim; `role="dialog"` e `aria-modal` no diálogo interno, não no scrim. Divergência intencional: os dois no mesmo elemento anunciariam dois diálogos aninhados ao leitor de tela. Foco preso e ESC funcionam. |
| **Tooltip** | `aria-describedby`, carência antes de fechar e balão alcançável pelo cursor (WCAG 1.4.13) ✓. |

### Templates

- **Template Base**: a estrutura que o template define já é a do painel —
  `br-skiplink` com âncoras numeradas, `header[role=banner]` dentro de
  `container-lg`, navegação, `main#main[role=main]` com o conteúdo e
  `footer[role=contentinfo]`. Sem mudança.
- **Template Erro — aplicado**: a tela de "não encontrado" era um card com um
  parágrafo. Passou à anatomia do template: ilustração, mensagem em dois níveis
  (o que houve + como seguir), texto explicativo, **campo de busca** e os três
  **botões auxiliares** — Página Anterior, Página Principal e Envie um Feedback
  (que abre o relato de erro do painel). A busca leva à tela de pesquisa com o
  termo digitado.

## Revisão dos Padrões de Design, Mobile e Writing

### Corrigido

- **Empty States em tabela**: a diretriz é explícita — o estado vazio deve
  substituir **toda** a estrutura da tabela (cabeçalho, rodapé, linhas e
  colunas), senão o usuário perde tempo lendo o cabeçalho de uma tabela sem
  dados. Em Documentos publicados, a mensagem vinha numa célula com `colspan`,
  com o cabeçalho e a paginação ainda na tela. Agora a tabela inteira dá lugar
  ao estado vazio.
- **Ajuda e Comunicação — iconografia**: o botão de relato usava o ícone de
  inseto. A diretriz reserva o **balão de comentário** (`comment` / `comment-alt`)
  ao envio de feedback do usuário — trocado. O ícone de ajuda contextual do
  painel já é o `info-circle` que a diretriz indica para esse caso.
- **Densidade + Mobile — área mínima de ação**: a diretriz separa dois mínimos,
  24px para cursor e **40px para toque**. O painel dimensionava tudo para os
  24px do mouse. Em aparelho de toque (`pointer:coarse`) a área de ação passa a
  40px; onde o elemento não pode crescer sem quebrar o layout, a área
  **ultrapassa** as dimensões da superfície, que é a saída prevista na própria
  diretriz. O desktop não muda.
- **Microcopy — visualização de dados**: número em tabela alinha à direita
  (probabilidade, impacto e P×I dos riscos), com algarismos de largura fixa.
  Texto segue à esquerda e o **ID também** — é rótulo, não quantidade, como a
  diretriz distingue.
- **Writing — navegação**: os nomes acessíveis das abas passaram a usar iniciais
  maiúsculas ("Portfólio de Processos", "Repositório de Materiais"), a regra que
  a diretriz dá para componentes de navegação.

### Conferido, sem alteração

| Padrão | Situação |
| --- | --- |
| **Ajuda e Comunicação** | Ajuda global na aba **Perguntas Frequentes**, acessível de qualquer tela pela navegação (a diretriz indica cabeçalho, rodapé ou menu); ajuda contextual em Tooltip com ícone "i", como o padrão pede para texto curto. |
| **Content Overflow** | Rolagem vertical como norma; a horizontal só nas tabelas largas, com sombra de rolagem. Truncamento sempre com acesso ao texto completo, e nunca em título, rótulo de botão ou mensagem. Link externo sinalizado com `external-link-alt` e aviso "(abre em nova aba)" para leitor de tela. |
| **Dropdown** | Acionador + identificador `caret` + superfície flutuante com elevação a partir da camada 2; inverte a direção junto às bordas; fecha pelo acionador, por clique fora e por ESC. |
| **Collapse** | Identificador `chevron` apontando para a direção da expansão; conteúdo empurra os elementos adjacentes; múltiplos elementos abrem ao mesmo tempo (o painel é de consulta comparativa). |
| **Formulário** | Rótulos em sentença acima do campo, sempre visíveis; par de botões com a ênfase primária à direita; sem "Redefinir"/"Limpar Formulário" em entrada de dados. |
| **Gráfico** | Anatomia com cabeçalho (título + legenda), área principal e `role="img"` com descrição; matriz de risco com legenda direta e tags que não dependem de cor. |
| **Onboarding** | Carousel + overlay + Step como indicador de progresso, saída disponível em todas as páginas, tela de boas-vindas e tela de conclusão. |
| **Navegação** | Lateral pelas abas (sempre visíveis), progressiva pelos cards e pelo Wizard, reversa por breadcrumb — os três tipos que a diretriz define. |
| **Mobile** | O painel é **web responsivo**: pela tabela comparativa da diretriz, todos os dez fundamentos seguem o DS (a coluna "Nativo" só vale para Android/iOS nativos). Nada a substituir. |
| **Princípios de UX Writing** | Voz neutra e direta, sem gíria nem jargão; hiperlinks descritivos (sem "clique aqui"); classificação por cor sempre acompanhada de texto e ícone. |

## Medidas do Componente Tag conferidas contra o código oficial

A pasta `govbr-ds-core-v3.7.0` trouxe o **SCSS de origem** dos componentes —
era o que faltava para conferir medidas em vez de estimá-las. Começando pela
Tag (`src/components/tag/_mixins.scss`, sobre a escala de espaçamento de base
8px), as nossas estavam abaixo do padrão em todos os tipos:

| Tipo | Tinha | Oficial |
| --- | --- | --- |
| Texto — small / medium / large | 16 · 24 · 32 | **20 · 28 · 36** (2xh · 3xh · 4xh) |
| Texto — canto | pílula | **4px** (`surface-rounder-sm`) |
| Texto — fonte | 12,44px | **14px** (`font-size-scale-base`) |
| Status | 8 · 12 · 16 | **12 · 16 · 24** (baseh · 2x · 3x) |
| Contagem | 16 · 20 · 24 | **20 · 24 · 28** (2xh · 3x · 3xh) |
| Ícone | — | **28 · 32 · 44** (3xh · 4x · 5xh), círculo |
| Interação | 16 · 24 · 32 | **32 · 40 · 44** (4x · 5x · 5xh) |

O **canto** é o achado que mais muda a aparência: a Tag de texto do DS é um
retângulo de canto pequeno, não uma pílula. A pílula pertence só ao tipo
contagem, o único que o SCSS declara com `border-radius:100em`. Entraram também
o espaçamento mínimo de 4px entre tags vizinhas e a borda branca de 1px que
status e contagem levam para se destacarem sobre qualquer fundo.

Como as tags de status do mapeamento e de nível de risco herdam do componente,
elas acompanham a correção.

## Ilustrações oficiais conferidas cena a cena

O pacote completo de ilustrações (46 empty-space, 24 de erro, 8 personagens ×
16 poses, objetos e fundos) permitiu conferir se cada tela mostra a **cena**
certa — não bastava usar arquivos do pacote, era preciso usar o arquivo que
conta o que aconteceu.

Estados vazios, antes e depois:

| Tela | Tinha | Agora |
| --- | --- | --- |
| Portfólio, filtros sem retorno | 00 (mensagem enviada) | **07** — lupa sobre documento, com "?" |
| Busca livre sem retorno | 00 (mensagem enviada) | **44** — a mesma cena em fundo orgânico |
| Documentos publicados | 05 (segurança de dados) | **03** — maleta cheia de papéis |
| Repositório de materiais | 10 (relógio) | **04** — tela com pastas e arquivos |
| NUGEP sem integrantes | 20 (engrenagem) | **14** — pessoa saindo de um dispositivo |
| Glossário | 05 (segurança de dados) | **07** — pesquisa sem retorno |

A busca livre e o filtro do portfólio usam variações da **mesma** cena (07 e
44): são o mesmo tipo de vazio, e a variação evita repetir a imagem quando o
usuário passa de um para o outro.

No **Template Erro** a escolha se confirmou: `erro/error01.png` é a cena do 404,
exatamente o caso de código inexistente na URL.

No **tutorial**, duas páginas prometiam no texto alternativo uma cena que o
arquivo não mostrava — a primeira dizia "foguete decolando" exibindo a cena de
mensagem enviada, e a quarta dizia "medidor de desempenho" exibindo um relógio.
A abertura passou a mostrar a **personagem acenando** (Lina, do pacote oficial),
que é onde a saudação faz sentido; a página de indicadores mostra a tela com
painéis; e a de encerramento, a maleta de trabalho.

As demais pastas do pacote foram avaliadas e ficaram de fora: **objetos** e
**fundos geométricos** são peças de composição para arte editorial, não para
estado de interface, e o painel não tem superfície que peça isso. Os
**personagens** entram apenas na saudação do tutorial — o fundamento Ilustração
reserva figura humana a boas-vindas, erro e vazio, e usá-los em mais pontos
daria ao painel um tom de campanha que ele não tem.

## Seleção de linhas e exportação em CSV em todas as tabelas

A tabela de Documentos publicados tinha seleção de linhas e exportação; as
demais, não. O Componente Table define isso como o **comportamento 4** da sua
anatomia, então passou a valer para todas.

A coluna de seleção é **injetada em JS**, não escrita em cada tabela: as linhas
são montadas por dezenas de trechos diferentes em `app.js`, e acrescentar uma
célula em cada um multiplicaria o mesmo código por dezenas de lugares. O CSV
também sai do DOM — o texto visível de cada célula —, de modo que qualquer
tabela nova ganha a exportação sem escrever uma linha.

O que entrou, seguindo a anatomia:

- **Caixa de seleção** no início de cada linha e **"selecionar tudo"** no header,
  com o estado intermediário do utilitário Checkgroup (`checked` + `indeterminate`)
  quando parte das linhas está marcada;
- **Barra Contextual** surgindo sob a Barra de Título com a contagem ("1 item
  selecionado" / "N itens selecionados") e as ações da seleção — exportar em CSV
  e limpar;
- **Estado selecionado** na linha (`is-selected`);
- O CSV sai com BOM e ponto e vírgula, para abrir direto no Excel em pt-BR, e o
  nome do arquivo vem do título da tabela.

A linha de conteúdo expandido tem o `colspan` ajustado junto, para não
desalinhar com a coluna nova.

## Menu gov.br e textos editáveis pela planilha

### Submenus do menu gov.br

Os painéis de submenu (Serviços → Buscar serviços por → Categorias…) se
sobrepunham ao texto do nível anterior a cada abertura. O submenu do menu gov.br
é um painel que **desliza por cima** do nível que o abriu — é o que dá sentido ao
botão "Voltar" que cada um traz. Aqui ele estava expandindo em linha, como filho
flex do próprio `<li>`: cada nível empurrava o conteúdo e, do segundo em diante,
os painéis se atropelavam. Agora cada `.off` cobre o painel que o abriu (o de
nível 1 cobre a lista base, o de nível 2 cobre o de nível 1) com rolagem própria.

### Textos institucionais na planilha

A aba **Parâmetros** ganhou três chaves, para o texto mudar sem tocar no código:

| Chave | O que controla |
| --- | --- |
| `Titulo_Inicio` | Título da tela inicial |
| `Subtitulo_Inicio` | Linha de apoio sob o título (fica oculta se vazia) |
| `Titulo_Repositorio` | Título da aba Repositório de materiais |

Somadas às que já existiam — `Contato_Unidade`, `Contato_Email`,
`Contato_Telefone`, `Link_Metodologia` e `Link_Guia` —, são oito pontos de texto
institucional editáveis pela planilha.

A leitura passa por uma função `par(chave, padrao)` que **cai no texto atual**
quando a chave não existe: planilha antiga continua funcionando e o painel nunca
abre com buraco. Para acrescentar outro texto editável, basta uma linha nova na
aba Parâmetros e uma chamada a `par()` no ponto correspondente.

As três linhas foram gravadas **direto no .xlsx**, com a listra zebrada da aba, e
entraram também em `js/dados.js` (fallback offline) e em
`scripts/dados_conteudo.py` (gerador). Nenhum dado existente foi alterado.

## Filtros, ordenação e expansão de linha

### Filtro que não voltava para "todos"

O bug tinha duas faces. No **Select de seleção única** o item é um radio:
clicar de novo no já escolhido não dispara mudança, e apagar o texto do campo
só filtra a lista — ao fechar, o rótulo era reescrito. Não havia como voltar a
"Todos os tipos". Agora o item já escolhido desmarca no segundo clique, e o
campo ganha um **botão de limpar** que aparece enquanto há escolha.

A outra face era de tela: o recorte continuava valendo depois de sair da aba,
sem nada que o explicasse. **Documentos, Repositório e Glossário** ganharam o
mesmo rodapé de filtros ativos que o Portfólio já tinha — tags com o que está
aplicado, cada uma removível, e o botão **Limpar Filtros**.

### Ordenação em três estados

O componente prevê três (seta dupla = ordenação padrão, crescente,
decrescente), mas o terceiro clique voltava a "crescente" e não havia como
desfazer. Agora o ciclo é **crescente → decrescente → ordem original**, que é a
ordem em que a tela montou a tabela, guardada na primeira ordenação.

### Expandir e retrair linha em todas as tabelas

O comportamento 3 da anatomia do Table, que só existia em Documentos
publicados, passou a valer para todas. A coluna do chevron é **injetada em JS**,
pelo mesmo motivo da coluna de seleção: as linhas são montadas por dezenas de
trechos em `app.js`. Quem quiser o comportamento só emite, depois da linha, um
`<tr class="collapse">` — a coluna, o `colspan`, o `aria-controls` e o chevron saem
daqui. O conteúdo expandido leva sombra interna, que é como o componente marca
a hierarquia entre o detalhe e a linha que o abriu.

A informação secundária saiu das colunas e foi para o detalhe. Uma informação
aparece **ou** na coluna **ou** no detalhe, nunca nas duas:

| Tabela | Ficou na tabela | Foi para o detalhe |
| --- | --- | --- |
| Atividades | Código · Atividade · Responsável · Prazo | Entradas, saídas, sistemas, executor, base normativa, corresponsáveis |
| Tarefas | Código · Tarefa · Tipo · Duração | Descrição, responsável, sistema, observações |
| Riscos | ID · Risco · P · I · P×I · Nível · Status | Tratamento, resposta, categoria, responsável |
| Indicadores | ID · Indicador · Meta · Resultado · Situação · Última medição | Fórmula, periodicidade, polaridade, fonte |
| Subprocessos | Código · Subprocesso · Ação | Entradas, saídas, sistemas, unidade responsável |

A **tabela de atividades** passou ao formato da de tarefas (código, nome,
classificação e tempo), que era o pedido: mesma disposição, dados próprios de
cada nível.

### Outros

- **Tique do checkbox no Select** ficava no canto superior esquerdo do quadrado:
  a caixa começa em 14px/8px e o tique estava em 19/9, quando o deslocamento
  correto dentro dela é 9/4 — daí 23/12.
- **NUGEP**: e-mail institucional e nome de unidade são cadeias longas sem
  espaço, que em tela estreita passavam da borda do cartão. O `min-width:0`
  desarma o mínimo automático do item flex e a quebra forçada vale para as
  cadeias sem espaço.

## Planilha: o que se calcula sozinho

Revisão de campo repetido — informação que o usuário digitava duas vezes, ou
que o painel já derivava do que estava ao lado. Sete colunas passaram a ser
**fórmula**, no cinza que a planilha já usava para "não digite aqui":

| Aba · coluna | De onde vem |
| --- | --- |
| Processos · `Macroprocesso` | O código do processo já carrega o do macroprocesso (P-06.01 → MP-06) |
| Processos · `Percentual` | Marcos "Sim" ÷ marcos aplicáveis; os "Não se aplica" saem da conta |
| Documentos, Riscos, Indicadores · `Vinculo_Nivel` | O prefixo do código em `Vinculo_Codigo` (MP-, P-, SP-, A-, T-) |
| Riscos · `Nivel_PxI` | Probabilidade × Impacto |
| Riscos · `Classificacao` | Faixa do P×I: 20+ Extremo · 12+ Alto · 5+ Moderado · abaixo, Baixo |
| Indicadores · `Situacao` | Resultado contra Meta, conforme a Polaridade |
| NUGEP · `Unidade_Nome` | A sigla, pela tabela de unidades na aba Listas |

As três últimas o painel **já recalculava** ao carregar (`classeRisco` e
`situacaoInd` em `app.js`) — o que estava digitado era ignorado. As fórmulas usam
exatamente as mesmas faixas e a mesma regra, então planilha e tela dizem a
mesma coisa.

A aba **Listas** ganhou a tabela de referência de unidades (colunas V e W):
a sigla puxa o nome por extenso, que antes era redigitado a cada pessoa do
NUGEP.

O **LEIA-ME** passou a ter um bloco próprio sobre as colunas cinza — o que cada
uma calcula e por que não se digita nelas — e a nota sobre vínculo múltiplo: a
fórmula resolve o vínculo único; para apontar para mais de um item, os códigos
vão separados por ponto e vírgula e os níveis à mão, na mesma ordem.
