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

`M1` Reunião de contextualização · `M2` Macroprocesso e processo modelados ·
`M3` Subprocessos modelados · `M4` AS-IS modelado · `M5` AS-IS validado ·
`M6` Procedimento validado · `M7` Procedimento aprovado · `M8` TO-BE
elaborado · `M9` TO-BE validado · `M10` Publicado no repositório.

M2 e M3 representam as diversas oficinas de modelagem entre a contextualização
e o AS-IS consolidado — desenhando macroprocesso, processo e subprocessos
(inclusive descobrindo subprocessos ainda não mapeados). M6/M7 tratam do POP
(Procedimento Operacional Padrão, CBOK 4.0): primeiro validado tecnicamente
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
  AS-IS, TO-BE, KPI, POP).

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
