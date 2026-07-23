# -*- coding: utf-8 -*-
"""
Monta o index.html do Painel de Gestão de Processos reaproveitando,
byte a byte, o cabeçalho, o menu, a seção de equipe/atendimento, o
rodapé e o VLibras do Painel de Transformação Digital (index original),
e injetando o <main>, as abas e os modais próprios deste painel.

Também extrai o <style> completo do original para css/govbr-ds.css.
"""
import os, re, sys

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
ORIG = "/mnt/user-data/uploads/index.html"
s = open(ORIG, encoding="utf-8").read()


def entre(texto, ini, fim, incluir_fim=False):
    a = texto.find(ini)
    if a < 0:
        sys.exit(f"Marcador inicial não encontrado: {ini[:60]!r}")
    b = texto.find(fim, a)
    if b < 0:
        sys.exit(f"Marcador final não encontrado: {fim[:60]!r}")
    return texto[a:b + (len(fim) if incluir_fim else 0)]


# ── 1. CSS do DS (conteúdo do <style>) ─────────────────────────────────
css = entre(s, "  <style>", "</style>")[len("  <style>"):]
os.makedirs(os.path.join(BASE, "css"), exist_ok=True)
open(os.path.join(BASE, "css", "govbr-ds.css"), "w", encoding="utf-8").write(
    "/* Extraído automaticamente do Painel de Transformação Digital (index original)\n"
    "   — gov.br DS v4 reimplementado. Não edite: ajustes vão em painel.css. */\n" + css)

# ── 2. Fatias do corpo ─────────────────────────────────────────────────
skiplink = entre(s, '<nav class="br-skiplink"', "</nav>", True)
header = entre(s, "<!-- ══════════════════════════════════════════════\n     br-header",
               "</header>", True)
headerbot = entre(s, "<!-- ── header-bottom-wrap", "<!-- br-tab")
alertas = re.search(r'^\s*<div class="tab-bar-actions">.*$', s, re.M).group(0).strip()
menu = entre(s, '<div class="br-menu" id="sectionMenu">',
             "<!-- ══════════════════════════════════════════════\n     MAGIC BUTTON")
team = entre(s, '<div class="page-cnt team-section">', '  <footer class="br-footer"')
footer = entre(s, '<footer class="br-footer"', "</footer>", True)
vlibras = entre(s, "<!-- VLibras Widget", "</body>")

# ── 3. Ajustes de conteúdo (PTD → Processos) ───────────────────────────
header = header.replace(
    "Painel de Acompanhamento do Plano de Transformação Digital (PTD).",
    "Painel de Gestão de Processos — cadeia de valor, mapeamento e repositório de processos.")
# os dois PDFs do PTD viram Planilha de dados + Metodologia
pdfs = re.findall(r'<a class="br-item" href="https://www\.gov\.br/governodigital/pt-br/estrategias-e-governanca-digital/planos-de-transformacao-digital[^>]*>.*?</a>', header, re.S)
if len(pdfs) >= 2:
    header = header.replace(pdfs[0],
        '<a class="br-item" href="data/painel-processos-dados.xlsx" download><i class="fas fa-file-excel" aria-hidden="true"></i> Planilha de dados (modelo)</a>')
    header = header.replace(pdfs[1],
        '<a class="br-item" href="#/metodologia"><i class="fas fa-book" aria-hidden="true"></i> Metodologia (CBOK · PMBOK)</a>')
headerbot = (headerbot
    .replace("Painel de Acompanhamento do Plano de Transformação Digital (PTD).",
             "Painel de Gestão de Processos — cadeia de valor, mapeamento e repositório de processos.")
    .replace("Pesquisar termos visíveis nesta página", "Pesquisar processos, documentos e registros")
    .replace("Pesquisar nesta página…", "Pesquisar no painel…"))

# equipe: remove o 1º card (Ponto Focal do PTD) e ajusta textos
cards = [m.start() for m in re.finditer(r'<div class="br-card team-card">', team)]
if len(cards) >= 2:
    team = team[:cards[0]] + team[cards[1]:]
team = (team
    .replace("Conheça a equipe responsável pelo Plano de Transformação Digital e por este painel na Codevasf.",
             "Conheça a equipe responsável pela gestão de processos (mapeamento, repositório e melhoria) e por este painel na Codevasf.")
    .replace("Analista em Desenvolvimento Regional — Suporte Técnico do PTD",
             "Analista em Desenvolvimento Regional — Gestão de Processos"))

footer = footer.replace("Processo e-Codevasf nº 59500.003375/2025-08-e",
                        "Unidade de Gestão Normativa e de Processos — AE/GPE/UNP")

# gatilhos de ajuda/acessibilidade passam a abrir os modais deste painel
for bloco in ("header", "menu", "footer"):
    pass
def liga_modais(txt):
    txt = re.sub(r'id="(menuHelpTrigger|footerHelpTrigger)"',
                 r'id="\1" data-modal-open="modalAjuda"', txt)
    txt = re.sub(r'id="(a11yStatementTrigger|menuA11yStatementTrigger)"',
                 r'id="\1" data-modal-open="modalA11y"', txt)
    return txt
header, menu, footer = liga_modais(header), liga_modais(menu), liga_modais(footer)

# ── 4. Fragmentos próprios ─────────────────────────────────────────────
HEAD = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Painel de Gestão de Processos · Codevasf · gov.br</title>
  <meta name="description" content="Painel de gestão de processos da Codevasf: cadeia de valor, mapeamento (BPMN/Bizagi), documentos, riscos, indicadores e diário de mapeamento — CBOK 4.0 e PMBOK.">
  <meta name="keywords" content="gestão de processos, BPM, cadeia de valor, mapeamento de processos, codevasf, gov.br, CBOK, PMBOK, bizagi">
  <meta name="robots" content="noindex, nofollow">
  <meta name="author" content="Codevasf — Ministério da Integração e do Desenvolvimento Regional">
  <link rel="icon" href="https://gov.br/favicon.ico">
  <link rel="icon" type="image/png" sizes="32x32" href="https://gov.br/++theme++padrao_govbr/favicons/favicon-32x32.png">
  <link rel="icon" type="image/png" sizes="48x48" href="https://gov.br/++theme++padrao_govbr/favicons/favicon-48x48.png">
  <link rel="apple-touch-icon" href="https://gov.br/++theme++padrao_govbr/favicons/apple-touch-icon.png">
  <meta name="theme-color" content="#0040b4">
  <meta property="og:title" content="Painel de Gestão de Processos · Codevasf">
  <meta property="og:description" content="Cadeia de valor, mapeamento de processos, documentos, riscos, indicadores e diário de mapeamento da Codevasf.">
  <meta property="og:type" content="website">
  <meta property="og:locale" content="pt_BR">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&family=Noto+Sans+Mono:wght@100..900&display=swap">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" crossorigin="anonymous" referrerpolicy="no-referrer">
  <link rel="stylesheet" href="css/govbr-ds.css">
  <link rel="stylesheet" href="css/painel.css">
  <script>
    /* ── Configuração do painel ──────────────────────────────────────
       Para alimentar via Google Sheets (recomendado, igual ao painel
       do PTD): importe data/painel-processos-dados.xlsx no Google
       Sheets, compartilhe como "qualquer pessoa com o link pode ver"
       e cole o ID da planilha abaixo (o trecho entre /d/ e /edit). */
    window.PAINEL_CONFIG = {
      googleSheetId: '',
      arquivoXlsx: 'data/painel-processos-dados.xlsx'
    };
  </script>
  <script defer src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
  <script defer src="js/dados.js"></script>
  <script defer src="js/govbr-ui.js"></script>
  <script defer src="js/app.js"></script>
</head>
<body>
"""

def aba(painel, rota, icone, cheio, curto, cnt=None, ativa=False):
    count = f' <span class="tab-count" id="{cnt}">—</span>' if cnt else ""
    return (f'<li class="tab-item{" active" if ativa else ""}">'
            f'<button type="button" id="tab-{painel}" role="tab" aria-selected="{"true" if ativa else "false"}" '
            f'aria-controls="panel-{painel}" data-rota="{rota}" data-painel="{painel}">'
            f'<span class="name"><i class="fas {icone}" aria-hidden="true"></i> '
            f'<span class="tab-label"><span class="tab-label-full">{cheio}</span>'
            f'<span class="tab-label-short">{curto}</span></span>{count}</span></button></li>')

TABS = ('''
<!-- br-tab — faixa de navegação principal do Painel de Gestão de Processos
     (mesma anatomia da faixa do painel do PTD; os Alertas operacionais
     agora listam prazos de mapeamento vencidos e riscos Alto/Extremo). -->
<div class="br-tab" id="navigation">
  <div class="tab-bar-row">
    <nav class="tab-nav" aria-label="Seções do painel">
      <ul role="tablist">'''
    + aba("inicio", "#/", "fa-house", "Início · Cadeia de Valor", "Início", ativa=True)
    + aba("catalogo", "#/catalogo", "fa-layer-group", "Catálogo de Processos", "Processos", "cntCatalogo")
    + aba("documentos", "#/documentos", "fa-folder-open", "Documentos", "Docs", "cntDocumentos")
    + aba("riscos", "#/riscos", "fa-shield-halved", "Radar de Riscos", "Riscos", "cntRiscos")
    + aba("indicadores", "#/indicadores", "fa-chart-line", "Indicadores", "Indic.", "cntIndicadores")
    + aba("diario", "#/diario", "fa-timeline", "Diário de Mapeamento", "Diário", "cntDiario")
    + aba("metodologia", "#/metodologia", "fa-book", "Metodologia", "Método")
    + '''</ul>
    </nav>
    ''' + alertas + '''
  </div>
</div>
''')

MAIN = """
<main id="main" role="main">
  <div id="mainTabContent" class="tab-content">
    <section class="tab-panel active pp-cnt" id="panel-inicio" role="tabpanel" aria-labelledby="tab-inicio"><div id="viewInicio"></div></section>
    <section class="tab-panel pp-cnt" id="panel-catalogo" role="tabpanel" aria-labelledby="tab-catalogo" hidden><div id="viewCatalogo"></div></section>
    <section class="tab-panel pp-cnt" id="panel-detalhe" role="tabpanel" aria-label="Ficha do item selecionado" hidden><div id="viewDetalhe"></div></section>
    <section class="tab-panel pp-cnt" id="panel-documentos" role="tabpanel" aria-labelledby="tab-documentos" hidden><div id="viewDocumentos"></div></section>
    <section class="tab-panel pp-cnt" id="panel-riscos" role="tabpanel" aria-labelledby="tab-riscos" hidden><div id="viewRiscos"></div></section>
    <section class="tab-panel pp-cnt" id="panel-indicadores" role="tabpanel" aria-labelledby="tab-indicadores" hidden><div id="viewIndicadores"></div></section>
    <section class="tab-panel pp-cnt" id="panel-diario" role="tabpanel" aria-labelledby="tab-diario" hidden><div id="viewDiario"></div></section>
    <section class="tab-panel pp-cnt" id="panel-metodologia" role="tabpanel" aria-labelledby="tab-metodologia" hidden><div id="viewMetodologia"></div></section>
    <section class="tab-panel pp-cnt" id="panel-busca" role="tabpanel" aria-label="Resultados da busca" hidden><div id="viewBusca"></div></section>
  </div>
</main>
"""

MAGIC = """
<!-- MAGIC BUTTON — único CTA do painel: canal direto com a UNP para
     sugerir melhorias de processos (substitui o CTA do PTD). -->
<div class="magic-button-area">
  <div class="br-magic-button">
    <a class="br-button" id="magicButtonProcessos" href="mailto:ae.gpe.unp@codevasf.gov.br?subject=Sugest%C3%A3o%20de%20melhoria%20de%20processo">
      Sugerir melhoria de processo
    </a>
  </div>
</div>
"""

MODAIS = """
<!-- ─── Modais do painel (ajuda, acessibilidade, cookies) ─── -->
<div class="br-scrim foco" id="modalAjuda" hidden>
  <div class="br-modal medium" role="dialog" aria-modal="true" aria-labelledby="modalAjudaTitulo">
    <div class="br-modal-header"><h1 class="modal-title" id="modalAjudaTitulo">Como usar este painel</h1>
      <button class="br-button close circle small" type="button" data-dismiss="true" aria-label="Fechar"><i class="fas fa-times" aria-hidden="true"></i></button></div>
    <div class="br-modal-body">
      <p><strong>Navegação hierárquica:</strong> na aba Início, clique em um macroprocesso da Cadeia de Valor para abrir a ficha; de lá, avance para processos, subprocessos e atividades. A trilha (breadcrumb) no topo de cada ficha mostra o caminho completo e permite voltar a qualquer nível.</p>
      <p><strong>Fichas:</strong> cada nível traz objetivo, responsáveis, SIPOC, diagrama BPMN (Bizagi), documentos, riscos (P×I), indicadores e — nos processos — os marcos M1–M9 e o diário de mapeamento.</p>
      <p><strong>Busca:</strong> use a lupa do cabeçalho para pesquisar por código, nome, documento ou registro.</p>
      <p><strong>Alertas:</strong> o sino da faixa de navegação lista prazos de mapeamento vencidos e riscos Alto/Extremo em aberto.</p>
      <p><strong>Dados:</strong> tudo vem da planilha (Google Sheets ou <code>data/painel-processos-dados.xlsx</code>). Edite a planilha e o painel reflete — detalhes na aba Metodologia.</p>
    </div>
  </div>
</div>
<div class="br-scrim foco" id="modalA11y" hidden>
  <div class="br-modal medium" role="dialog" aria-modal="true" aria-labelledby="modalA11yTitulo">
    <div class="br-modal-header"><h1 class="modal-title" id="modalA11yTitulo">Declaração de Acessibilidade</h1>
      <button class="br-button close circle small" type="button" data-dismiss="true" aria-label="Fechar"><i class="fas fa-times" aria-hidden="true"></i></button></div>
    <div class="br-modal-body">
      <p>Este painel segue o Design System do governo federal (gov.br DS v4) e as diretrizes do eMAG/WCAG: navegação por teclado com anel de foco visível, atalhos de acesso rápido (teclas 1–4), modo de alto contraste, textos alternativos nas imagens, VLibras e estrutura semântica com marcos ARIA.</p>
      <p>Encontrou uma barreira de acessibilidade? Escreva para <a href="mailto:ae.gpe.unp@codevasf.gov.br">ae.gpe.unp@codevasf.gov.br</a>.</p>
    </div>
  </div>
</div>
<div class="br-scrim foco" id="modalCookies" hidden>
  <div class="br-modal small" role="dialog" aria-modal="true" aria-labelledby="modalCookiesTitulo">
    <div class="br-modal-header"><h1 class="modal-title" id="modalCookiesTitulo">Cookies e privacidade</h1>
      <button class="br-button close circle small" type="button" data-dismiss="true" aria-label="Fechar"><i class="fas fa-times" aria-hidden="true"></i></button></div>
    <div class="br-modal-body">
      <p>Este painel não utiliza cookies de rastreamento. Apenas preferências locais (como o modo de alto contraste) são guardadas no seu navegador, e você pode limpá-las a qualquer momento nas configurações do próprio navegador.</p>
    </div>
  </div>
</div>
"""

# ── 5. Montagem final ──────────────────────────────────────────────────
partes = [HEAD, "\n", skiplink, "\n\n", header, "\n", headerbot, "\n", TABS,
          MAIN, "\n", menu, MAGIC, "\n", team, "\n", footer, "\n", MODAIS,
          "\n", vlibras, "</body>\n</html>\n"]
html = "".join(partes)
open(os.path.join(BASE, "index.html"), "w", encoding="utf-8").write(html)
print(f"index.html montado: {len(html):,} bytes | css/govbr-ds.css: {len(css):,} bytes")
for chave in ["Painel de Gestão de Processos", "viewInicio", "notifPanel",
              "sectionMenu", "br-footer", "vlibras-plugin", "modalAjuda"]:
    print(("ok  " if chave in html else "FALTA "), chave)
