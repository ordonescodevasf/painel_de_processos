# -*- coding: utf-8 -*-
"""Gera SVGs ilustrativos (estilo BPMN) para img/diagramas/.
Substitua cada arquivo pela exportação real do Bizagi (PNG/SVG) mantendo o nome."""
import os, html

ITENS = {
    "mp-01": "Gestão Estratégica e Governança",
    "mp-02": "Gestão de Riscos, Integridade e Controles",
    "mp-03": "Desenvolvimento Territorial e Estruturação Produtiva",
    "mp-04": "Gestão de Empreendimentos de Irrigação",
    "mp-05": "Revitalização de Bacias Hidrográficas",
    "mp-06": "Gestão de Licitações e Contratos",
    "mp-07": "Gestão de Pessoas",
    "mp-08": "Gestão de Tecnologia da Informação",
    "p-06-01": "Planejamento da Contratação",
    "p-06-02": "Seleção do Fornecedor",
    "p-06-03": "Gestão e Fiscalização Contratual",
    "p-04-01": "Operação e Manutenção de Perímetros Irrigados",
    "p-05-01": "Recuperação de Nascentes e Matas Ciliares",
    "p-01-01": "Formulação e Monitoramento do Planejamento Estratégico",
    "p-07-01": "Admissão e Integração de Empregados",
    "sp-06-01-01": "Estudo Técnico Preliminar (ETP)",
    "sp-06-01-02": "Termo de Referência (TR)",
    "sp-06-01-03": "Pesquisa de Preços",
    "sp-06-02-01": "Condução da Sessão Pública",
    "sp-04-01-01": "Distribuição de Água aos Irrigantes",
}

TPL = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 420" font-family="'Noto Sans',Arial,sans-serif">
  <rect width="960" height="420" fill="#ffffff"/>
  <rect x="8" y="8" width="944" height="404" fill="none" stroke="#1351b4" stroke-width="2" rx="6"/>
  <rect x="8" y="8" width="36" height="404" fill="#1351b4" rx="6"/>
  <text x="26" y="214" fill="#fff" font-size="15" font-weight="700" text-anchor="middle" transform="rotate(-90 26 214)">{codigo}</text>
  <text x="60" y="42" fill="#0c326f" font-size="20" font-weight="700">{titulo}</text>
  <text x="60" y="64" fill="#555" font-size="12">Diagrama BPMN ilustrativo — substitua este arquivo pela exportação do Bizagi Modeler (mesmo nome).</text>
  <line x1="44" y1="140" x2="952" y2="140" stroke="#ccc" stroke-dasharray="6 4"/>
  <line x1="44" y1="280" x2="952" y2="280" stroke="#ccc" stroke-dasharray="6 4"/>
  <text x="60" y="128" fill="#888" font-size="11" font-weight="700">RAIA — ÁREA DEMANDANTE</text>
  <text x="60" y="268" fill="#888" font-size="11" font-weight="700">RAIA — UNIDADE EXECUTORA</text>
  <circle cx="110" cy="205" r="16" fill="#fff" stroke="#168821" stroke-width="3"/>
  <rect x="170" y="180" width="150" height="50" rx="8" fill="#e6f0fa" stroke="#1351b4" stroke-width="2"/>
  <text x="245" y="209" fill="#0c326f" font-size="12" font-weight="600" text-anchor="middle">Receber demanda</text>
  <rect x="370" y="180" width="150" height="50" rx="8" fill="#e6f0fa" stroke="#1351b4" stroke-width="2"/>
  <text x="445" y="203" fill="#0c326f" font-size="12" font-weight="600" text-anchor="middle">Analisar e</text>
  <text x="445" y="218" fill="#0c326f" font-size="12" font-weight="600" text-anchor="middle">instruir</text>
  <path d="M600 205 l28 -24 28 24 -28 24 z" fill="#fff8e1" stroke="#b38600" stroke-width="2"/>
  <rect x="700" y="320" width="150" height="50" rx="8" fill="#e6f0fa" stroke="#1351b4" stroke-width="2"/>
  <text x="775" y="349" fill="#0c326f" font-size="12" font-weight="600" text-anchor="middle">Ajustar</text>
  <rect x="700" y="180" width="150" height="50" rx="8" fill="#e6f0fa" stroke="#1351b4" stroke-width="2"/>
  <text x="775" y="209" fill="#0c326f" font-size="12" font-weight="600" text-anchor="middle">Aprovar</text>
  <circle cx="910" cy="205" r="16" fill="#fff" stroke="#e52207" stroke-width="4"/>
  <g stroke="#333" stroke-width="2" fill="none" marker-end="url(#seta)">
    <line x1="126" y1="205" x2="168" y2="205"/>
    <line x1="320" y1="205" x2="368" y2="205"/>
    <line x1="520" y1="205" x2="598" y2="205"/>
    <line x1="656" y1="205" x2="698" y2="205"/>
    <path d="M628 229 v116 h70"/>
    <line x1="850" y1="205" x2="892" y2="205"/>
  </g>
  <defs><marker id="seta" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0 0 L8 4 L0 8 z" fill="#333"/></marker></defs>
  <text x="612" y="176" fill="#b38600" font-size="10" text-anchor="middle">Conforme?</text>
</svg>
"""

out = os.path.join(os.path.dirname(__file__), "..", "img", "diagramas")
os.makedirs(out, exist_ok=True)
for cod, nome in ITENS.items():
    svg = TPL.format(codigo=html.escape(cod.upper()), titulo=html.escape(nome))
    with open(os.path.join(out, f"{cod}.svg"), "w", encoding="utf-8") as f:
        f.write(svg)
print(f"{len(ITENS)} diagramas gerados em img/diagramas/")
