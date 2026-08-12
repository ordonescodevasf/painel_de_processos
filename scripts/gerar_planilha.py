# -*- coding: utf-8 -*-
"""
Gera a planilha data/painel-processos-dados.xlsx com DADOS FICTÍCIOS,
fonte de dados do Painel de Processos (GitHub Pages).

Estrutura (alinhada ao BPM CBOK 4.0 e ao PMBOK):
  LEIA-ME | Macroprocessos | Processos | Subprocessos | Atividades |
  Documentos | Riscos | Indicadores | Listas

ATENÇÃO — este script RECRIA a planilha de exemplo do zero (dados fictícios).
Se a planilha em data/ já tem conteúdo real ou colunas acrescentadas depois
(por exemplo "Objetivo" em Atividades), NÃO rode este script: use
scripts/atualizar_planilha.py, que altera a planilha existente no lugar,
preservando dados, colunas extras, formatação e validações.
"""
import datetime as dt
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dados_conteudo as CONTEUDO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

AZUL = "0040B4"      # azul gov.br v4 (blue-warm-vivid-70)
AZUL_ESCURO = "0C326F"
CINZA_FORMULA = "F0F0F0"
CINZA_ZEBRA = "F5F8FC"   # listra de leitura (blue-warm-vivid-5 clareado)
# Arial, e não Noto Sans (a família da v4): a planilha é preenchida no
# Excel/LibreOffice de cada pessoa, onde a Noto Sans não está instalada —
# uma fonte ausente vira substituição imprevisível. Arial mantém a métrica
# estável em qualquer máquina.
FONTE = "Arial"

# Cor da guia por grupo de conteúdo — a planilha tem 16 abas e a barra
# inferior é a única navegação que o usuário tem.
COR_GUIA = {
    "LEIA-ME": AZUL_ESCURO,
    "Macroprocessos": AZUL, "Processos": AZUL, "Subprocessos": AZUL,
    "Atividades": AZUL, "Tarefas": AZUL,
    "Documentos": "168821", "Riscos": "168821", "Indicadores": "168821",
    "Jornada": "155BCB", "Repositorio": "155BCB", "NUGEP": "155BCB",
    "Glossario": "155BCB", "FAQ": "155BCB", "Siglas": "155BCB",
    "Parametros": "888888", "Listas": "888888",
}

th = Side(style="thin", color="CCCCCC")
BORDA = Border(left=th, right=th, top=th, bottom=th)
F_HEAD = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
F_CELL = Font(name=FONTE, size=10)
FILL_HEAD = PatternFill("solid", fgColor=AZUL)
FILL_FORM = PatternFill("solid", fgColor=CINZA_FORMULA)
FILL_ZEBRA = PatternFill("solid", fgColor=CINZA_ZEBRA)
AL_HEAD = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_WRAP = Alignment(vertical="top", wrap_text=True)
AL_TOP = Alignment(vertical="top")
AL_CENTER = Alignment(horizontal="center", vertical="top")

# ── Colunas novas das fichas (dados fictícios, como o resto do arquivo) ──
# "Unidades_Corresponsaveis" entra nas cinco abas da hierarquia e "Executor"
# (um cargo) só em Atividades. Ficam em tabelas à parte, indexadas pelo
# código, para não obrigar a mexer em cada literal de linha.
CORRESP_MP = {
    "MP-01": "AR/GDT; AI/GOM",
    "MP-02": "AI/GOM; AR/GRB",
    "MP-03": "AR/GRB; AA/GLC",
    "MP-04": "AA/GLC; AG/GGP",
    "MP-05": "AG/GGP; AT/GTI",
    "MP-06": "AT/GTI; AE/GPE",
    "MP-07": "AE/GPE; AE/GAG",
    "MP-08": "AE/GAG; AR/GDT",
}

CORRESP_PR = {
    "P-06.01": "AE/GAG; AR/GDT",
    "P-06.02": "AR/GDT; AI/GOM",
    "P-06.03": "AI/GOM; AR/GRB",
    "P-04.01": "AR/GRB; AA/GLC",
    "P-05.01": "AA/GLC; AG/GGP",
    "P-01.01": "AG/GGP; AT/GTI",
    "P-07.01": "AT/GTI; AE/GPE",
}

CORRESP_SP = {
    "SP-06.01.01": "AI/GOM",
    "SP-06.01.02": "AR/GRB",
    "SP-06.01.03": "AG/GGP",
    "SP-06.01.03.01": "AG/GGP",
    "SP-06.02.01": "AT/GTI",
    "SP-04.01.01": "AE/GPE",
}

# Subprocesso reutilizável (Call Activity do BPMN 2.0 — no Bizagi, que este
# painel usa para os diagramas, o próprio elemento já se chama "Subprocesso
# Reutilizável"): SP-06.01.03.01 mora nativamente dentro de SP-06.01.03 (2º
# nível de aninhamento), mas o mesmo tratamento de preços discrepantes é
# chamado também no julgamento de propostas (P-06.02, mesmo macroprocesso) e
# em pequenas compras da operação de perímetros (P-04.01, MACROPROCESSO
# DIFERENTE) — o caso que a marcação existe para evidenciar.
REUTILIZAVEL_SP = {"SP-06.01.03.01": "Sim"}
REUTILIZADO_EM_SP = {"SP-06.01.03.01": "P-06.02; P-04.01"}

CORRESP_AT = {
    "A-06.01.01.01": "AR/GRB",
    "A-06.01.01.02": "AA/GLC",
    "A-06.01.01.03": "AG/GGP",
    "A-06.01.01.04": "AT/GTI",
    "A-06.01.02.01": "AE/GPE",
    "A-06.01.02.02": "AE/GAG",
    "A-06.01.03.01": "AR/GDT",
    "A-06.01.03.02": "AI/GOM",
    "A-06.01.03.01.01": "AR/GRB",
    "A-06.02.01.01": "AA/GLC",
    "A-06.02.01.02": "AG/GGP",
    "A-04.01.01.01": "AT/GTI",
    "A-04.01.01.02": "AE/GPE",
    "A-06.03.01": "AE/GAG",
    "A-06.03.02": "AR/GDT",
}

CORRESP_TF = {
    "T-06.01.01.01.01": "AA/GLC",
    "T-06.01.01.01.02": "AG/GGP",
    "T-06.01.01.01.03": "AT/GTI",
    "T-06.01.03.01.01": "AE/GPE",
    "T-06.01.03.01.02": "AE/GAG",
    "T-06.01.03.01.03": "AR/GDT",
    "T-06.01.03.01.04": "AI/GOM",
    "T-06.02.01.01.01": "AR/GRB",
    "T-06.02.01.01.02": "AA/GLC",
    "T-04.01.01.01.01": "AG/GGP",
    "T-06.03.02.01": "AT/GTI",
    "T-06.03.02.02": "AE/GPE",
    "T-06.03.02.03": "AE/GAG",
}

EXECUTOR_AT = {
    "A-06.01.01.01": "Analista em Desenvolvimento Regional",
    "A-06.01.01.02": "Técnico em Desenvolvimento Regional",
    "A-06.01.01.03": "Assistente em Desenvolvimento Regional",
    "A-06.01.01.04": "Analista em Desenvolvimento Regional",
    "A-06.01.02.01": "Técnico em Desenvolvimento Regional",
    "A-06.01.02.02": "Assistente em Desenvolvimento Regional",
    "A-06.01.03.01": "Analista em Desenvolvimento Regional",
    "A-06.01.03.02": "Técnico em Desenvolvimento Regional",
    "A-06.01.03.01.01": "Assistente em Desenvolvimento Regional",
    "A-06.02.01.01": "Analista em Desenvolvimento Regional",
    "A-06.02.01.02": "Técnico em Desenvolvimento Regional",
    "A-04.01.01.01": "Assistente em Desenvolvimento Regional",
    "A-04.01.01.02": "Analista em Desenvolvimento Regional",
    "A-06.03.01": "Técnico em Desenvolvimento Regional",
    "A-06.03.02": "Assistente em Desenvolvimento Regional",
}


def com_colunas(linhas, base, *mapas):
    """Completa cada linha até `base` colunas — as listas de dados são
    propositalmente ragged — e anexa, na ordem, o valor de cada mapa,
    resolvido pelo código da 1ª coluna."""
    for ln in linhas:
        while len(ln) < base:
            ln.append("")
        for mapa in mapas:
            ln.append(mapa.get(ln[0], ""))
    return linhas


D = dt.date  # atalho


# NOTA: os cabeçalhos declarados nas chamadas de cabecalho() abaixo têm de
# bater com scripts/esquema_planilha.py, que é o que atualizar_planilha.py
# usa para acrescentar colunas à planilha já preenchida. Ao criar uma
# coluna nova, mexa nos DOIS arquivos — esquema_planilha.conferir() compara
# um dicionário {aba: [colunas]} com o esquema e aponta a diferença.
def cabecalho(ws, headers, widths, formula_cols=()):
    for j, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = AL_HEAD
        c.border = BORDA
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = 30
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws._formula_cols = set(formula_cols)  # marcação interna
    ws.sheet_properties.tabColor = COR_GUIA.get(ws.title, AZUL)
    # Zoom um pouco menor: as abas largas cabem sem rolagem horizontal
    # constante, que é o que mais cansa em preenchimento longo.
    ws.sheet_view.zoomScale = 90


def escreve(ws, linhas, wrap_cols=(), center_cols=(), pct_cols=(), date_cols=()):
    for i, linha in enumerate(linhas, start=2):
        for j, v in enumerate(linha, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = F_CELL
            c.border = BORDA
            if j in getattr(ws, "_formula_cols", ()):
                c.fill = FILL_FORM
            elif i % 2 == 0:
                c.fill = FILL_ZEBRA
            if j in wrap_cols:
                c.alignment = AL_WRAP
            elif j in center_cols:
                c.alignment = AL_CENTER
            else:
                c.alignment = AL_TOP
            if j in pct_cols:
                c.number_format = "0%"
            if j in date_cols:
                c.number_format = "DD/MM/YYYY"


def dv(ws, col_letter, listas_ref, ultima_linha=300):
    v = DataValidation(
        type="list", formula1=listas_ref, allow_blank=True,
        showErrorMessage=True, errorStyle="stop",
        errorTitle="Valor fora da lista",
        error=("Escolha uma das opções da seta ao lado da célula. "
               "A lista completa está na aba Listas."),
        showInputMessage=True,
        promptTitle="Lista de opções",
        prompt="Clique na seta ao lado da célula e escolha uma das opções.",
    )
    ws.add_data_validation(v)
    v.add(f"{col_letter}2:{col_letter}{ultima_linha}")


wb = Workbook()

# ----------------------------------------------------------------------------
# LISTAS (validação de dados)
# ----------------------------------------------------------------------------
ls = wb.active
ls.title = "Listas"
listas = {
    "Categoria_Macroprocesso": ["Gerencial", "Finalístico", "Suporte"],
    "Status_Mapeamento": ["Não iniciado", "Em andamento", "Concluído", "Suspenso"],
    "Fase_Ciclo_BPM": [
        "1. Alinhamento à estratégia e metas",
        "2. Arquitetar mudanças",
        "3. Desenvolver iniciativas",
        "4. Implementar mudanças",
        "5. Medir o sucesso",
    ],
    "Prioridade": ["Alta", "Média", "Baixa"],
    "Complexidade": ["Alta", "Média", "Baixa"],
    "Sim_Nao": ["Sim", "Não", "Não se aplica"],
    "Nivel_Vinculo": ["Macroprocesso", "Processo", "Subprocesso", "Atividade", "Tarefa"],
    "Tipo_Tarefa": ["Manual", "Automatizada", "Regra de negócio"],
    "Tipo_Documento": [
        "Procedimento (PRO)", "Manual", "Norma interna",
        "Formulário/Modelo", "Ata de reunião", "Diagrama BPMN", "Checklist",
        "Relatório", "Plano", "Outro",
    ],
    "Situacao_Documento": ["Vigente", "Em elaboração", "Em revisão", "Revogado"],
    "Categoria_Risco": [
        "Operacional", "Legal/Conformidade", "Pessoas",
        "Tecnologia da Informação", "Financeiro/Orçamentário", "Imagem/Reputação",
    ],
    "Resposta_Risco": ["Mitigar", "Aceitar", "Transferir", "Evitar"],
    "Status_Risco": ["Aberto", "Em tratamento", "Encerrado"],
    "Polaridade": ["Maior melhor", "Menor melhor"],
    "Periodicidade": ["Mensal", "Bimestral", "Trimestral", "Semestral", "Anual"],
    "Tipo_Registro": ["Reunião", "Oficina", "Entrevista", "Validação",
                      "Decisão", "Entrega", "Marco", "Nota"],
    "Fase_Jornada": ["Descobrir", "Definir", "Desenvolver", "Entregar", "Evoluir"],
    "Categoria_Repositorio": ["Documento oficial", "Template", "Instrumento",
                              "Ferramenta", "Referência"],
    "Fase_Instrumento": ["Planejamento", "Análise", "Desenho", "Implementação",
                         "Monitoramento", "Refinamento"],
    "Categoria_Glossario": ["BPM (CBOK)", "Projetos (PMBOK)", "Metodologia Codevasf",
                            "SIPOC e Modelagem", "Indicadores e Riscos", "Governança e Papéis"],
    "Categoria_FAQ": ["Conceitos básicos", "Modelagem e SIPOC", "Cadeia de Valor e governança",
                      "Indicadores, metas e riscos", "Plano de Ações AE/GPE", "Como usar o painel"],
}
for j, (nome, itens) in enumerate(listas.items(), start=1):
    c = ls.cell(row=1, column=j, value=nome)
    c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
    ls.column_dimensions[get_column_letter(j)].width = max(len(nome), max(len(x) for x in itens)) + 3
    for i, item in enumerate(itens, start=2):
        cc = ls.cell(row=i, column=j, value=item)
        cc.font = F_CELL; cc.border = BORDA
ls.freeze_panes = "A2"
ls.sheet_properties.tabColor = COR_GUIA["Listas"]
ls.sheet_view.zoomScale = 90

def ref(nome):
    j = list(listas).index(nome) + 1
    col = get_column_letter(j)
    n = len(listas[nome])
    return f"=Listas!${col}$2:${col}${n + 1}"

# ----------------------------------------------------------------------------
# MACROPROCESSOS
# ----------------------------------------------------------------------------
mp = wb.create_sheet("Macroprocessos")
cabecalho(mp,
    ["Codigo", "Nome", "Categoria", "Ordem", "Descricao", "Objetivo",
     "Unidade_Responsavel", "Dono_Processo", "Entregas",
     "Clientes_Beneficiarios", "Partes_Interessadas", "Sistemas",
     "Normativos_Aplicaveis", "Imagem_Bizagi", "Observacoes",
     "Unidades_Corresponsaveis"],
    [9, 34, 12, 7, 46, 46, 16, 26, 40, 34, 34, 30, 40, 30, 26, 30])
macros = [
    ["MP-01", "Gestão Estratégica e Governança", "Gerencial", 1,
     "Formulação, desdobramento e monitoramento da estratégia corporativa, da governança e do desempenho institucional.",
     "Garantir que a atuação da Companhia esteja alinhada à estratégia, com decisões baseadas em evidências.",
     "AE/GPE", "Helena Castro — Gerente de Planejamento",
     "Plano Estratégico; Plano de Ação Anual; Relatórios de desempenho",
     "Diretoria Executiva; Conselho de Administração; Unidades internas",
     "Ministério supervisor; Órgãos de controle; Sociedade",
     "e-Codevasf; Painéis de BI",
     "Estatuto Social; Regimento Interno; Lei nº 13.303/2016",
     "", ""],
    ["MP-02", "Gestão de Riscos, Integridade e Controles", "Gerencial", 2,
     "Identificação, avaliação e tratamento de riscos corporativos, integridade e controles internos (2ª linha).",
     "Assegurar razoável segurança para o alcance dos objetivos institucionais.",
     "AE/GAG", "Fernanda Alves — Gerente de Governança",
     "Política de Gestão de Riscos; Matriz de riscos corporativa; Plano de integridade",
     "Alta administração; Gestores de 1ª linha",
     "Auditoria interna; CGU; TCU",
     "e-Codevasf; Sistema de gestão de riscos",
     "Política de Gestão de Riscos (fictícia); IN Conjunta MP/CGU nº 01/2016",
     "", ""],
    ["MP-03", "Desenvolvimento Territorial e Estruturação Produtiva", "Finalístico", 3,
     "Apoio a arranjos produtivos, inclusão socioeconômica e estruturação de cadeias produtivas nos vales.",
     "Promover o desenvolvimento regional integrado e sustentável.",
     "AR/GDT", "Eduardo Martins — Gerente de Desenvolvimento",
     "Projetos de estruturação produtiva; Convênios e instrumentos de repasse",
     "Produtores rurais; Cooperativas; Municípios",
     "Parlamentares; Entes federados; Entidades parceiras",
     "TransfereGov; e-Codevasf",
     "Lei nº 14.133/2021; Portarias de transferências voluntárias",
     "", ""],
    ["MP-04", "Gestão de Empreendimentos de Irrigação", "Finalístico", 4,
     "Implantação, operação, manutenção e transferência de gestão de perímetros públicos de irrigação.",
     "Ampliar a área irrigada produtiva com sustentabilidade hídrica e econômica.",
     "AI/GOM", "Marcos Vinícius — Gerente de Operação",
     "Perímetros em operação; Água distribuída; Relatórios de O&M",
     "Irrigantes; Distritos de irrigação",
     "ANA; Agências estaduais; Associações de usuários",
     "SIG-Irrigação (fictício); e-Codevasf",
     "Lei nº 12.787/2013 (Política Nacional de Irrigação)",
     "", ""],
    ["MP-05", "Revitalização de Bacias Hidrográficas", "Finalístico", 5,
     "Ações de recuperação hidroambiental, segurança hídrica e uso sustentável dos recursos naturais.",
     "Contribuir para a revitalização das bacias dos rios São Francisco e Parnaíba.",
     "AR/GRB", "Luciana Prado — Gerente de Revitalização",
     "Nascentes recuperadas; Obras hidroambientais; Sistemas de abastecimento",
     "Comunidades ribeirinhas; Municípios",
     "Comitês de bacia; MMA; Órgãos ambientais",
     "e-Codevasf; GeoPortal (fictício)",
     "Legislação ambiental aplicável",
     "", ""],
    ["MP-06", "Gestão de Licitações e Contratos", "Suporte", 6,
     "Planejamento das contratações, seleção de fornecedores e gestão dos contratos administrativos da Companhia.",
     "Prover contratações tempestivas, vantajosas e conformes à legislação.",
     "AA/GLC", "Ricardo Nogueira — Gerente de Licitações",
     "Editais publicados; Contratos firmados; Atas de registro de preços",
     "Todas as unidades demandantes",
     "Fornecedores; Assessoria Jurídica; Órgãos de controle; PNCP",
     "Compras.gov.br; PNCP; Painel de Preços; e-Codevasf",
     "Lei nº 14.133/2021; Lei nº 13.303/2016; IN SEGES nº 65/2021; NI-027/2024 (fictícia)",
     "",
     "Macroprocesso priorizado no ciclo 2026 de mapeamento."],
    ["MP-07", "Gestão de Pessoas", "Suporte", 7,
     "Provimento, desenvolvimento, remuneração e qualidade de vida do corpo funcional.",
     "Assegurar pessoas qualificadas e engajadas para a missão institucional.",
     "AG/GGP", "Patrícia Ramos — Gerente de Pessoas",
     "Empregados admitidos e capacitados; Folha de pagamento",
     "Empregados; Gestores",
     "Sindicatos; SEST; Ministério supervisor",
     "SIGEP (fictício); e-Codevasf",
     "CLT; Normas internas de pessoal",
     "", ""],
    ["MP-08", "Gestão de Tecnologia da Informação", "Suporte", 8,
     "Planejamento, desenvolvimento, sustentação e segurança dos serviços de TI.",
     "Prover soluções digitais seguras que habilitem os processos de negócio.",
     "AT/GTI", "Daniela Ribeiro — Gerente de TI",
     "Sistemas em produção; Serviços de infraestrutura; Suporte ao usuário",
     "Todas as unidades",
     "SGD/MGI; Fornecedores de TI",
     "e-Codevasf; Service Desk",
     "IN SGD nº 94/2022; Política de Segurança da Informação (fictícia)",
     "", ""],
]
com_colunas(macros, 15, CORRESP_MP)
escreve(mp, macros, wrap_cols={5, 6, 9, 10, 11, 12, 13, 15, 16}, center_cols={3, 4})
mp.freeze_panes = "C2"
dv(mp, "C", ref("Categoria_Macroprocesso"))

# ----------------------------------------------------------------------------
# PROCESSOS
# ----------------------------------------------------------------------------
pr = wb.create_sheet("Processos")
cabecalho(pr,
    ["Codigo", "Macroprocesso", "Nome", "Descricao", "Objetivo",
     "Area_Responsavel", "Dono_Processo", "Interlocutor", "Prioridade",
     "Complexidade", "Status_Mapeamento", "Percentual", "Fase_Ciclo_BPM",
     "Inicio_Mapeamento", "Prazo_Previsto", "Data_Conclusao",
     "Fornecedores", "Entradas", "Saidas", "Beneficiarios", "Sistemas",
     "Normativos_Relacionados", "Processo_ECodevasf", "Imagem_Bizagi",
     "M1_Reuniao_Contextualizacao", "M2_Macro_Processo_Modelados",
     "M3_Subprocessos_Modelados", "M4_ASIS_Modelado", "M5_ASIS_Validado",
     "M6_Procedimento_Validado", "M7_Procedimento_Aprovado", "M8_TOBE_Elaborado",
     "M9_TOBE_Validado", "M10_Publicado_Repositorio", "Proxima_Acao", "Pendencia",
     "Ultima_Atualizacao", "Unidades_Corresponsaveis"],
    [10, 14, 34, 44, 40, 16, 24, 24, 10, 12, 14, 10, 30, 13, 13, 13,
     30, 34, 34, 28, 28, 40, 20, 30, 9, 9, 9, 9, 9, 9, 9, 9, 9, 34, 28, 13, 30])
S, N = "Sim", "Não"
# Marco que não se aplica ao processo — ex.: M3 "Subprocessos modelados" num
# processo que se decompõe direto em atividades, sem nível de subprocesso.
NA = "Não se aplica"
procs = [
    ["P-06.01", "MP-06", "Planejamento da Contratação",
     "Da identificação da necessidade (DFD) até a aprovação do ETP, TR e pesquisa de preços que instruem o certame.",
     "Instruir as contratações com estudos e artefatos completos, reduzindo retrabalho e impugnações.",
     "AA/GLC", "Ricardo Nogueira", "Carlos Eduardo Lima (UNP)",
     "Alta", "Alta", "Concluído", 1.00, "5. Medir o sucesso",
     D(2026, 1, 5), D(2026, 5, 29), D(2026, 5, 18),
     "Áreas demandantes; Painel de Preços; Fornecedores (cotações)",
     "DFD; Plano de Contratações Anual; Requisitos da área",
     "ETP aprovado; TR aprovado; Pesquisa de preços validada",
     "AA/GLC (Seleção do Fornecedor); Assessoria Jurídica",
     "e-Codevasf; Compras.gov.br; Painel de Preços",
     "Lei nº 14.133/2021; IN SEGES nº 65/2021; NI-027/2024 (fictícia)",
     "59500.000123/2026-11", "https://fluxosti.codevasf.gov.br/incidentes/",
     S, S, S, S, S, S, S, S, S, S,
     "Monitorar indicadores do processo e revisar o PRO em 12 meses.", "",
     D(2026, 7, 10)],
    ["P-06.02", "MP-06", "Seleção do Fornecedor",
     "Da divulgação do edital à homologação do resultado, incluindo sessão pública, julgamento e recursos.",
     "Selecionar a proposta mais vantajosa com celeridade e segurança jurídica.",
     "AA/GLC", "Ricardo Nogueira", "Carlos Eduardo Lima (UNP)",
     "Alta", "Média", "Em andamento", 0.60, "2. Arquitetar mudanças",
     D(2026, 6, 1), D(2026, 9, 30), None,
     "P-06.01 (artefatos); Compras.gov.br",
     "Edital minutado; Parecer jurídico; ETP/TR",
     "Resultado homologado; Contrato/ata para assinatura",
     "Unidades demandantes; Fornecedores",
     "Compras.gov.br; PNCP; e-Codevasf",
     "Lei nº 14.133/2021; NI-027/2024 (fictícia)",
     "59500.000456/2026-22", "",
     S, S, S, S, N, N, N, N, N, N,
     "Validar AS-IS com o dono do processo (reunião marcada).",
     "Pendente retorno da área sobre fluxo de recursos administrativos.",
     D(2026, 7, 15)],
    ["P-06.03", "MP-06", "Gestão e Fiscalização Contratual",
     "Da assinatura do contrato ao encerramento, incluindo fiscalização, medições, pagamentos e sanções.",
     "Garantir a entrega do objeto contratado no prazo, custo e qualidade pactuados.",
     "AA/GLC", "Ricardo Nogueira", "Bruna Souza (UNP)",
     "Alta", "Alta", "Em andamento", 0.35, "2. Arquitetar mudanças",
     D(2026, 6, 15), D(2026, 11, 30), None,
     "P-06.02 (contrato); Fornecedores contratados",
     "Contrato assinado; Cronograma; Garantias",
     "Objeto recebido; Pagamentos efetuados; Termo de encerramento",
     "Unidades demandantes; Fornecedores",
     "e-Codevasf; Compras.gov.br",
     "Lei nº 14.133/2021; Manual de Fiscalização (fictício)",
     "59500.000789/2026-33", "",
     # M3 (Subprocessos modelados) não se aplica: este processo não tem
     # subprocessos — as atividades penduram direto nele (ver aba Atividades).
     S, S, NA, N, N, N, N, N, N, N,
     "Concluir modelagem AS-IS das medições e pagamentos.", "",
     D(2026, 7, 20)],
    ["P-04.01", "MP-04", "Operação e Manutenção de Perímetros Irrigados",
     "Programação e distribuição de água, manutenção da infraestrutura de uso comum e relacionamento com irrigantes.",
     "Assegurar a distribuição hídrica confiável e a conservação dos ativos.",
     "AI/GOM", "Marcos Vinícius", "Eduardo Martins (UNP)",
     "Média", "Alta", "Em andamento", 0.45, "2. Arquitetar mudanças",
     D(2026, 5, 4), D(2026, 10, 30), None,
     "ANA (outorgas); Distritos de irrigação",
     "Outorga de uso; Demanda de irrigação; Plano de cultivo",
     "Água distribuída; Infraestrutura mantida; Relatórios de O&M",
     "Irrigantes; Distritos",
     "SIG-Irrigação (fictício); e-Codevasf",
     "Lei nº 12.787/2013",
     "59500.000901/2026-44", "",
     S, S, S, N, N, N, N, N, N, N,
     "Agendar oficina de validação do AS-IS com o distrito.", "",
     D(2026, 7, 12)],
    ["P-05.01", "MP-05", "Recuperação de Nascentes e Matas Ciliares",
     "Seleção de áreas, execução de cercamento e plantio, e monitoramento da regeneração.",
     "Recuperar áreas degradadas prioritárias das bacias.",
     "AR/GRB", "Luciana Prado", "Eduardo Martins (UNP)",
     "Média", "Média", "Não iniciado", 0.00, "1. Alinhamento à estratégia e metas",
     None, D(2027, 3, 31), None,
     "Municípios; Comitês de bacia",
     "Diagnóstico hidroambiental; Termos de cooperação",
     "Nascentes recuperadas; Relatórios de monitoramento",
     "Comunidades; Órgãos ambientais",
     "e-Codevasf; GeoPortal (fictício)",
     "Legislação ambiental aplicável",
     "", "",
     N, N, N, N, N, N, N, N, N, N,
     "Enviar formulário de levantamento à área (previsto ago/2026).", "",
     D(2026, 7, 1)],
    ["P-01.01", "MP-01", "Formulação e Monitoramento do Planejamento Estratégico",
     "Construção do plano estratégico, desdobramento em planos de ação e monitoramento periódico de resultados.",
     "Manter a estratégia viva, monitorada e comunicada.",
     "AE/GPE", "Helena Castro", "Gustavo Pereira (UNP)",
     "Alta", "Média", "Concluído", 1.00, "5. Medir o sucesso",
     D(2025, 9, 1), D(2026, 2, 27), D(2026, 2, 20),
     "Diretoria; Unidades",
     "Diretrizes de governo; Diagnóstico institucional",
     "Plano Estratégico; Painel de indicadores; RAG",
     "Diretoria Executiva; Conselhos",
     "e-Codevasf; Painéis de BI",
     "Lei nº 13.303/2016",
     "59500.000015/2025-77", "",
     S, S, S, S, S, S, S, S, S, S,
     "Ciclo de monitoramento trimestral (próximo: set/2026).", "",
     D(2026, 6, 30)],
    ["P-07.01", "MP-07", "Admissão e Integração de Empregados",
     "Da homologação do concurso à integração do novo empregado, incluindo exames, posse e ambientação.",
     "Admitir e integrar novos empregados com agilidade e conformidade.",
     "AG/GGP", "Patrícia Ramos", "Bruna Souza (UNP)",
     "Baixa", "Baixa", "Não iniciado", 0.00, "1. Alinhamento à estratégia e metas",
     None, D(2027, 6, 30), None,
     "Banca do concurso; Candidatos",
     "Resultado homologado; Documentação do candidato",
     "Empregado admitido e integrado",
     "Novos empregados; Unidades de lotação",
     "SIGEP (fictício); e-Codevasf",
     "CLT; Normas internas de pessoal",
     "", "",
     N, N, N, N, N, N, N, N, N, N,
     "Aguardando priorização no ciclo 2027.", "",
     D(2026, 6, 15)],
]
com_colunas(procs, 37, CORRESP_PR)
escreve(pr, procs,
        wrap_cols={4, 5, 17, 18, 19, 20, 21, 22, 35, 36},
        center_cols={9, 10, 11, 12} | set(range(25, 35)),
        pct_cols={12}, date_cols={14, 15, 16, 37})
pr.freeze_panes = "D2"
dv(pr, "I", ref("Prioridade")); dv(pr, "J", ref("Complexidade"))
dv(pr, "K", ref("Status_Mapeamento")); dv(pr, "M", ref("Fase_Ciclo_BPM"))
for k in range(25, 35):
    dv(pr, get_column_letter(k), ref("Sim_Nao"))

# ----------------------------------------------------------------------------
# SUBPROCESSOS
# ----------------------------------------------------------------------------
sp = wb.create_sheet("Subprocessos")
cabecalho(sp,
    ["Codigo", "Vinculo_Pai", "Ordem", "Nome", "Descricao", "Objetivo",
     "Unidade_Responsavel", "Dono", "Entradas", "Saidas", "Sistemas", "Imagem_Bizagi",
     "Unidades_Corresponsaveis", "Reutilizavel", "Reutilizado_Em"],
    [16, 16, 7, 34, 46, 40, 16, 24, 40, 40, 28, 30, 30, 13, 34])
subs = [
    ["SP-06.01.01", "P-06.01", 1, "Estudo Técnico Preliminar (ETP)",
     "Caracterização da necessidade, análise de soluções de mercado e demonstração da viabilidade da contratação.",
     "Fundamentar tecnicamente a melhor solução para a necessidade.",
     "AA/GLC", "Ricardo Nogueira",
     "DFD; Levantamento de soluções de mercado", "ETP aprovado no e-Codevasf", "e-Codevasf; Compras.gov.br", ""],
    ["SP-06.01.02", "P-06.01", 2, "Termo de Referência (TR)",
     "Definição do objeto, requisitos, modelo de execução e gestão contratual, e critérios de julgamento.",
     "Especificar com precisão o objeto e as condições da contratação.",
     "AA/GLC", "Ricardo Nogueira",
     "ETP aprovado; Modelos padronizados de TR", "TR aprovado e validado juridicamente", "e-Codevasf", ""],
    ["SP-06.01.03", "P-06.01", 3, "Pesquisa de Preços",
     "Levantamento de preços em fontes admitidas (Painel de Preços, PNCP, cotações) e consolidação do valor estimado.",
     "Estimar o valor da contratação conforme a IN SEGES nº 65/2021.",
     "AA/GLC", "Ricardo Nogueira",
     "TR validado; Fontes de preços admitidas",
     "Relatório de pesquisa de preços; Mapa comparativo", "Painel de Preços; PNCP; e-Codevasf",
     ""],
    # Exemplo de subprocesso DENTRO de outro subprocesso — o CBOK 4.0 não fixa a
    # profundidade da decomposição ("Levels Vary in Number and Name"): ao
    # mapear a Pesquisa de Preços, a equipe percebeu que o tratamento de
    # valores inexequíveis/excessivos merecia seu próprio subprocesso,
    # descoberto só durante as oficinas de modelagem (marco M3).
    # Vinculo_Pai aqui é OUTRO SUBPROCESSO (SP-...), não um Processo (P-...).
    ["SP-06.01.03.01", "SP-06.01.03", 1, "Tratamento de Preços Inexequíveis ou Excessivos",
     "Identificação e exclusão de valores discrepantes coletados na pesquisa, com justificativa técnica registrada.",
     "Evitar que outliers de preço distorçam o valor estimado da contratação.",
     "AA/GLC", "Ricardo Nogueira",
     "Amostra de preços coletados na pesquisa",
     "Valores discrepantes identificados e justificados", "Planilha padrão; e-Codevasf",
     ""],
    ["SP-06.02.01", "P-06.02", 1, "Condução da Sessão Pública",
     "Abertura da sessão no sistema, fase de lances, julgamento, habilitação e registro em ata.",
     "Conduzir o certame com transparência e celeridade.",
     "AA/GLC", "Ricardo Nogueira",
     "Edital publicado; Propostas dos licitantes",
     "Ata da sessão; Resultado por item", "Compras.gov.br", ""],
    ["SP-04.01.01", "P-04.01", 1, "Distribuição de Água aos Irrigantes",
     "Programação semanal, operação de comportas e bombas, e registro de volumes distribuídos.",
     "Entregar a água programada com eficiência e equidade.",
     "AI/GOM", "Marcos Vinícius",
     "Outorga de uso; Demanda de irrigação",
     "Programação hídrica executada; Registros de volume", "SIG-Irrigação (fictício)",
     ""],
]
com_colunas(subs, 12, CORRESP_SP, REUTILIZAVEL_SP, REUTILIZADO_EM_SP)
escreve(sp, subs, wrap_cols={5, 6, 9, 10, 11, 13, 15}, center_cols={3, 14})
sp.freeze_panes = "D2"

# ----------------------------------------------------------------------------
# ATIVIDADES
# ----------------------------------------------------------------------------
at = wb.create_sheet("Atividades")
cabecalho(at,
    # Vinculo_Pai (era "Subprocesso"): aceita o código de um SUBPROCESSO
    # (SP-...) ou de um PROCESSO (P-...). O painel identifica pelo prefixo do
    # código e monta ficha, breadcrumb e navegação sem o nível que não existe.
    ["Codigo", "Vinculo_Pai", "Ordem", "Nome", "Descricao",
     "Responsavel_Ator", "Entradas", "Saidas", "Sistemas",
     "Base_Normativa", "Imagem_Bizagi",
     "Executor", "Unidades_Corresponsaveis"],
    [16, 13, 7, 36, 46, 24, 38, 38, 24, 30, 44, 30, 30])
ativs = [
    ["A-06.01.01.01", "SP-06.01.01", 1, "Formalizar a necessidade (DFD)",
     "Registrar o Documento de Formalização da Demanda com justificativa, quantitativos e alinhamento ao PCA.",
     "Área demandante", "Necessidade identificada; Plano de Contratações Anual",
     "DFD assinado no e-Codevasf", "e-Codevasf",  "NI-027/2024 (fictícia)", "https://placehold.co/960x420/155bcb/ffffff?text=Diagrama+BPMN+A-06.01.01.01"],
    ["A-06.01.01.02", "SP-06.01.01", 2, "Levantar soluções de mercado",
     "Pesquisar soluções disponíveis, contratações similares e alternativas (inclusive não contratar).",
     "Equipe de planejamento da contratação", "DFD; Catálogos; Contratações similares (PNCP)",
     "Levantamento de soluções documentado", "PNCP; Compras.gov.br", 
     "Lei nº 14.133/2021, art. 18"],
    ["A-06.01.01.03", "SP-06.01.01", 3, "Estimar quantidades e resultados",
     "Definir quantitativos com memória de cálculo e resultados pretendidos com a contratação.",
     "Equipe de planejamento da contratação", "Levantamento de soluções; Séries históricas",
     "Memória de cálculo de quantitativos", "e-Codevasf",  ""],
    ["A-06.01.01.04", "SP-06.01.01", 4, "Elaborar e aprovar o ETP",
     "Consolidar o ETP no sistema e submeter à aprovação da autoridade competente.",
     "Equipe de planejamento; Autoridade competente", "Levantamentos e memórias anteriores",
     "ETP aprovado", "e-Codevasf; Compras.gov.br",  "Lei nº 14.133/2021, art. 18"],
    ["A-06.01.02.01", "SP-06.01.02", 1, "Redigir o Termo de Referência",
     "Elaborar o TR a partir do ETP, definindo objeto, requisitos, modelo de execução e critérios de medição.",
     "Equipe de planejamento da contratação", "ETP aprovado; Modelos padronizados",
     "Minuta de TR", "e-Codevasf",  "Lei nº 14.133/2021, art. 6º, XXIII"],
    ["A-06.01.02.02", "SP-06.01.02", 2, "Validar o TR com a Assessoria Jurídica",
     "Submeter a minuta ao órgão jurídico e ajustar conforme o parecer.",
     "AA/GLC; Assessoria Jurídica", "Minuta de TR",
     "TR validado; Parecer jurídico", "e-Codevasf",  "Lei nº 14.133/2021, art. 53"],
    ["A-06.01.03.01", "SP-06.01.03", 1, "Realizar a pesquisa de preços",
     "Consultar fontes admitidas, registrar parâmetros e tratar valores inexequíveis ou excessivos.",
     "Equipe de planejamento da contratação", "TR validado; Fontes de preços",
     "Relatório de pesquisa de preços", "Painel de Preços; PNCP", 
     "IN SEGES nº 65/2021"],
    ["A-06.01.03.02", "SP-06.01.03", 2, "Consolidar o valor estimado",
     "Aplicar o método definido (média/mediana/menor preço) e justificar a estimativa final.",
     "Equipe de planejamento da contratação", "Relatório de pesquisa",
     "Valor estimado consolidado", "e-Codevasf",  "IN SEGES nº 65/2021, art. 6º"],
    # Atividade do subprocesso ANINHADO SP-06.01.03.01 (subprocesso dentro de subprocesso)
    ["A-06.01.03.01.01", "SP-06.01.03.01", 1, "Aplicar critério de exclusão de valores discrepantes",
     "Comparar cada preço coletado à mediana da amostra e excluir, com justificativa registrada, os valores inexequíveis ou excessivos.",
     "Equipe de planejamento da contratação", "Amostra de preços coletados",
     "Amostra tratada; Justificativa de exclusão", "Planilha padrão",  "IN SEGES nº 65/2021, art. 6º, §3º", ""],
    # ── Processo SEM subprocessos: atividades ligadas DIRETO ao processo ──
    # Vinculo_Pai aqui é um PROCESSO (P-06.03), não um subprocesso. Acontece
    # quando o processo não precisa de decomposição intermediária — o CBOK 4.0
    # não obriga o nível subprocesso ("Levels Vary in Number and Name"). As
    # tarefas continuam penduradas na atividade, como sempre.
    ["A-06.03.01", "P-06.03", 1, "Designar fiscais e gestor do contrato",
     "Formalizar a designação do fiscal técnico, do fiscal administrativo e do gestor do contrato, com ciência dos designados.",
     "AA/GLC", "Contrato assinado; Minuta de portaria de designação",
     "Portaria de designação publicada", "e-Codevasf", 
     "Lei nº 14.133/2021, art. 117", ""],
    ["A-06.03.02", "P-06.03", 2, "Registrar medição e atestar a execução",
     "Conferir o objeto entregue no período, registrar a medição e atestar a nota fiscal para pagamento.",
     "Fiscal técnico do contrato", "Relatório de execução; Nota fiscal",
     "Medição registrada; Nota fiscal atestada", "e-Codevasf", 
     "Lei nº 14.133/2021, art. 140", ""],
    ["A-06.02.01.01", "SP-06.02.01", 1, "Publicar o edital",
     "Divulgar o edital no PNCP e no sistema de compras, observando prazos mínimos legais.",
     "Agente de contratação", "Edital aprovado; Parecer jurídico",
     "Edital publicado (PNCP)", "Compras.gov.br; PNCP", 
     "Lei nº 14.133/2021, art. 54"],
    ["A-06.02.01.02", "SP-06.02.01", 2, "Conduzir a sessão e julgar propostas",
     "Operar a fase de lances, julgar propostas, habilitar o vencedor e registrar a ata.",
     "Agente de contratação; Equipe de apoio", "Edital publicado; Propostas",
     "Ata da sessão; Resultado do julgamento", "Compras.gov.br", 
     "Lei nº 14.133/2021"],
    ["A-04.01.01.01", "SP-04.01.01", 1, "Programar a distribuição hídrica",
     "Elaborar a programação semanal de distribuição conforme demanda dos lotes e disponibilidade hídrica.",
     "Equipe de operação do perímetro", "Plano de cultivo; Disponibilidade hídrica",
     "Programação semanal aprovada", "SIG-Irrigação (fictício)",  ""],
    ["A-04.01.01.02", "SP-04.01.01", 2, "Operar e registrar a distribuição",
     "Operar comportas e bombas conforme programação e registrar volumes por ponto de entrega.",
     "Equipe de operação do perímetro", "Programação semanal",
     "Volumes registrados; Ocorrências", "SIG-Irrigação (fictício)",  ""],
]
com_colunas(ativs, 11, EXECUTOR_AT, CORRESP_AT)
escreve(at, ativs, wrap_cols={5, 7, 8, 9, 10, 12, 13}, center_cols={3})
at.freeze_panes = "D2"

# ----------------------------------------------------------------------------
# DOCUMENTOS
# ----------------------------------------------------------------------------
dc = wb.create_sheet("Documentos")
cabecalho(dc,
    ["ID", "Vinculo_Nivel", "Vinculo_Codigo", "Tipo_Documento", "Titulo",
     "Versao", "Data", "Situacao", "Link", "Observacoes"],
    [10, 15, 14, 26, 52, 8, 12, 14, 44, 30])
docs = [
    ["DOC-001", "Processo", "P-06.01", "Procedimento (PRO)",
     "PRO 06.01 — Planejamento da Contratação", "2.0", D(2026, 5, 18), "Vigente",
     "https://exemplo.codevasf.gov.br/repositorio/pop-06-01.pdf",
     "Publicado após validação do TO-BE."],
    ["DOC-002", "Processo", "P-06.01", "Diagrama BPMN",
     "Diagrama AS-IS — Planejamento da Contratação (Bizagi)", "1.0", D(2026, 3, 2),
     "Vigente", "https://fluxosti.codevasf.gov.br/incidentes/", "Exportado do Bizagi Modeler. Mesmo link da ficha do processo (P-06.01)."],
    ["DOC-003", "Processo", "P-06.01", "Diagrama BPMN",
     "Diagrama TO-BE — Planejamento da Contratação (Bizagi)", "1.0", D(2026, 5, 4),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/p-06-01-tobe.pdf", ""],
    ["DOC-004", "Processo", "P-06.01", "Ata de reunião",
     "Ata — Reunião de contextualização com a GLC", "1.0", D(2026, 1, 12), "Vigente",
     "https://exemplo.codevasf.gov.br/sei/ata-contextualizacao-p0601", "e-Codevasf 59500.000123/2026-11."],
    ["DOC-005", "Processo", "P-06.01", "Relatório",
     "Relatório de oportunidades de melhoria (AS-IS → TO-BE)", "1.0", D(2026, 4, 13),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/rel-melhorias-p0601.pdf", ""],
    ["DOC-006", "Subprocesso", "SP-06.01.01", "Formulário/Modelo",
     "Modelo de DFD — Documento de Formalização da Demanda", "3.1", D(2026, 2, 10),
     "Vigente", "https://exemplo.codevasf.gov.br/modelos/dfd.docx", ""],
    ["DOC-007", "Subprocesso", "SP-06.01.03", "Procedimento (PRO)",
     "Roteiro de pesquisa de preços (IN SEGES nº 65/2021)", "1.2", D(2026, 4, 27),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/roteiro-precos.pdf", ""],
    ["DOC-008", "Macroprocesso", "MP-06", "Norma interna",
     "NI-027/2024 — Norma de Licitações e Contratos (fictícia)", "4.0", D(2024, 11, 20),
     "Vigente", "https://exemplo.codevasf.gov.br/normativos/ni-027", ""],
    ["DOC-009", "Macroprocesso", "MP-06", "Manual",
     "Manual de Gestão e Fiscalização de Contratos (fictício)", "2.3", D(2025, 8, 14),
     "Em revisão", "https://exemplo.codevasf.gov.br/normativos/manual-fiscalizacao", ""],
    ["DOC-010", "Processo", "P-06.02", "Ata de reunião",
     "Ata — Contextualização do processo Seleção do Fornecedor", "1.0", D(2026, 6, 2),
     "Vigente", "https://exemplo.codevasf.gov.br/sei/ata-contextualizacao-p0602", ""],
    ["DOC-011", "Processo", "P-04.01", "Diagrama BPMN",
     "Diagrama AS-IS parcial — O&M de Perímetros (Bizagi)", "0.3", D(2026, 7, 8),
     "Em elaboração", "", "Modelagem em curso."],
    ["DOC-012", "Atividade", "A-06.01.01.01", "Formulário/Modelo",
     "Passo a passo do DFD no e-Codevasf", "1.0", D(2026, 2, 12), "Vigente",
     "https://exemplo.codevasf.gov.br/modelos/dfd-sei-passoapasso.pdf", ""],
    ["DOC-013", "Processo", "P-01.01", "Relatório",
     "Caderno de Indicadores Estratégicos 2026", "1.0", D(2026, 3, 31), "Vigente",
     "https://exemplo.codevasf.gov.br/estrategia/caderno-2026.pdf", ""],
    ["DOC-014", "Processo", "P-06.01", "Plano",
     "Plano de implantação do TO-BE (PMBOK — plano do projeto)", "1.0", D(2026, 5, 6),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/plano-tobe-p0601.pdf", ""],
    # Exemplo fictício — combinado internamente que TODOS os processos devem
    # ter um checklist (ainda pendente de padronização/rollout completo).
    ["DOC-015", "Processo", "P-06.01", "Checklist",
     "Checklist de Conformidade — Planejamento da Contratação",
     "1.0", D(2026, 5, 18), "Vigente",
     "https://exemplo.codevasf.gov.br/repositorio/checklist-p0601.pdf",
     "Exemplo — checklist ainda não padronizado para os demais processos da carteira."],
]
escreve(dc, docs, wrap_cols={5, 9, 10}, center_cols={6, 8}, date_cols={7})
dc.freeze_panes = "D2"
dv(dc, "B", ref("Nivel_Vinculo")); dv(dc, "D", ref("Tipo_Documento"))
dv(dc, "H", ref("Situacao_Documento"))

# ----------------------------------------------------------------------------
# RISCOS (Nivel e Classificacao por fórmula)
# ----------------------------------------------------------------------------
rs = wb.create_sheet("Riscos")
cabecalho(rs,
    ["ID", "Vinculo_Nivel", "Vinculo_Codigo", "Descricao_Risco", "Categoria",
     "Probabilidade_1a5", "Impacto_1a5", "Nivel_PxI", "Classificacao",
     "Resposta", "Controles_Tratamento", "Responsavel", "Status"],
    [9, 15, 14, 52, 22, 12, 10, 9, 13, 11, 46, 22, 14],
    formula_cols={8, 9})
riscos = [
    ["R-001", "Processo", "P-06.01",
     "Estimativas de preço defasadas gerando sobrepreço ou licitação deserta.",
     "Operacional", 4, 4, None, None, "Mitigar",
     "Roteiro de pesquisa com múltiplas fontes (Painel de Preços, PNCP) e revisão pela UNP.",
     "Ricardo Nogueira", "Em tratamento"],
    ["R-002", "Processo", "P-06.01",
     "TR genérico ou restritivo provocando impugnações e atrasos no certame.",
     "Legal/Conformidade", 3, 4, None, None, "Mitigar",
     "Checklist de revisão e validação obrigatória pela Assessoria Jurídica.",
     "Ricardo Nogueira", "Em tratamento"],
    ["R-003", "Processo", "P-06.03",
     "Fiscalização intempestiva de contratos, com medições e pagamentos atrasados.",
     "Operacional", 3, 5, None, None, "Mitigar",
     "Designação formal de fiscais, agenda de medições e alertas no e-Codevasf.",
     "Ricardo Nogueira", "Aberto"],
    ["R-004", "Subprocesso", "SP-06.01.01",
     "Conhecimento concentrado em um único empregado (pessoa-chave) na elaboração de ETP.",
     "Pessoas", 4, 3, None, None, "Mitigar",
     "Publicar o PRO, treinar substitutos e revezar a equipe de planejamento.",
     "Patrícia Ramos", "Em tratamento"],
    ["R-005", "Processo", "P-04.01",
     "Indisponibilidade hídrica comprometendo a programação de distribuição.",
     "Operacional", 2, 5, None, None, "Mitigar",
     "Plano de contingência hídrica e priorização de culturas conforme regras do perímetro.",
     "Marcos Vinícius", "Aberto"],
    ["R-006", "Processo", "P-06.02",
     "Instabilidade do Compras.gov.br durante a sessão pública.",
     "Tecnologia da Informação", 2, 4, None, None, "Mitigar",
     "Protocolo de suspensão formal da sessão e comunicação imediata aos licitantes.",
     "Daniela Ribeiro", "Aberto"],
    ["R-007", "Macroprocesso", "MP-06",
     "Alterações normativas frequentes exigindo atualização contínua de procedimentos.",
     "Legal/Conformidade", 4, 2, None, None, "Mitigar",
     "Monitoramento normativo mensal (resenha) e revisão programada dos PROs.",
     "Bruna Souza", "Em tratamento"],
    ["R-008", "Atividade", "A-06.01.03.01",
     "Uso de fontes de preço não admitidas pela IN SEGES nº 65/2021.",
     "Legal/Conformidade", 2, 3, None, None, "Mitigar",
     "Validação da pesquisa pela UNP antes da aprovação do valor estimado.",
     "Carlos Eduardo Lima", "Encerrado"],
    ["R-009", "Processo", "P-06.01",
     "Baixo engajamento das áreas demandantes no preenchimento do DFD.",
     "Pessoas", 3, 3, None, None, "Mitigar",
     "Oficinas de capacitação e modelo simplificado de DFD no e-Codevasf.",
     "Bruna Souza", "Encerrado"],
]
escreve(rs, riscos, wrap_cols={4, 11}, center_cols={6, 7, 8, 9, 10, 13})
for i in range(2, 2 + len(riscos)):
    rs.cell(row=i, column=8, value=f"=F{i}*G{i}")
    rs.cell(row=i, column=9,
            value=f'=IF(H{i}>=20,"Extremo",IF(H{i}>=12,"Alto",IF(H{i}>=5,"Moderado","Baixo")))')
    for col in (8, 9):
        c = rs.cell(row=i, column=col)
        c.font = F_CELL; c.border = BORDA; c.fill = FILL_FORM; c.alignment = AL_CENTER
rs.freeze_panes = "D2"
dv(rs, "B", ref("Nivel_Vinculo")); dv(rs, "E", ref("Categoria_Risco"))
dv(rs, "J", ref("Resposta_Risco")); dv(rs, "M", ref("Status_Risco"))

# ----------------------------------------------------------------------------
# INDICADORES (Situacao por fórmula)
# ----------------------------------------------------------------------------
ind = wb.create_sheet("Indicadores")
cabecalho(ind,
    ["ID", "Vinculo_Nivel", "Vinculo_Codigo", "Nome", "Descricao_Formula",
     "Unidade", "Polaridade", "Meta", "Resultado_Atual", "Situacao",
     "Periodicidade", "Fonte", "Ultima_Medicao"],
    [10, 15, 14, 40, 46, 9, 14, 8, 12, 15, 13, 24, 13],
    formula_cols={10})
inds = [
    ["IND-001", "Macroprocesso", "MP-06", "Tempo médio do ciclo de contratação",
     "Média de dias entre o DFD e a homologação do certame.", "dias", "Menor melhor",
     120, 148, None, "Trimestral", "Compras.gov.br", D(2026, 6, 30)],
    ["IND-002", "Processo", "P-06.01", "Prazo médio de elaboração do ETP",
     "Média de dias úteis entre o DFD e a aprovação do ETP.", "dias", "Menor melhor",
     30, 26, None, "Mensal", "e-Codevasf", D(2026, 6, 30)],
    ["IND-003", "Processo", "P-06.02", "Impugnações por edital",
     "Nº médio de impugnações recebidas por edital publicado.", "nº", "Menor melhor",
     1, 0.8, None, "Trimestral", "Compras.gov.br", D(2026, 6, 30)],
    ["IND-004", "Processo", "P-06.03", "Contratos com fiscal designado",
     "% de contratos vigentes com fiscal formalmente designado.", "%", "Maior melhor",
     100, 92, None, "Mensal", "e-Codevasf", D(2026, 6, 30)],
    ["IND-005", "Processo", "P-04.01", "Eficiência no uso da água",
     "Relação % entre volume faturado e volume captado no perímetro.", "%", "Maior melhor",
     75, 71, None, "Mensal", "SIG-Irrigação (fictício)", D(2026, 6, 30)],
    ["IND-006", "Macroprocesso", "MP-06", "Processos com mapeamento concluído",
     "% de processos do macroprocesso com marco M9 concluído.", "%", "Maior melhor",
     100, 33, None, "Trimestral", "Painel de Processos", D(2026, 6, 30)],
    ["IND-007", "Processo", "P-01.01", "Ações estratégicas monitoradas no prazo",
     "% de ações do plano com status atualizado no ciclo.", "%", "Maior melhor",
     95, 97, None, "Trimestral", "Painéis de BI", D(2026, 6, 30)],
    ["IND-008", "Subprocesso", "SP-06.01.03", "Pesquisas de preço com 3+ fontes",
     "% de pesquisas de preços com três ou mais fontes admitidas.", "%", "Maior melhor",
     100, 88, None, "Mensal", "e-Codevasf", D(2026, 6, 30)],
    ["IND-009", "Macroprocesso", "MP-05", "Nascentes recuperadas no ano",
     "Nº de nascentes com recuperação concluída no exercício.", "nº", "Maior melhor",
     120, None, None, "Anual", "GeoPortal (fictício)", None],
]
escreve(ind, inds, wrap_cols={4, 5, 12}, center_cols={6, 7, 8, 9, 10, 11},
        date_cols={13})
for i in range(2, 2 + len(inds)):
    ind.cell(row=i, column=10,
             value=(f'=IF(I{i}="","Sem medição",IF(G{i}="Maior melhor",'
                    f'IF(I{i}>=H{i},"Meta atingida","Abaixo da meta"),'
                    f'IF(I{i}<=H{i},"Meta atingida","Acima da meta")))'))
    c = ind.cell(row=i, column=10)
    c.font = F_CELL; c.border = BORDA; c.fill = FILL_FORM; c.alignment = AL_CENTER
ind.freeze_panes = "D2"
dv(ind, "B", ref("Nivel_Vinculo")); dv(ind, "G", ref("Polaridade"))
dv(ind, "K", ref("Periodicidade"))

# ----------------------------------------------------------------------------
# DIÁRIO DE MAPEAMENTO (evidências do trabalho — CBOK/PMBOK)
# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# TAREFAS (menor unidade de trabalho — CBOK 4.0)
# ----------------------------------------------------------------------------
tf = wb.create_sheet("Tarefas")
cabecalho(tf,
    ["Codigo", "Atividade", "Ordem", "Nome", "Descricao", "Tipo_Tarefa",
     "Responsavel", "Sistema", "Duracao_Estimada", "Observacoes", "Imagem_Bizagi",
     "Unidades_Corresponsaveis"],
    # Duracao_Estimada é NÚMERO (horas úteis; 1 dia útil = 8h) — a duração da
    # atividade é o somatório das tarefas, e a do processo, o somatório
    # recursivo das atividades (caminho crítico, não caminho feliz).
    [20, 16, 7, 38, 46, 16, 24, 22, 14, 28, 44, 30])
tarefas = [
    ["T-06.01.01.01.01", "A-06.01.01.01", 1, "Reunir informações da demanda",
     "Levantar justificativa, quantitativos preliminares e alinhamento ao PCA junto ao gestor da área.",
     "Manual", "Área demandante", "e-Codevasf", 4, ""],
    ["T-06.01.01.01.02", "A-06.01.01.01", 2, "Preencher o formulário DFD no e-Codevasf",
     "Registrar o Documento de Formalização da Demanda no modelo padronizado.",
     "Manual", "Área demandante", "e-Codevasf", 4, "Modelo DOC-006/DOC-012."],
    ["T-06.01.01.01.03", "A-06.01.01.01", 3, "Colher assinatura eletrônica do gestor",
     "Encaminhar o DFD para assinatura da autoridade competente da unidade.",
     "Manual", "Área demandante", "e-Codevasf", 8, ""],
    ["T-06.01.03.01.01", "A-06.01.03.01", 1, "Consultar o Painel de Preços",
     "Pesquisar contratações públicas similares e extrair os relatórios de preços.",
     "Manual", "Equipe de planejamento", "Painel de Preços", 4, ""],
    ["T-06.01.03.01.02", "A-06.01.03.01", 2, "Consultar contratações no PNCP",
     "Verificar atas e contratos vigentes de objetos equivalentes no PNCP.",
     "Manual", "Equipe de planejamento", "PNCP", 4, ""],
    ["T-06.01.03.01.03", "A-06.01.03.01", 3, "Registrar cotações de fornecedores",
     "Solicitar e registrar cotações diretas quando as fontes oficiais forem insuficientes.",
     "Manual", "Equipe de planejamento", "e-Codevasf", 24, "Mínimo de 3 fontes (IN 65/2021)."],
    ["T-06.01.03.01.04", "A-06.01.03.01", 4, "Aplicar tratamento estatístico",
     "Calcular média/mediana, excluir valores inexequíveis ou excessivos e justificar o método.",
     "Regra de negócio", "Equipe de planejamento", "Planilha padrão", 4, ""],
    # Tarefas de uma atividade que pende DIRETO do processo (P-06.03, sem
    # subprocessos) — o nível tarefa não muda em nada nesse caso.
    ["T-06.03.02.01", "A-06.03.02", 1, "Conferir o objeto entregue",
     "Verificar quantidade, qualidade e conformidade da entrega com o Termo de Referência.",
     "Manual", "Fiscal técnico do contrato", "e-Codevasf", 8, ""],
    ["T-06.03.02.02", "A-06.03.02", 2, "Registrar a medição no processo",
     "Lançar a medição do período no processo do contrato, anexando relatórios e evidências.",
     "Manual", "Fiscal técnico do contrato", "e-Codevasf", 4, ""],
    ["T-06.03.02.03", "A-06.03.02", 3, "Atestar a nota fiscal",
     "Atestar a nota fiscal depois da conferência e encaminhá-la para pagamento.",
     "Manual", "Gestor do contrato", "e-Codevasf", 4,
     "O prazo de pagamento conta a partir do ateste."],
    ["T-06.02.01.01.01", "A-06.02.01.01", 1, "Cadastrar o edital no Compras.gov.br",
     "Inserir o edital aprovado, anexos e cronograma do certame no sistema.",
     "Manual", "Agente de contratação", "Compras.gov.br", 4, ""],
    ["T-06.02.01.01.02", "A-06.02.01.01", 2, "Publicar o aviso no PNCP",
     "Divulgação automática do aviso de licitação a partir do cadastro no sistema.",
     "Automatizada", "Compras.gov.br", "PNCP", 0, ""],
    ["T-04.01.01.01.01", "A-04.01.01.01", 1, "Consolidar demandas semanais dos lotes",
     "Compilar os pedidos de água dos irrigantes por setor hidráulico para a programação.",
     "Manual", "Equipe de operação", "SIG-Irrigação (fictício)", 8, ""],
]
com_colunas(tarefas, 11, CORRESP_TF)
escreve(tf, tarefas, wrap_cols={5, 10, 12}, center_cols={3, 6, 9})
# Duracao_Estimada (coluna 9) fica numérica — sem formatação de texto.
tf.freeze_panes = "D2"
dv(tf, "F", ref("Tipo_Tarefa"))

# ----------------------------------------------------------------------------
# JORNADA · REPOSITORIO · NUGEP · GLOSSARIO · FAQ · PARAMETROS
# (conteúdo em scripts/dados_conteudo.py — edite depois direto na planilha)
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# SIGLAS (Lista de Nomes e Siglas das Unidades Orgânicas — Decisão da
# Presidência nº 601, de 11/04/2025). Fica na planilha, não no código, para
# a UNP corrigir um nome de unidade sem precisar editar nada além da célula.
# ----------------------------------------------------------------------------
sg = wb.create_sheet("Siglas")
cabecalho(sg, ["Sigla", "Nome"], [20, 60])
SIGLAS = [
    ("10ª/AJ", "Assessoria Jurídica Regional"),
    ("10ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("10ª/EAG", "Escritório de Apoio Técnico de Araguaína"),
    ("10ª/GB", "Gabinete da Superintendência Regional"),
    ("10ª/GGR", "Gerência de Gestão Regional"),
    ("10ª/GTR", "Gerência Técnica Regional"),
    ("10ª/SL", "Secretaria Regional de Licitações"),
    ("10ª/SR", "10ª Superintendência Regional"),
    ("11ª/AJ", "Assessoria Jurídica Regional"),
    ("11ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("11ª/GB", "Gabinete da Superintendência Regional"),
    ("11ª/GGR", "Gerência de Gestão Regional"),
    ("11ª/GGR/UOF", "Unidade Regional de Orçamento e Finanças"),
    ("11ª/GGR/UOF/SCO", "Setor de Contabilidade"),
    ("11ª/GGR/URA", "Unidade Regional de Administração"),
    ("11ª/GRD", "Gerência Regional de Infraestrutura"),
    ("11ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("11ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("11ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("11ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("11ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("11ª/SL", "Secretaria Regional de Licitações"),
    ("11ª/SR", "11ª Superintendência Regional"),
    ("12ª/AJ", "Assessoria Jurídica Regional"),
    ("12ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("12ª/EMO", "Escritório de Apoio Técnico de Mossoró"),
    ("12ª/GB", "Gabinete da Superintendência Regional"),
    ("12ª/GGR", "Gerência de Gestão Regional"),
    ("12ª/GTR", "Gerência Técnica Regional"),
    ("12ª/SL", "Secretaria Regional de Licitações"),
    ("12ª/SR", "12ª Superintendência Regional"),
    ("13ª/AJ", "Assessoria Jurídica Regional"),
    ("13ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("13ª/EPT", "Escritório de Apoio Técnico de Patos"),
    ("13ª/GB", "Gabinete da Superintendência Regional"),
    ("13ª/GGR", "Gerência de Gestão Regional"),
    ("13ª/GTR", "Gerência Técnica Regional"),
    ("13ª/SL", "Secretaria Regional de Licitações"),
    ("13ª/SR", "13ª Superintendência Regional"),
    ("14ª/AJ", "Assessoria Jurídica Regional"),
    ("14ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("14ª/ETA", "Escritório de Apoio Técnico de Tauá"),
    ("14ª/GB", "Gabinete da Superintendência Regional"),
    ("14ª/GGR", "Gerência de Gestão Regional"),
    ("14ª/GTR", "Gerência Técnica Regional"),
    ("14ª/SL", "Secretaria Regional de Licitações"),
    ("14ª/SR", "14ª Superintendência Regional"),
    ("15ª/AJ", "Assessoria Jurídica Regional"),
    ("15ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("15ª/GB", "Gabinete da Superintendência Regional"),
    ("15ª/GGR", "Gerência de Gestão Regional"),
    ("15ª/GTR", "Gerência Técnica Regional"),
    ("15ª/SL", "Secretaria Regional de Licitações"),
    ("15ª/SR", "15ª Superintendência Regional"),
    ("16ª/AJ", "Assessoria Jurídica Regional"),
    ("16ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("16ª/EPI", "Escritório de Apoio Técnico de Piumhi"),
    ("16ª/GB", "Gabinete da Superintendência Regional"),
    ("16ª/GGR", "Gerência de Gestão Regional"),
    ("16ª/GTR", "Gerência Técnica Regional"),
    ("16ª/SL", "Secretaria Regional de Licitações"),
    ("16ª/SR", "16ª Superintendência Regional"),
    ("1ª/AJ", "Assessoria Jurídica Regional"),
    ("1ª/CIG", "Centro Integrado de Recursos Pesqueiros e Aquicultura de Gorutuba"),
    ("1ª/CIM", "Centro Integrado de Recursos Pesqueiros e Aquicultura de Três Marias"),
    ("1ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("1ª/EAL", "Escritório de Apoio Técnico de Almenara"),
    ("1ª/GB", "Gabinete da Superintendência Regional"),
    ("1ª/GB/URC", "Unidade Regional de Comunicação"),
    ("1ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("1ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("1ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("1ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("1ª/GRD", "Gerência Regional de Infraestrutura"),
    ("1ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("1ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("1ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("1ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("1ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("1ª/GRG/UFN/SCB", "Setor de Cobrança"),
    ("1ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("1ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("1ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("1ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("1ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("1ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("1ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("1ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("1ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("1ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("1ª/SL", "Secretaria Regional de Licitações"),
    ("1ª/SR", "1ª Superintendência Regional"),
    ("2ª/AJ", "Assessoria Jurídica Regional"),
    ("2ª/CIX", "Centro Integrado de Recursos Pesqueiros e Aquicultura de Xique-Xique"),
    ("2ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("2ª/EGU", "Escritório de Apoio Técnico de Guanambi"),
    ("2ª/EIR", "Escritório de Apoio Técnico de Irecê"),
    ("2ª/GB", "Gabinete da Superintendência Regional"),
    ("2ª/GB/URC", "Unidade Regional de Comunicação"),
    ("2ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("2ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("2ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("2ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("2ª/GRD", "Gerência Regional de Infraestrutura"),
    ("2ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("2ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("2ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("2ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("2ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("2ª/GRG/UFN/SCB", "Setor de Cobrança"),
    ("2ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("2ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("2ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("2ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("2ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("2ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("2ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("2ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("2ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("2ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("2ª/SL", "Secretaria Regional de Licitações"),
    ("2ª/SR", "2ª Superintendência Regional"),
    ("2ª/UBA", "Unidade Regional Descentralizada de Barreiras"),
    ("2ª/UVC", "Unidade Regional Descentralizada de Vitória da Conquista"),
    ("3ª/AJ", "Assessoria Jurídica Regional"),
    ("3ª/CIB", "Centro Integrado de Recursos Pesqueiros e Aquicultura de Bebedouro"),
    ("3ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("3ª/GB", "Gabinete da Superintendência Regional"),
    ("3ª/GB/URC", "Unidade Regional de Comunicação"),
    ("3ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("3ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("3ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("3ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("3ª/GRD", "Gerência Regional de Infraestrutura"),
    ("3ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("3ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("3ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("3ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("3ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("3ª/GRG/UFN/SCB", "Setor de Cobrança"),
    ("3ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("3ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("3ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("3ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("3ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("3ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("3ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("3ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("3ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("3ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("3ª/SL", "Secretaria Regional de Licitações"),
    ("3ª/SR", "3ª Superintendência Regional"),
    ("4ª/AJ", "Assessoria Jurídica Regional"),
    ("4ª/CIT", "Centro Integrado de Recursos Pesqueiros e Aquicultura de Betume"),
    ("4ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("4ª/EPR", "Escritório de Apoio Técnico de Propriá"),
    ("4ª/GB", "Gabinete da Superintendência Regional"),
    ("4ª/GB/URC", "Unidade Regional de Comunicação"),
    ("4ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("4ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("4ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("4ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("4ª/GRD", "Gerência Regional de Infraestrutura"),
    ("4ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("4ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("4ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("4ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("4ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("4ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("4ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("4ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("4ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("4ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("4ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("4ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("4ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("4ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("4ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("4ª/SL", "Secretaria Regional de Licitações"),
    ("4ª/SR", "4ª Superintendência Regional"),
    ("5ª/AJ", "Assessoria Jurídica Regional"),
    ("5ª/CII", "Centro Integrado de Recursos Pesqueiros e Aquicultura de Itiúba"),
    ("5ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("5ª/EPE", "Escritório de Apoio Técnico de Penedo"),
    ("5ª/GB", "Gabinete da Superintendência Regional"),
    ("5ª/GB/URC", "Unidade Regional de Comunicação"),
    ("5ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("5ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("5ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("5ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("5ª/GRD", "Gerência Regional de Infraestrutura"),
    ("5ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("5ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("5ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("5ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("5ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("5ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("5ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("5ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("5ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("5ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("5ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("5ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("5ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("5ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("5ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("5ª/SL", "Secretaria Regional de Licitações"),
    ("5ª/SR", "5ª Superintendência Regional"),
    ("6ª/AJ", "Assessoria Jurídica Regional"),
    ("6ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("6ª/ESA", "Escritório de Representação de Salvador"),
    ("6ª/GB", "Gabinete da Superintendência Regional"),
    ("6ª/GB/URC", "Unidade Regional de Comunicação"),
    ("6ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("6ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("6ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("6ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("6ª/GRD", "Gerência Regional de Infraestrutura"),
    ("6ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("6ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("6ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("6ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("6ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("6ª/GRG/UFN/SCB", "Setor de Cobrança"),
    ("6ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("6ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("6ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("6ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("6ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("6ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("6ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("6ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("6ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("6ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("6ª/SL", "Secretaria Regional de Licitações"),
    ("6ª/SR", "6ª Superintendência Regional"),
    ("6ª/UPF", "Unidade Regional Descentralizada de Paulo Afonso"),
    ("7ª/AJ", "Assessoria Jurídica Regional"),
    ("7ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("7ª/EOE", "Escritório de Apoio Técnico de Oeiras"),
    ("7ª/EPA", "Escritório de Apoio Técnico de Parnaíba"),
    ("7ª/ERN", "Escritório de Apoio Técnico de São Raimundo Nonato"),
    ("7ª/GB", "Gabinete da Superintendência Regional"),
    ("7ª/GB/URC", "Unidade Regional de Comunicação"),
    ("7ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("7ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("7ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("7ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("7ª/GRD", "Gerência Regional de Infraestrutura"),
    ("7ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("7ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("7ª/GRD/UPS", "Unidade Regional de Projetos Especiais"),
    ("7ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("7ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("7ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("7ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("7ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("7ª/GRI", "Gerência Regional de Irrigação e Operações"),
    ("7ª/GRI/UEI", "Unidade Regional de Gestão dos Empreendimentos de Irrigação"),
    ("7ª/GRI/URP", "Unidade Regional de Apoio à Produção"),
    ("7ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("7ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("7ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("7ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("7ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("7ª/SL", "Secretaria Regional de Licitações"),
    ("7ª/SR", "7ª Superintendência Regional"),
    ("8ª/AJ", "Assessoria Jurídica Regional"),
    ("8ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("8ª/EIM", "Escritório de Apoio Técnico de Imperatriz"),
    ("8ª/EPD", "Escritório de Apoio Técnico de Presidente Dutra"),
    ("8ª/GB", "Gabinete da Superintendência Regional"),
    ("8ª/GB/URC", "Unidade Regional de Comunicação"),
    ("8ª/GRA", "Gerência Regional de Administração e Tecnologia"),
    ("8ª/GRA/UGP", "Unidade Regional de Gestão de Pessoas"),
    ("8ª/GRA/USA", "Unidade Regional de Patrimônio, Logística e Serviços Auxiliares"),
    ("8ª/GRA/UTI", "Unidade Regional de Tecnologia da Informação"),
    ("8ª/GRD", "Gerência Regional de Infraestrutura"),
    ("8ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("8ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("8ª/GRD/UPS", "Unidade Regional de Projetos Especiais"),
    ("8ª/GRG", "Gerência Regional de Estratégia e Finanças"),
    ("8ª/GRG/UCB", "Unidade Regional de Contabilidade"),
    ("8ª/GRG/UFN", "Unidade Regional de Finanças e Cobrança"),
    ("8ª/GRG/UMC", "Unidade Regional de Monitoramento e de Controle de Contratos e Convênios"),
    ("8ª/GRG/URO", "Unidade Regional Orçamentária"),
    ("8ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("8ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("8ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("8ª/GRS", "Gerência Regional de Suporte Integrado"),
    ("8ª/GRS/ULF", "Unidade Regional de Licenciamento Ambiental e Administração Fundiária"),
    ("8ª/SL", "Secretaria Regional de Licitações"),
    ("8ª/SR", "8ª Superintendência Regional"),
    ("9ª/AJ", "Assessoria Jurídica Regional"),
    ("9ª/Cogex", "Comitê de Gestão Executiva da Superintendência Regional"),
    ("9ª/GB", "Gabinete da Superintendência Regional"),
    ("9ª/GGR", "Gerência de Gestão Regional"),
    ("9ª/GGR/UOF", "Unidade Regional de Orçamento e Finanças"),
    ("9ª/GGR/UOF/SCO", "Setor de Contabilidade"),
    ("9ª/GGR/URA", "Unidade Regional de Administração"),
    ("9ª/GRD", "Gerência Regional de Infraestrutura"),
    ("9ª/GRD/UEP", "Unidade Regional de Estudos e Projetos"),
    ("9ª/GRD/UIP", "Unidade Regional de Implantação e Acompanhamento de Projetos"),
    ("9ª/GRR", "Gerência Regional de Revitalização e Desenvolvimento Territorial"),
    ("9ª/GRR/UDT", "Unidade Regional de Desenvolvimento Territorial"),
    ("9ª/GRR/UES", "Unidade Regional de Empreendimentos Socioambientais"),
    ("9ª/SL", "Secretaria Regional de Licitações"),
    ("9ª/SR", "9ª Superintendência Regional"),
    ("AA", "Área de Administração e Tecnologia"),
    ("AA/GGP", "Gerência de Gestão de Pessoas"),
    ("AA/GGP/UBS", "Unidade de Benefícios e Saúde Ocupacional"),
    ("AA/GGP/UCP", "Unidade de Cadastro e Pagamento"),
    ("AA/GGP/UDP", "Unidade de Desenvolvimento de Pessoas"),
    ("AA/GGP/URT", "Unidade de Relações de Trabalho"),
    ("AA/GPA", "Gerência de Patrimônio, Logística e Serviços Auxiliares"),
    ("AA/GPA/SBI", "Setor de Biblioteca"),
    ("AA/GPA/SDP", "Setor de Documentação e Protocolo"),
    ("AA/GPA/UAL", "Unidade de Administração Predial e Logística"),
    ("AA/GPA/UCS", "Unidade de Compras Administrativas"),
    ("AA/GPA/UPM", "Unidade de Patrimônio e Material"),
    ("AA/GTI", "Gerência de Tecnologia da Informação"),
    ("AA/GTI/UIT", "Unidade de Infraestrutura e Tecnologia"),
    ("AA/GTI/UPC", "Unidade de Conformidade e Controle de Tecnologia da Informação"),
    ("AA/GTI/USC", "Unidade de Segurança Cibernética"),
    ("AA/GTI/USI", "Unidade de Sistemas de Informações"),
    ("AD", "Área de Desenvolvimento e Infraestrutura"),
    ("AD/GEP", "Gerência de Estudos e Projetos"),
    ("AD/GEP/UPE", "Unidade de Estudos e Projetos de Infraestrutura Urbana e Edificações"),
    ("AD/GEP/UPH", "Unidade de Estudos e Projetos de Infraestrutura Hídrica e Irrigação"),
    ("AD/GIM", "Gerência de Implantação de Obras"),
    ("AD/GIM/UOE", "Unidade de Implantação de Obras de Infraestrutura Urbana e Edificações"),
    ("AD/GIM/UOH", "Unidade de Implantação de Obras de Infraestrutura Hídrica e Irrigação"),
    ("AD/GPI", "Gerência de Parcerias e Investimentos"),
    ("AD/GQV", "Gerência de Qualificação Viária"),
    ("AD/GQV/UGV", "Unidade de Gestão e Controle de Projetos de Qualificação Viária"),
    ("AD/GQV/UIO", "Unidade de Implantação de Obras de Qualificação Viária"),
    ("AD/SE", "Secretaria Executiva"),
    ("AE", "Área de Estratégia e Finanças"),
    ("AE/GCB", "Gerência de Contabilidade"),
    ("AE/GCB/UCC", "Unidade de Conformidade Contábil"),
    ("AE/GCB/UEC", "Unidade de Execução Contábil"),
    ("AE/GCB/UER", "Unidade de Escrituração e Revisão Contábil"),
    ("AE/GFN", "Gerência de Finanças"),
    ("AE/GFN/UCR", "Unidade de Gestão da Cobrança"),
    ("AE/GFN/UEF", "Unidade de Programação e Execução Financeira"),
    ("AE/GGO", "Gerência de Gestão Orçamentária"),
    ("AE/GGO/UEO", "Unidade de Execução Orçamentária"),
    ("AE/GGO/UPO", "Unidade de Programação Orçamentária"),
    ("AE/GPE", "Gerência de Planejamento Estratégico"),
    ("AE/GPE/UAT", "Unidade de Acompanhamento de Transferências"),
    ("AE/GPE/UNP", "Unidade de Gestão Normativa e de Processos"),
    ("AE/GPE/UPL", "Unidade de Planejamento Institucional"),
    ("AG", "Área de Governança e Sustentabilidade"),
    ("AG/GAF", "Gerência de Administração Fundiária e Geotecnologia"),
    ("AG/GAF/UDF", "Unidade de Documentação Fundiária"),
    ("AG/GAF/UGG", "Unidade de Gestão Geotecnológica"),
    ("AG/GCT", "Gerência de Custos"),
    ("AG/GCT/UCT", "Unidade de Custos e Orçamentação"),
    ("AG/GCT/UTR", "Unidade de Procedimentos Técnicos Referenciais"),
    ("AG/GMA", "Gerência de Meio Ambiente"),
    ("AG/GMA/UCA", "Unidade de Conservação Ambiental"),
    ("AG/GMA/ULA", "Unidade de Licenciamento Ambiental"),
    ("AG/SE", "Secretaria Executiva"),
    ("AGE", "Assembleia Geral"),
    ("AI", "Área de Irrigação e Operações"),
    ("AI/GAP", "Gerência de Apoio à Produção Irrigada"),
    ("AI/GAP/UGI", "Unidade de Gestão de Áreas Irrigadas"),
    ("AI/GAP/UPI", "Unidade de Apoio aos Projetos Públicos de Irrigação"),
    ("AI/GEE", "Gerência de Eficiência Energética"),
    ("AI/GEE/UEE", "Unidade de Gestão de Energia e Apoio ao PISF"),
    ("AI/GEE/USB", "Unidade de Gestão e Segurança de Barragens"),
    ("AI/GGE", "Gerência de Gestão de Empreendimentos"),
    ("AI/GGE/UGM", "Unidade de Gestão e Monitoramento de Empreendimentos"),
    ("AI/GGE/UOR", "Unidade de Operação e Reabilitação de Empreendimentos"),
    ("AI/SE", "Secretaria Executiva"),
    ("AR", "Área de Revitalização e Desenvolvimento Territorial"),
    ("AR/GDT", "Gerência de Desenvolvimento Territorial"),
    ("AR/GDT/UAP", "Unidade de Desenvolvimento dos Arranjos Produtivos"),
    ("AR/GDT/UPA", "Unidade de Recursos Pesqueiros, Aquicultura e Beneficiamento"),
    ("AR/GMT", "Gerência de Mecanização e Modernização Territorial"),
    ("AR/GMT/UCM", "Unidade de Gestão e Controle da Modernização Territorial"),
    ("AR/GMT/UME", "Unidade de Aquisição de Máquinas e Equipamentos"),
    ("AR/GSA", "Gerência de Saneamento e Acesso à Água"),
    ("AR/GSA/UAS", "Unidade de Água e Saneamento Básico"),
    ("AR/GSA/UEA", "Unidade de Esgotamento Rural e Acesso a Água"),
    ("AR/SE", "Secretaria Executiva"),
    ("CEC", "Comissão de Ética da Codevasf"),
    ("Coaud", "Comitê de Auditoria Estatutário"),
    ("Cogid", "Comitê de Governança Interna e Digital"),
    ("Cogid/SPDP", "Subcomitê de Proteção de Dados Pessoais"),
    ("Cogid/SSI", "Subcomitê de Segurança da Informação"),
    ("Cogid/STIC", "Subcomitê de Tecnologia da Informação"),
    ("Confis", "Conselho Fiscal"),
    ("Consad", "Conselho de Administração"),
    ("Consad/AUD", "Auditoria Interna"),
    ("Consad/AUD/UGE", "Unidade de Acompanhamento e Avaliação da Gestão"),
    ("Consad/AUD/UIN", "Unidade de Apoio e Informação"),
    ("Consad/AUD/UPR", "Unidade de Acompanhamento e Avaliação de Programas"),
    ("Consad/COR", "Corregedoria"),
    ("Consad/COR/UIV", "Unidade Especial de Admissibilidade e Investigação"),
    ("Consad/COR/UPJ", "Unidade Especial de Processos Acusatórios e Julgamento"),
    ("Consad/OUV", "Ouvidoria"),
    ("Copes", "Comitê de Pessoas, Elegibilidade, Sucessão e Remuneração"),
    ("DEX", "Diretoria Executiva"),
    ("PR", "Presidência"),
    ("PR/AC", "Assessoria de Comunicação"),
    ("PR/AC/UCE", "Unidade Especial de Cerimonial e Eventos"),
    ("PR/AC/UIM", "Unidade Especial de Imagem e Promoção Institucional"),
    ("PR/AJ", "Assessoria Jurídica"),
    ("PR/AJ/UAA", "Unidade de Assuntos Administrativos"),
    ("PR/AJ/UCO", "Unidade do Contencioso"),
    ("PR/EBE", "Escritório de Representação de Belém"),
    ("PR/GB", "Gabinete da Presidência"),
    ("PR/SC", "Secretaria de Órgãos Colegiados"),
    ("PR/SI", "Secretaria de Integridade"),
    ("PR/SLC", "Secretaria de Licitações e Contratos"),
    ("PR/SLC/UGC", "Unidade Especial de Gestão de Contratos"),
    ("PR/SLC/UGL", "Unidade Especial de Gestão de Licitações"),
    ("PR/SRC", "Secretaria de Gestão de Riscos e Controle Interno"),
    ("PR/SRC/UGR", "Unidade Especial de Gerenciamento de Riscos Corporativos"),
]
escreve(sg, [list(x) for x in SIGLAS], wrap_cols={2})
sg.freeze_panes = "A2"

jn = wb.create_sheet("Jornada")
cabecalho(jn, ["Ordem", "Fase", "Nome", "Duracao", "Objetivo", "Atividades_Chave",
               "Quem_Faz", "Entregaveis", "Sentimento_Usuario"],
          [7, 13, 30, 12, 44, 60, 30, 40, 40])
escreve(jn, [list(x) for x in CONTEUDO.JORNADA], wrap_cols={5, 6, 7, 8, 9}, center_cols={1})
jn.freeze_panes = "D2"
dv(jn, "B", ref("Fase_Jornada"))

rp = wb.create_sheet("Repositorio")
cabecalho(rp, ["ID", "Categoria", "Fase_Ciclo", "Codigo", "Titulo", "Descricao",
               "Fonte", "Link", "Ordem"],
          [10, 18, 15, 12, 42, 56, 12, 50, 7])
escreve(rp, [list(x) for x in CONTEUDO.REPOSITORIO], wrap_cols={5, 6, 8}, center_cols={9})
rp.freeze_panes = "E2"
dv(rp, "B", ref("Categoria_Repositorio")); dv(rp, "C", ref("Fase_Instrumento"))

ng = wb.create_sheet("NUGEP")
cabecalho(ng, ["Ordem", "Nome", "Papel", "Unidade_Sigla", "Unidade_Nome", "Email", "Telefone",
               "Foto", "Hierarquia"],
          [7, 24, 34, 14, 38, 32, 16, 46, 11])
escreve(ng, [list(x) for x in CONTEUDO.NUGEP], wrap_cols={3, 5}, center_cols={1, 9})
ng.freeze_panes = "C2"

gl = wb.create_sheet("Glossario")
cabecalho(gl, ["Termo", "Categoria", "Definicao", "Fonte", "Termos_Relacionados"],
          [34, 22, 80, 18, 40])
escreve(gl, [list(x) for x in CONTEUDO.GLOSSARIO], wrap_cols={3, 5})
gl.freeze_panes = "B2"
dv(gl, "B", ref("Categoria_Glossario"))

fq = wb.create_sheet("FAQ")
cabecalho(fq, ["Ordem", "Categoria", "Pergunta", "Resposta"], [7, 26, 50, 90])
escreve(fq, [list(x) for x in CONTEUDO.FAQ], wrap_cols={3, 4}, center_cols={1})
fq.freeze_panes = "C2"
dv(fq, "B", ref("Categoria_FAQ"))

pm = wb.create_sheet("Parametros")
cabecalho(pm, ["Chave", "Valor"], [24, 90])
escreve(pm, [list(x) for x in CONTEUDO.PARAMETROS], wrap_cols={2})

# ----------------------------------------------------------------------------
# LEIA-ME
# ----------------------------------------------------------------------------
lm = wb.create_sheet("LEIA-ME", 0)
lm.sheet_view.showGridLines = False
lm.column_dimensions["A"].width = 3
lm.column_dimensions["B"].width = 30
lm.column_dimensions["C"].width = 100
lm.column_dimensions["D"].width = 16

def titulo(r, txt, size=14, cor=AZUL_ESCURO):
    c = lm.cell(row=r, column=2, value=txt)
    c.font = Font(name=FONTE, size=size, bold=True, color=cor)

def linha(r, rotulo, texto):
    a = lm.cell(row=r, column=2, value=rotulo)
    a.font = Font(name=FONTE, size=10, bold=True)
    a.alignment = AL_TOP
    b = lm.cell(row=r, column=3, value=texto)
    b.font = F_CELL
    b.alignment = AL_WRAP

titulo(2, "Painel de Processos — Base de Dados (DADOS FICTÍCIOS)", 16)
lm.cell(row=3, column=2, value="Fonte única de dados do painel publicado no GitHub Pages. "
        "Preencha as abas e o site refletirá o conteúdo.").font = Font(name=FONTE, size=10, italic=True)
c = lm.cell(row=4, column=2, value="ATENÇÃO: todos os nomes, números, normativos internos e resultados são "
            "FICTÍCIOS, criados apenas para demonstração do painel.")
c.font = Font(name=FONTE, size=10, bold=True, color="C0392B")

titulo(6, "Como o painel usa esta planilha")
linha(7, "1. Publicação", "Salve este arquivo em data/painel-processos-dados.xlsx no repositório. "
      "O site lê a planilha diretamente no navegador (SheetJS).")
linha(8, "2. Fallback", "Opcionalmente, gere js/dados.js com o script scripts/planilha_para_js.py "
      "(usado quando o site é aberto sem servidor/offline).")
linha(9, "3. Vínculos", "Hierarquia (CBOK 4.0): Macroprocesso → Processo de negócio (aba Processos) → "
      "Subprocesso → Atividade → Tarefa. Os relacionamentos usam os CÓDIGOS: "
      "Processos→Macroprocesso; Subprocessos→Vinculo_Pai (Processo P-... ou outro Subprocesso SP-..., quando aninhado); "
      "Atividades→Vinculo_Pai (Subprocesso SP-... OU Processo P-..., quando o processo não tem subprocessos); Tarefas→Atividade; "
      "Documentos/Riscos/Indicadores usam Vinculo_Nivel + Vinculo_Codigo.")

titulo(11, "Convenções de preenchimento")
linha(12, "Listas na célula", "Separe múltiplos itens com ponto e vírgula ( ; ). "
      "Ex.: 'e-Codevasf; Compras.gov.br; PNCP'.")
linha(14, "Percentual", "Na aba Processos, use percentual (0% a 100%).")
linha(15, "Datas", "Formato dd/mm/aaaa.")
linha(16, "Diagrama (Bizagi Web Publish)", "O painel incorpora o diagrama por iframe direto da publicação web "
      "do Bizagi (Bizagi Modeler → Publish → Web) — mesmo padrão já usado na Base de Conhecimento da Codevasf "
      "(AA/GTI) para o Gerenciamento de Incidentes de TI. Publique o modelo, suba o pacote num servidor e cole "
      "a URL da publicação em Imagem_Bizagi (Macroprocessos, Processos ou Subprocessos — único nível sem "
      "diagrama são Atividades e Tarefas). Exige que o servidor permita ser incorporado por outra origem "
      "(o painel roda no GitHub Pages); sem isso, o quadro aparece em branco, mas o botão \"Abrir em nova "
      "aba\" sempre funciona. Deixe em branco enquanto não houver publicação — o painel mostra um aviso.")
linha(17, "Células de fórmula", "Colunas com fundo cinza (Riscos: Nivel_PxI e Classificacao; "
      "Indicadores: Situacao) são CALCULADAS — não digite valores; ao inserir linhas, copie a fórmula da linha acima.")
linha(18, "Validação de dados", "Campos com lista suspensa buscam os valores na aba 'Listas' "
      "(edite lá para incluir novas opções).")

titulo(20, "Dicionário de abas")
abas_desc = [
    ("Macroprocessos", "1º nível da cadeia de valor (CBOK: processos primários/finalísticos, de suporte e gerenciais)."),
    ("Processos", "2º nível — processos de negócio, com ficha completa: SIPOC (fornecedores, entradas, saídas, beneficiários), status e marcos do mapeamento (M1–M10), fase do ciclo BPM e dados do projeto."),
    ("Subprocessos", "3º nível (CBOK): subprocessos do processo de negócio. Vinculo_Pai aceita o código de um Processo (P-...) OU de OUTRO subprocesso (SP-...) — o CBOK não fixa a profundidade da decomposição, então um subprocesso pode conter outro subprocesso, tantos níveis quanto o processo exigir."),
    ("Atividades", "4º nível (CBOK), com entradas, saídas, ator, sistemas e prazos. Vinculo_Pai aceita o código de um Subprocesso (SP-...) OU de um Processo (P-...): processo que não tem subprocesso se decompõe direto em atividades, e estas em tarefas. Nesse caso, marque M3 (Subprocessos modelados) como \"Não se aplica\" na aba Processos."),
    ("Tarefas", "5º e último nível (CBOK): menor unidade de trabalho de uma atividade — manual, automatizada ou regra de negócio."),
    ("Documentos", "Repositório: Procedimentos (PRO), manuais, atas, diagramas BPMN (Bizagi), relatórios — vinculados a qualquer nível."),
    ("Riscos", "Riscos vinculados a qualquer nível; nível = Probabilidade × Impacto (matriz 5×5)."),
    ("Indicadores", "Indicadores de desempenho por nível, com meta, resultado e situação calculada."),
    ("Jornada", "Etapas da jornada de mapeamento (Descobrir → Evoluir), exibidas na aba Repositório do painel."),
    ("Repositorio", "Materiais e ferramentas: metodologia e guia oficiais (RES 031/2025), templates, instrumentos por fase do ciclo BPM, ferramentas e referências."),
    ("NUGEP", "Integrantes do Núcleo de Gestão Normativa e de Processos (aba NUGEP do painel). "
              "Foto: URL de imagem pública que aparece no avatar. Hierarquia: 1 = Gerente-Executivo (AE), "
              "2 = Gerente (AE/GPE), 3 = equipe da Unidade (AE/GPE/UNP), 0 = interlocutor de outra área."),
    ("Glossario", "Termos BPM (CBOK), PMBOK e metodologia Codevasf (aba Glossário)."),
    ("FAQ", "Perguntas e respostas exibidas na aba FAQ."),
    ("Parametros", "Configurações chave/valor: contato do NUGEP e links da metodologia e do guia."),
    ("Listas", "Domínios das listas suspensas (validação de dados)."),
]
r = 21
for nome, desc in abas_desc:
    linha(r, nome, desc)
    r += 1

titulo(r + 1, "Conteúdo atual (calculado)")
contagens = [
    ("Macroprocessos", "=COUNTA(Macroprocessos!$A$2:$A$500)"),
    ("Processos", "=COUNTA(Processos!$A$2:$A$500)"),
    ("Subprocessos", "=COUNTA(Subprocessos!$A$2:$A$500)"),
    ("Atividades", "=COUNTA(Atividades!$A$2:$A$500)"),
    ("Tarefas", "=COUNTA(Tarefas!$A$2:$A$500)"),
    ("Documentos", "=COUNTA(Documentos!$A$2:$A$500)"),
    ("Riscos", "=COUNTA(Riscos!$A$2:$A$500)"),
    ("Indicadores", "=COUNTA(Indicadores!$A$2:$A$500)"),
    ("Etapas da jornada", "=COUNTA(Jornada!$A$2:$A$500)"),
    ("Itens do repositório", "=COUNTA(Repositorio!$A$2:$A$500)"),
    ("Integrantes do NUGEP", "=COUNTA(NUGEP!$A$2:$A$500)"),
    ("Termos do glossário", "=COUNTA(Glossario!$A$2:$A$500)"),
    ("Perguntas do FAQ", "=COUNTA(FAQ!$A$2:$A$500)"),
]
r2 = r + 2
for nome, f in contagens:
    a = lm.cell(row=r2, column=2, value=nome)
    a.font = Font(name=FONTE, size=10, bold=True)
    b = lm.cell(row=r2, column=3, value=f)
    b.font = F_CELL
    b.fill = FILL_FORM
    b.alignment = Alignment(horizontal="left")
    r2 += 1

titulo(r2 + 1, "Referências metodológicas")
linha(r2 + 2, "BPM CBOK 4.0", "Hierarquia e tipos de processos, ciclo de vida BPM (5 fases), papéis "
      "(dono do processo), SIPOC, indicadores e repositório de processos (ABPMP, 2019).")
linha(r2 + 3, "PMBOK", "Gestão do projeto de mapeamento: termo de abertura, escopo, partes interessadas, "
      "riscos do projeto, entregáveis e lições aprendidas (PMI).")

ordem_final = ["LEIA-ME", "Macroprocessos", "Processos", "Subprocessos", "Atividades",
               "Tarefas", "Documentos", "Riscos", "Indicadores", "Jornada",
               "Repositorio", "NUGEP", "Glossario", "FAQ", "Siglas", "Parametros", "Listas"]
wb._sheets = [wb[n] for n in ordem_final]
wb.active = 0
from pathlib import Path

DESTINO = Path(__file__).resolve().parent.parent / "data" / "painel-processos-dados.xlsx"
wb.save(DESTINO)

# O openpyxl grava a fórmula, nunca o resultado dela — o cache de valor do XML
# fica vazio e qualquer leitor que não seja o Excel (o painel, entre eles) lê
# célula nula nas colunas calculadas. Este passo recalcula e grava `<f>` e `<v>`
# juntos, como o Excel faria. Sem ele a planilha sai quebrada para o painel.
from cachear_formulas import cachear  # noqa: E402

print(f"Planilha gerada com sucesso ({cachear(DESTINO)} fórmulas com valor calculado).")
