# -*- coding: utf-8 -*-
"""
Gera a planilha data/painel-processos-dados.xlsx com DADOS FICTÍCIOS,
fonte de dados do Painel de Processos (GitHub Pages).

Estrutura (alinhada ao BPM CBOK 4.0 e ao PMBOK):
  LEIA-ME | Macroprocessos | Processos | Subprocessos | Atividades |
  Documentos | Riscos | Indicadores | Diario_Mapeamento | Listas
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

AZUL = "1351B4"      # azul gov.br (blue-warm-vivid-70)
AZUL_ESCURO = "0C326F"
CINZA_FORMULA = "F0F0F0"
FONTE = "Arial"

th = Side(style="thin", color="CCCCCC")
BORDA = Border(left=th, right=th, top=th, bottom=th)
F_HEAD = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
F_CELL = Font(name=FONTE, size=10)
FILL_HEAD = PatternFill("solid", fgColor=AZUL)
FILL_FORM = PatternFill("solid", fgColor=CINZA_FORMULA)
AL_HEAD = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_WRAP = Alignment(vertical="top", wrap_text=True)
AL_TOP = Alignment(vertical="top")
AL_CENTER = Alignment(horizontal="center", vertical="top")

D = dt.date  # atalho


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


def escreve(ws, linhas, wrap_cols=(), center_cols=(), pct_cols=(), date_cols=()):
    for i, linha in enumerate(linhas, start=2):
        for j, v in enumerate(linha, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = F_CELL
            c.border = BORDA
            if j in getattr(ws, "_formula_cols", ()):
                c.fill = FILL_FORM
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
    v = DataValidation(type="list", formula1=listas_ref, allow_blank=True, showErrorMessage=False)
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
    "Sim_Nao": ["Sim", "Não"],
    "Nivel_Vinculo": ["Macroprocesso", "Processo", "Subprocesso", "Atividade", "Tarefa"],
    "Tipo_Tarefa": ["Manual", "Automatizada", "Regra de negócio"],
    "Tipo_Documento": [
        "Procedimento Operacional (POP)", "Manual", "Norma interna",
        "Formulário/Modelo", "Ata de reunião", "Diagrama BPMN",
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
     "Normativos_Aplicaveis", "Imagem_Bizagi", "Observacoes"],
    [9, 34, 12, 7, 46, 46, 16, 26, 40, 34, 34, 30, 40, 30, 26])
macros = [
    ["MP-01", "Gestão Estratégica e Governança", "Gerencial", 1,
     "Formulação, desdobramento e monitoramento da estratégia corporativa, da governança e do desempenho institucional.",
     "Garantir que a atuação da Companhia esteja alinhada à estratégia, com decisões baseadas em evidências.",
     "AE/GPE", "Helena Castro — Gerente de Planejamento",
     "Plano Estratégico; Plano de Ação Anual; Relatórios de desempenho",
     "Diretoria Executiva; Conselho de Administração; Unidades internas",
     "Ministério supervisor; Órgãos de controle; Sociedade",
     "e-Codevasf; SEI; Painéis de BI",
     "Estatuto Social; Regimento Interno; Lei nº 13.303/2016",
     "img/diagramas/mp-01.svg", ""],
    ["MP-02", "Gestão de Riscos, Integridade e Controles", "Gerencial", 2,
     "Identificação, avaliação e tratamento de riscos corporativos, integridade e controles internos (2ª linha).",
     "Assegurar razoável segurança para o alcance dos objetivos institucionais.",
     "AE/GAG", "Fernanda Alves — Gerente de Governança",
     "Política de Gestão de Riscos; Matriz de riscos corporativa; Plano de integridade",
     "Alta administração; Gestores de 1ª linha",
     "Auditoria interna; CGU; TCU",
     "SEI; Sistema de gestão de riscos",
     "Política de Gestão de Riscos (fictícia); IN Conjunta MP/CGU nº 01/2016",
     "img/diagramas/mp-02.svg", ""],
    ["MP-03", "Desenvolvimento Territorial e Estruturação Produtiva", "Finalístico", 3,
     "Apoio a arranjos produtivos, inclusão socioeconômica e estruturação de cadeias produtivas nos vales.",
     "Promover o desenvolvimento regional integrado e sustentável.",
     "AR/GDT", "Eduardo Martins — Gerente de Desenvolvimento",
     "Projetos de estruturação produtiva; Convênios e instrumentos de repasse",
     "Produtores rurais; Cooperativas; Municípios",
     "Parlamentares; Entes federados; Entidades parceiras",
     "TransfereGov; SEI",
     "Lei nº 14.133/2021; Portarias de transferências voluntárias",
     "img/diagramas/mp-03.svg", ""],
    ["MP-04", "Gestão de Empreendimentos de Irrigação", "Finalístico", 4,
     "Implantação, operação, manutenção e transferência de gestão de perímetros públicos de irrigação.",
     "Ampliar a área irrigada produtiva com sustentabilidade hídrica e econômica.",
     "AI/GOM", "Marcos Vinícius — Gerente de Operação",
     "Perímetros em operação; Água distribuída; Relatórios de O&M",
     "Irrigantes; Distritos de irrigação",
     "ANA; Agências estaduais; Associações de usuários",
     "SIG-Irrigação (fictício); SEI",
     "Lei nº 12.787/2013 (Política Nacional de Irrigação)",
     "img/diagramas/mp-04.svg", ""],
    ["MP-05", "Revitalização de Bacias Hidrográficas", "Finalístico", 5,
     "Ações de recuperação hidroambiental, segurança hídrica e uso sustentável dos recursos naturais.",
     "Contribuir para a revitalização das bacias dos rios São Francisco e Parnaíba.",
     "AR/GRB", "Luciana Prado — Gerente de Revitalização",
     "Nascentes recuperadas; Obras hidroambientais; Sistemas de abastecimento",
     "Comunidades ribeirinhas; Municípios",
     "Comitês de bacia; MMA; Órgãos ambientais",
     "SEI; GeoPortal (fictício)",
     "Legislação ambiental aplicável",
     "img/diagramas/mp-05.svg", ""],
    ["MP-06", "Gestão de Licitações e Contratos", "Suporte", 6,
     "Planejamento das contratações, seleção de fornecedores e gestão dos contratos administrativos da Companhia.",
     "Prover contratações tempestivas, vantajosas e conformes à legislação.",
     "AA/GLC", "Ricardo Nogueira — Gerente de Licitações",
     "Editais publicados; Contratos firmados; Atas de registro de preços",
     "Todas as unidades demandantes",
     "Fornecedores; Assessoria Jurídica; Órgãos de controle; PNCP",
     "Compras.gov.br; PNCP; Painel de Preços; SEI",
     "Lei nº 14.133/2021; Lei nº 13.303/2016; IN SEGES nº 65/2021; NI-027/2024 (fictícia)",
     "img/diagramas/mp-06.svg",
     "Macroprocesso priorizado no ciclo 2026 de mapeamento."],
    ["MP-07", "Gestão de Pessoas", "Suporte", 7,
     "Provimento, desenvolvimento, remuneração e qualidade de vida do corpo funcional.",
     "Assegurar pessoas qualificadas e engajadas para a missão institucional.",
     "AG/GGP", "Patrícia Ramos — Gerente de Pessoas",
     "Empregados admitidos e capacitados; Folha de pagamento",
     "Empregados; Gestores",
     "Sindicatos; SEST; Ministério supervisor",
     "SIGEP (fictício); SEI",
     "CLT; Normas internas de pessoal",
     "img/diagramas/mp-07.svg", ""],
    ["MP-08", "Gestão de Tecnologia da Informação", "Suporte", 8,
     "Planejamento, desenvolvimento, sustentação e segurança dos serviços de TI.",
     "Prover soluções digitais seguras que habilitem os processos de negócio.",
     "AT/GTI", "Daniela Ribeiro — Gerente de TI",
     "Sistemas em produção; Serviços de infraestrutura; Suporte ao usuário",
     "Todas as unidades",
     "SGD/MGI; Fornecedores de TI",
     "e-Codevasf; Service Desk",
     "IN SGD nº 94/2022; Política de Segurança da Informação (fictícia)",
     "img/diagramas/mp-08.svg", ""],
]
escreve(mp, macros, wrap_cols={5, 6, 9, 10, 11, 12, 13, 15}, center_cols={3, 4})
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
     "Fornecedores", "Entradas", "Saidas", "Clientes", "Sistemas",
     "Normativos_Relacionados", "Processo_SEI", "Imagem_Bizagi",
     "M1_Formulario_Enviado", "M2_Formulario_Retornado",
     "M3_Reuniao_Contextualizacao", "M4_ASIS_Modelado", "M5_ASIS_Validado",
     "M6_Normativos_Identificados", "M7_TOBE_Elaborado", "M8_TOBE_Validado",
     "M9_Publicado_Repositorio", "Proxima_Acao", "Pendencia",
     "Ultima_Atualizacao"],
    [10, 14, 34, 44, 40, 16, 24, 24, 10, 12, 14, 10, 30, 13, 13, 13,
     30, 34, 34, 28, 28, 40, 20, 30, 9, 9, 9, 9, 9, 9, 9, 9, 9, 34, 28, 13])
S, N = "Sim", "Não"
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
     "SEI; Compras.gov.br; Painel de Preços",
     "Lei nº 14.133/2021; IN SEGES nº 65/2021; NI-027/2024 (fictícia)",
     "59500.000123/2026-11", "img/diagramas/p-06-01.svg",
     S, S, S, S, S, S, S, S, S,
     "Monitorar indicadores do processo e revisar POP em 12 meses.", "",
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
     "Compras.gov.br; PNCP; SEI",
     "Lei nº 14.133/2021; NI-027/2024 (fictícia)",
     "59500.000456/2026-22", "img/diagramas/p-06-02.svg",
     S, S, S, S, N, N, N, N, N,
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
     "SEI; Compras.gov.br",
     "Lei nº 14.133/2021; Manual de Fiscalização (fictício)",
     "59500.000789/2026-33", "img/diagramas/p-06-03.svg",
     S, S, S, N, N, N, N, N, N,
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
     "SIG-Irrigação (fictício); SEI",
     "Lei nº 12.787/2013",
     "59500.000901/2026-44", "img/diagramas/p-04-01.svg",
     S, S, S, S, N, N, N, N, N,
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
     "SEI; GeoPortal (fictício)",
     "Legislação ambiental aplicável",
     "", "img/diagramas/p-05-01.svg",
     N, N, N, N, N, N, N, N, N,
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
     "59500.000015/2025-77", "img/diagramas/p-01-01.svg",
     S, S, S, S, S, S, S, S, S,
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
     "SIGEP (fictício); SEI",
     "CLT; Normas internas de pessoal",
     "", "img/diagramas/p-07-01.svg",
     N, N, N, N, N, N, N, N, N,
     "Aguardando priorização no ciclo 2027.", "",
     D(2026, 6, 15)],
]
escreve(pr, procs,
        wrap_cols={4, 5, 17, 18, 19, 20, 21, 22, 34, 35},
        center_cols={9, 10, 11, 12} | set(range(25, 34)),
        pct_cols={12}, date_cols={14, 15, 16, 36})
pr.freeze_panes = "D2"
dv(pr, "I", ref("Prioridade")); dv(pr, "J", ref("Complexidade"))
dv(pr, "K", ref("Status_Mapeamento")); dv(pr, "M", ref("Fase_Ciclo_BPM"))
for col in "YZ" + "".join(get_column_letter(k) for k in range(27, 34)):
    pass
for k in range(25, 34):
    dv(pr, get_column_letter(k), ref("Sim_Nao"))

# ----------------------------------------------------------------------------
# SUBPROCESSOS
# ----------------------------------------------------------------------------
sp = wb.create_sheet("Subprocessos")
cabecalho(sp,
    ["Codigo", "Processo", "Ordem", "Nome", "Descricao", "Objetivo",
     "Unidade_Responsavel", "Dono", "Entregas", "Sistemas", "Imagem_Bizagi"],
    [13, 10, 7, 34, 46, 40, 16, 24, 40, 28, 30])
subs = [
    ["SP-06.01.01", "P-06.01", 1, "Estudo Técnico Preliminar (ETP)",
     "Caracterização da necessidade, análise de soluções de mercado e demonstração da viabilidade da contratação.",
     "Fundamentar tecnicamente a melhor solução para a necessidade.",
     "AA/GLC", "Ricardo Nogueira",
     "ETP aprovado no SEI", "SEI; Compras.gov.br", "img/diagramas/sp-06-01-01.svg"],
    ["SP-06.01.02", "P-06.01", 2, "Termo de Referência (TR)",
     "Definição do objeto, requisitos, modelo de execução e gestão contratual, e critérios de julgamento.",
     "Especificar com precisão o objeto e as condições da contratação.",
     "AA/GLC", "Ricardo Nogueira",
     "TR aprovado e validado juridicamente", "SEI", "img/diagramas/sp-06-01-02.svg"],
    ["SP-06.01.03", "P-06.01", 3, "Pesquisa de Preços",
     "Levantamento de preços em fontes admitidas (Painel de Preços, PNCP, cotações) e consolidação do valor estimado.",
     "Estimar o valor da contratação conforme a IN SEGES nº 65/2021.",
     "AA/GLC", "Ricardo Nogueira",
     "Relatório de pesquisa de preços; Mapa comparativo", "Painel de Preços; PNCP; SEI",
     "img/diagramas/sp-06-01-03.svg"],
    ["SP-06.02.01", "P-06.02", 1, "Condução da Sessão Pública",
     "Abertura da sessão no sistema, fase de lances, julgamento, habilitação e registro em ata.",
     "Conduzir o certame com transparência e celeridade.",
     "AA/GLC", "Ricardo Nogueira",
     "Ata da sessão; Resultado por item", "Compras.gov.br", "img/diagramas/sp-06-02-01.svg"],
    ["SP-04.01.01", "P-04.01", 1, "Distribuição de Água aos Irrigantes",
     "Programação semanal, operação de comportas e bombas, e registro de volumes distribuídos.",
     "Entregar a água programada com eficiência e equidade.",
     "AI/GOM", "Marcos Vinícius",
     "Programação hídrica executada; Registros de volume", "SIG-Irrigação (fictício)",
     "img/diagramas/sp-04-01-01.svg"],
]
escreve(sp, subs, wrap_cols={5, 6, 9, 10}, center_cols={3})
sp.freeze_panes = "D2"

# ----------------------------------------------------------------------------
# ATIVIDADES
# ----------------------------------------------------------------------------
at = wb.create_sheet("Atividades")
cabecalho(at,
    ["Codigo", "Subprocesso", "Ordem", "Nome", "Descricao",
     "Responsavel_Ator", "Entradas", "Saidas", "Sistemas",
     "Prazo_Padrao", "Base_Normativa"],
    [16, 13, 7, 36, 46, 24, 38, 38, 24, 14, 30])
ativs = [
    ["A-06.01.01.01", "SP-06.01.01", 1, "Formalizar a necessidade (DFD)",
     "Registrar o Documento de Formalização da Demanda com justificativa, quantitativos e alinhamento ao PCA.",
     "Área demandante", "Necessidade identificada; Plano de Contratações Anual",
     "DFD assinado no SEI", "SEI", "5 dias úteis", "NI-027/2024 (fictícia)"],
    ["A-06.01.01.02", "SP-06.01.01", 2, "Levantar soluções de mercado",
     "Pesquisar soluções disponíveis, contratações similares e alternativas (inclusive não contratar).",
     "Equipe de planejamento da contratação", "DFD; Catálogos; Contratações similares (PNCP)",
     "Levantamento de soluções documentado", "PNCP; Compras.gov.br", "10 dias úteis",
     "Lei nº 14.133/2021, art. 18"],
    ["A-06.01.01.03", "SP-06.01.01", 3, "Estimar quantidades e resultados",
     "Definir quantitativos com memória de cálculo e resultados pretendidos com a contratação.",
     "Equipe de planejamento da contratação", "Levantamento de soluções; Séries históricas",
     "Memória de cálculo de quantitativos", "SEI", "5 dias úteis", ""],
    ["A-06.01.01.04", "SP-06.01.01", 4, "Elaborar e aprovar o ETP",
     "Consolidar o ETP no sistema e submeter à aprovação da autoridade competente.",
     "Equipe de planejamento; Autoridade competente", "Levantamentos e memórias anteriores",
     "ETP aprovado", "SEI; Compras.gov.br", "7 dias úteis", "Lei nº 14.133/2021, art. 18"],
    ["A-06.01.02.01", "SP-06.01.02", 1, "Redigir o Termo de Referência",
     "Elaborar o TR a partir do ETP, definindo objeto, requisitos, modelo de execução e critérios de medição.",
     "Equipe de planejamento da contratação", "ETP aprovado; Modelos padronizados",
     "Minuta de TR", "SEI", "10 dias úteis", "Lei nº 14.133/2021, art. 6º, XXIII"],
    ["A-06.01.02.02", "SP-06.01.02", 2, "Validar o TR com a Assessoria Jurídica",
     "Submeter a minuta ao órgão jurídico e ajustar conforme o parecer.",
     "AA/GLC; Assessoria Jurídica", "Minuta de TR",
     "TR validado; Parecer jurídico", "SEI", "15 dias úteis", "Lei nº 14.133/2021, art. 53"],
    ["A-06.01.03.01", "SP-06.01.03", 1, "Realizar a pesquisa de preços",
     "Consultar fontes admitidas, registrar parâmetros e tratar valores inexequíveis ou excessivos.",
     "Equipe de planejamento da contratação", "TR validado; Fontes de preços",
     "Relatório de pesquisa de preços", "Painel de Preços; PNCP", "10 dias úteis",
     "IN SEGES nº 65/2021"],
    ["A-06.01.03.02", "SP-06.01.03", 2, "Consolidar o valor estimado",
     "Aplicar o método definido (média/mediana/menor preço) e justificar a estimativa final.",
     "Equipe de planejamento da contratação", "Relatório de pesquisa",
     "Valor estimado consolidado", "SEI", "3 dias úteis", "IN SEGES nº 65/2021, art. 6º"],
    ["A-06.02.01.01", "SP-06.02.01", 1, "Publicar o edital",
     "Divulgar o edital no PNCP e no sistema de compras, observando prazos mínimos legais.",
     "Agente de contratação", "Edital aprovado; Parecer jurídico",
     "Edital publicado (PNCP)", "Compras.gov.br; PNCP", "2 dias úteis",
     "Lei nº 14.133/2021, art. 54"],
    ["A-06.02.01.02", "SP-06.02.01", 2, "Conduzir a sessão e julgar propostas",
     "Operar a fase de lances, julgar propostas, habilitar o vencedor e registrar a ata.",
     "Agente de contratação; Equipe de apoio", "Edital publicado; Propostas",
     "Ata da sessão; Resultado do julgamento", "Compras.gov.br", "Conforme edital",
     "Lei nº 14.133/2021"],
    ["A-04.01.01.01", "SP-04.01.01", 1, "Programar a distribuição hídrica",
     "Elaborar a programação semanal de distribuição conforme demanda dos lotes e disponibilidade hídrica.",
     "Equipe de operação do perímetro", "Plano de cultivo; Disponibilidade hídrica",
     "Programação semanal aprovada", "SIG-Irrigação (fictício)", "Semanal", ""],
    ["A-04.01.01.02", "SP-04.01.01", 2, "Operar e registrar a distribuição",
     "Operar comportas e bombas conforme programação e registrar volumes por ponto de entrega.",
     "Equipe de operação do perímetro", "Programação semanal",
     "Volumes registrados; Ocorrências", "SIG-Irrigação (fictício)", "Diário", ""],
]
escreve(at, ativs, wrap_cols={5, 7, 8, 9, 11}, center_cols={3})
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
    ["DOC-001", "Processo", "P-06.01", "Procedimento Operacional (POP)",
     "POP 06.01 — Planejamento da Contratação", "2.0", D(2026, 5, 18), "Vigente",
     "https://exemplo.codevasf.gov.br/repositorio/pop-06-01.pdf",
     "Publicado após validação do TO-BE."],
    ["DOC-002", "Processo", "P-06.01", "Diagrama BPMN",
     "Diagrama AS-IS — Planejamento da Contratação (Bizagi)", "1.0", D(2026, 3, 2),
     "Vigente", "img/diagramas/p-06-01.svg", "Exportado do Bizagi Modeler."],
    ["DOC-003", "Processo", "P-06.01", "Diagrama BPMN",
     "Diagrama TO-BE — Planejamento da Contratação (Bizagi)", "1.0", D(2026, 5, 4),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/p-06-01-tobe.pdf", ""],
    ["DOC-004", "Processo", "P-06.01", "Ata de reunião",
     "Ata — Reunião de contextualização com a GLC", "1.0", D(2026, 1, 12), "Vigente",
     "https://exemplo.codevasf.gov.br/sei/ata-contextualizacao-p0601", "SEI 59500.000123/2026-11."],
    ["DOC-005", "Processo", "P-06.01", "Relatório",
     "Relatório de oportunidades de melhoria (AS-IS → TO-BE)", "1.0", D(2026, 4, 13),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/rel-melhorias-p0601.pdf", ""],
    ["DOC-006", "Subprocesso", "SP-06.01.01", "Formulário/Modelo",
     "Modelo de DFD — Documento de Formalização da Demanda", "3.1", D(2026, 2, 10),
     "Vigente", "https://exemplo.codevasf.gov.br/modelos/dfd.docx", ""],
    ["DOC-007", "Subprocesso", "SP-06.01.03", "Procedimento Operacional (POP)",
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
     "Em elaboração", "img/diagramas/p-04-01.svg", "Modelagem em curso."],
    ["DOC-012", "Atividade", "A-06.01.01.01", "Formulário/Modelo",
     "Passo a passo do DFD no SEI", "1.0", D(2026, 2, 12), "Vigente",
     "https://exemplo.codevasf.gov.br/modelos/dfd-sei-passoapasso.pdf", ""],
    ["DOC-013", "Processo", "P-01.01", "Relatório",
     "Caderno de Indicadores Estratégicos 2026", "1.0", D(2026, 3, 31), "Vigente",
     "https://exemplo.codevasf.gov.br/estrategia/caderno-2026.pdf", ""],
    ["DOC-014", "Processo", "P-06.01", "Plano",
     "Plano de implantação do TO-BE (PMBOK — plano do projeto)", "1.0", D(2026, 5, 6),
     "Vigente", "https://exemplo.codevasf.gov.br/repositorio/plano-tobe-p0601.pdf", ""],
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
     "Designação formal de fiscais, agenda de medições e alertas no SEI.",
     "Ricardo Nogueira", "Aberto"],
    ["R-004", "Subprocesso", "SP-06.01.01",
     "Conhecimento concentrado em um único empregado (pessoa-chave) na elaboração de ETP.",
     "Pessoas", 4, 3, None, None, "Mitigar",
     "Publicar POP, treinar substitutos e revezar a equipe de planejamento.",
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
     "Monitoramento normativo mensal (resenha) e revisão programada dos POPs.",
     "Bruna Souza", "Em tratamento"],
    ["R-008", "Atividade", "A-06.01.03.01",
     "Uso de fontes de preço não admitidas pela IN SEGES nº 65/2021.",
     "Legal/Conformidade", 2, 3, None, None, "Mitigar",
     "Validação da pesquisa pela UNP antes da aprovação do valor estimado.",
     "Carlos Eduardo Lima", "Encerrado"],
    ["R-009", "Processo", "P-06.01",
     "Baixo engajamento das áreas demandantes no preenchimento do DFD.",
     "Pessoas", 3, 3, None, None, "Mitigar",
     "Oficinas de capacitação e modelo simplificado de DFD no SEI.",
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
     30, 26, None, "Mensal", "SEI", D(2026, 6, 30)],
    ["IND-003", "Processo", "P-06.02", "Impugnações por edital",
     "Nº médio de impugnações recebidas por edital publicado.", "nº", "Menor melhor",
     1, 0.8, None, "Trimestral", "Compras.gov.br", D(2026, 6, 30)],
    ["IND-004", "Processo", "P-06.03", "Contratos com fiscal designado",
     "% de contratos vigentes com fiscal formalmente designado.", "%", "Maior melhor",
     100, 92, None, "Mensal", "SEI", D(2026, 6, 30)],
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
     100, 88, None, "Mensal", "SEI", D(2026, 6, 30)],
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
di = wb.create_sheet("Diario_Mapeamento")
cabecalho(di,
    ["ID", "Processo", "Data", "Tipo", "Titulo", "Descricao", "Autor",
     "Participantes", "Entradas_Insumos", "Saidas_Entregaveis",
     "Evidencias", "Memoria"],
    [10, 10, 12, 12, 36, 50, 20, 36, 34, 36, 46, 44])
diario = [
    ["REG-001", "P-06.01", D(2026, 1, 5), "Marco", "Abertura do projeto de mapeamento",
     "Aprovação do termo de abertura do projeto de mapeamento do processo, com escopo, equipe e cronograma.",
     "Bruna Souza", "Bruna Souza; Ricardo Nogueira; Carlos Eduardo Lima",
     "Plano de ação 2026; Priorização da carteira de processos",
     "Termo de abertura aprovado (TAP)",
     "Termo de abertura|https://exemplo.codevasf.gov.br/sei/tap-p0601",
     "Escopo limitado ao planejamento da contratação; seleção e gestão contratual tratadas em projetos próprios."],
    ["REG-002", "P-06.01", D(2026, 1, 12), "Reunião", "Reunião de contextualização com a GLC",
     "Apresentação da metodologia, entendimento do contexto do processo e definição dos interlocutores.",
     "Carlos Eduardo Lima",
     "Equipe GLC; Ricardo Nogueira; Bruna Souza; Carlos Eduardo Lima",
     "Formulário de levantamento preenchido",
     "Ata da reunião; Lista preliminar de atividades",
     "Ata|https://exemplo.codevasf.gov.br/sei/ata-contextualizacao-p0601",
     "A área destacou gargalo na pesquisa de preços e retrabalho na aprovação do ETP."],
    ["REG-003", "P-06.01", D(2026, 1, 26), "Entrevista", "Entrevista com a equipe de ETP",
     "Detalhamento das atividades de elaboração do ETP, entradas, saídas e sistemas utilizados.",
     "Carlos Eduardo Lima", "Equipe de planejamento da contratação",
     "Roteiro de entrevista; Amostras de ETP",
     "Registro de entrevista; Insumos para o AS-IS",
     "Registro de entrevista|https://exemplo.codevasf.gov.br/sei/entrevista-etp",
     ""],
    ["REG-004", "P-06.01", D(2026, 2, 9), "Oficina", "Oficina de modelagem AS-IS",
     "Modelagem colaborativa do fluxo atual em BPMN no Bizagi Modeler, com validação visual dos participantes.",
     "Daniela Ribeiro", "Equipe GLC; UNP",
     "Insumos das entrevistas; Notação BPMN 2.0",
     "Diagrama BPMN AS-IS v0.1",
     "Diagrama AS-IS v0.1|img/diagramas/p-06-01.svg",
     "Identificados 3 retrabalhos e 2 handoffs desnecessários (princípios de desenho do CBOK)."],
    ["REG-005", "P-06.01", D(2026, 3, 2), "Validação", "Validação do AS-IS com o dono do processo",
     "Revisão e aprovação formal do diagrama AS-IS pelo dono do processo.",
     "Carlos Eduardo Lima", "Ricardo Nogueira; Equipe GLC; UNP",
     "Diagrama AS-IS v0.1", "Diagrama AS-IS v1.0 validado",
     "Diagrama AS-IS v1.0|img/diagramas/p-06-01.svg;Ata de validação|https://exemplo.codevasf.gov.br/sei/ata-validacao-asis",
     ""],
    ["REG-006", "P-06.01", D(2026, 3, 16), "Decisão", "Definição do escopo do TO-BE",
     "Decisão de incorporar o roteiro de pesquisa de preços e o modelo simplificado de DFD no redesenho.",
     "Bruna Souza", "Ricardo Nogueira; Bruna Souza",
     "Relatório preliminar de melhorias", "Escopo do TO-BE registrado",
     "Registro de decisão|https://exemplo.codevasf.gov.br/sei/decisao-tobe",
     "Priorizado o redesenho 'de fora para dentro', a partir da área demandante (cliente interno)."],
    ["REG-007", "P-06.01", D(2026, 4, 13), "Entrega", "TO-BE elaborado",
     "Entrega do diagrama TO-BE e do relatório de oportunidades de melhoria.",
     "Daniela Ribeiro", "UNP; Equipe GLC",
     "AS-IS validado; Boas práticas (CBOK cap. 6)",
     "Diagrama TO-BE v0.1; Relatório de melhorias",
     "Relatório de melhorias|https://exemplo.codevasf.gov.br/repositorio/rel-melhorias-p0601.pdf",
     ""],
    ["REG-008", "P-06.01", D(2026, 5, 4), "Validação", "Validação do TO-BE",
     "Aprovação do redesenho pelo dono do processo e pela gerência da área.",
     "Carlos Eduardo Lima", "Ricardo Nogueira; Gerência AA",
     "Diagrama TO-BE v0.1", "Diagrama TO-BE v1.0 validado",
     "Ata de validação TO-BE|https://exemplo.codevasf.gov.br/sei/ata-validacao-tobe",
     ""],
    ["REG-009", "P-06.01", D(2026, 5, 18), "Marco", "Publicação no repositório de processos",
     "Publicação do POP 06.01 e dos diagramas no repositório corporativo; encerramento do projeto.",
     "Bruna Souza", "UNP",
     "TO-BE validado; POP revisado",
     "POP 06.01 v2.0 publicado; Página do processo no painel",
     "POP 06.01|https://exemplo.codevasf.gov.br/repositorio/pop-06-01.pdf",
     "Lições aprendidas registradas para o próximo ciclo (PMBOK — encerramento)."],
    ["REG-010", "P-06.01", D(2026, 6, 15), "Reunião", "Monitoramento de indicadores (fase 5)",
     "Primeira reunião do ciclo de medição do processo redesenhado (CBOK — medir o sucesso).",
     "Gustavo Pereira", "UNP; GLC",
     "Dados de mai/2026 (SEI e Compras.gov.br)",
     "Ata; Ajuste da meta do IND-002",
     "Ata de monitoramento|https://exemplo.codevasf.gov.br/sei/ata-monitoramento-jun26",
     ""],
    ["REG-011", "P-06.02", D(2026, 6, 2), "Reunião", "Contextualização — Seleção do Fornecedor",
     "Início do mapeamento do processo de seleção, com apresentação do método e coleta do formulário.",
     "Carlos Eduardo Lima", "Equipe GLC; UNP",
     "Formulário de levantamento preenchido",
     "Ata; Cronograma do projeto",
     "Ata|https://exemplo.codevasf.gov.br/sei/ata-contextualizacao-p0602",
     ""],
    ["REG-012", "P-06.02", D(2026, 6, 30), "Oficina", "Oficina de modelagem AS-IS (sessão 1)",
     "Modelagem do fluxo da sessão pública e do julgamento de propostas.",
     "Daniela Ribeiro", "Agentes de contratação; UNP",
     "Insumos da contextualização", "Diagrama AS-IS parcial v0.2",
     "Diagrama parcial|img/diagramas/p-06-02.svg",
     "Fluxo de recursos administrativos ainda pendente de detalhamento com a área."],
    ["REG-013", "P-04.01", D(2026, 7, 10), "Entrevista", "Entrevista — operação do perímetro",
     "Levantamento das rotinas de programação e distribuição hídrica com a equipe de campo.",
     "Eduardo Martins", "Equipe de operação; UNP",
     "Roteiro de entrevista", "Registro de entrevista; Insumos para o AS-IS",
     "Registro|https://exemplo.codevasf.gov.br/sei/entrevista-p0401",
     ""],
]
escreve(di, diario, wrap_cols={5, 6, 8, 9, 10, 11, 12}, center_cols={4}, date_cols={3})
di.freeze_panes = "E2"
dv(di, "D", ref("Tipo_Registro"))

# ----------------------------------------------------------------------------
# TAREFAS (menor unidade de trabalho — CBOK 4.0)
# ----------------------------------------------------------------------------
tf = wb.create_sheet("Tarefas")
cabecalho(tf,
    ["Codigo", "Atividade", "Ordem", "Nome", "Descricao", "Tipo_Tarefa",
     "Responsavel", "Sistema", "Duracao_Estimada", "Observacoes"],
    [20, 16, 7, 38, 46, 16, 24, 22, 14, 28])
tarefas = [
    ["T-06.01.01.01.01", "A-06.01.01.01", 1, "Reunir informações da demanda",
     "Levantar justificativa, quantitativos preliminares e alinhamento ao PCA junto ao gestor da área.",
     "Manual", "Área demandante", "SEI", "0,5 dia", ""],
    ["T-06.01.01.01.02", "A-06.01.01.01", 2, "Preencher o formulário DFD no SEI",
     "Registrar o Documento de Formalização da Demanda no modelo padronizado.",
     "Manual", "Área demandante", "SEI", "0,5 dia", "Modelo DOC-006/DOC-012."],
    ["T-06.01.01.01.03", "A-06.01.01.01", 3, "Colher assinatura eletrônica do gestor",
     "Encaminhar o DFD para assinatura da autoridade competente da unidade.",
     "Manual", "Área demandante", "SEI", "1 dia", ""],
    ["T-06.01.03.01.01", "A-06.01.03.01", 1, "Consultar o Painel de Preços",
     "Pesquisar contratações públicas similares e extrair os relatórios de preços.",
     "Manual", "Equipe de planejamento", "Painel de Preços", "0,5 dia", ""],
    ["T-06.01.03.01.02", "A-06.01.03.01", 2, "Consultar contratações no PNCP",
     "Verificar atas e contratos vigentes de objetos equivalentes no PNCP.",
     "Manual", "Equipe de planejamento", "PNCP", "0,5 dia", ""],
    ["T-06.01.03.01.03", "A-06.01.03.01", 3, "Registrar cotações de fornecedores",
     "Solicitar e registrar cotações diretas quando as fontes oficiais forem insuficientes.",
     "Manual", "Equipe de planejamento", "SEI", "3 dias", "Mínimo de 3 fontes (IN 65/2021)."],
    ["T-06.01.03.01.04", "A-06.01.03.01", 4, "Aplicar tratamento estatístico",
     "Calcular média/mediana, excluir valores inexequíveis ou excessivos e justificar o método.",
     "Regra de negócio", "Equipe de planejamento", "Planilha padrão", "0,5 dia", ""],
    ["T-06.02.01.01.01", "A-06.02.01.01", 1, "Cadastrar o edital no Compras.gov.br",
     "Inserir o edital aprovado, anexos e cronograma do certame no sistema.",
     "Manual", "Agente de contratação", "Compras.gov.br", "0,5 dia", ""],
    ["T-06.02.01.01.02", "A-06.02.01.01", 2, "Publicar o aviso no PNCP",
     "Divulgação automática do aviso de licitação a partir do cadastro no sistema.",
     "Automatizada", "Compras.gov.br", "PNCP", "Imediato", ""],
    ["T-04.01.01.01.01", "A-04.01.01.01", 1, "Consolidar demandas semanais dos lotes",
     "Compilar os pedidos de água dos irrigantes por setor hidráulico para a programação.",
     "Manual", "Equipe de operação", "SIG-Irrigação (fictício)", "1 dia", ""],
]
escreve(tf, tarefas, wrap_cols={5, 10}, center_cols={3, 6, 9})
tf.freeze_panes = "D2"
dv(tf, "F", ref("Tipo_Tarefa"))

# ----------------------------------------------------------------------------
# JORNADA · REPOSITORIO · NUGEP · GLOSSARIO · FAQ · PARAMETROS
# (conteúdo em scripts/dados_conteudo.py — edite depois direto na planilha)
# ----------------------------------------------------------------------------
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
cabecalho(ng, ["Ordem", "Nome", "Papel", "Unidade_Sigla", "Unidade_Nome", "Email", "Telefone"],
          [7, 24, 34, 14, 38, 32, 16])
escreve(ng, [list(x) for x in CONTEUDO.NUGEP], wrap_cols={3, 5}, center_cols={1})
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
      "Processo de trabalho (aba Subprocessos) → Atividade → Tarefa. Os relacionamentos usam os CÓDIGOS: "
      "Processos→Macroprocesso, Subprocessos→Processo, Atividades→Subprocesso, Tarefas→Atividade; "
      "Documentos/Riscos/Indicadores usam Vinculo_Nivel + Vinculo_Codigo; o Diário usa o código do Processo.")

titulo(11, "Convenções de preenchimento")
linha(12, "Listas na célula", "Separe múltiplos itens com ponto e vírgula ( ; ). "
      "Ex.: 'SEI; Compras.gov.br; PNCP'.")
linha(13, "Evidências (Diário)", "Formato Nome|URL, separando várias com ';'. "
      "Ex.: 'Ata|https://...;Diagrama|img/diagramas/x.svg'.")
linha(14, "Percentual", "Na aba Processos, use percentual (0% a 100%).")
linha(15, "Datas", "Formato dd/mm/aaaa.")
linha(16, "Imagens Bizagi", "Recomendado: publique a imagem exportada do Bizagi on-line "
      "(Drive público, intranet acessível etc.) e cole a URL na coluna Imagem_Bizagi — o painel exibe a "
      "imagem e o clique abre o link original. Links de compartilhamento do Google Drive são convertidos "
      "automaticamente para exibição. Caminhos relativos (img/diagramas/...) também funcionam.")
linha(17, "Células de fórmula", "Colunas com fundo cinza (Riscos: Nivel_PxI e Classificacao; "
      "Indicadores: Situacao) são CALCULADAS — não digite valores; ao inserir linhas, copie a fórmula da linha acima.")
linha(18, "Validação de dados", "Campos com lista suspensa buscam os valores na aba 'Listas' "
      "(edite lá para incluir novas opções).")

titulo(20, "Dicionário de abas")
abas_desc = [
    ("Macroprocessos", "1º nível da cadeia de valor (CBOK: processos primários/finalísticos, de suporte e gerenciais)."),
    ("Processos", "2º nível — processos de negócio, com ficha completa: SIPOC (fornecedores, entradas, saídas, clientes), status e marcos do mapeamento (M1–M9), fase do ciclo BPM e dados do projeto."),
    ("Subprocessos", "3º nível — processos de trabalho, vinculados ao processo de negócio."),
    ("Atividades", "4º nível (CBOK), com entradas, saídas, ator, sistemas e prazos."),
    ("Tarefas", "5º e último nível (CBOK): menor unidade de trabalho de uma atividade — manual, automatizada ou regra de negócio."),
    ("Documentos", "Repositório: POPs, manuais, atas, diagramas BPMN (Bizagi), relatórios — vinculados a qualquer nível."),
    ("Riscos", "Riscos vinculados a qualquer nível; nível = Probabilidade × Impacto (matriz 5×5)."),
    ("Indicadores", "Indicadores de desempenho por nível, com meta, resultado e situação calculada."),
    ("Diario_Mapeamento", "Registro rastreável do trabalho: reuniões, oficinas, entrevistas, decisões, entregas e marcos, com entradas, saídas/entregáveis e evidências (CBOK 4.0; PMBOK)."),
    ("Jornada", "Etapas da jornada de mapeamento (Descobrir → Evoluir), exibidas na aba Repositório do painel."),
    ("Repositorio", "Materiais e ferramentas: metodologia e guia oficiais (RES 031/2025), templates, instrumentos por fase do ciclo BPM, ferramentas e referências."),
    ("NUGEP", "Integrantes do Núcleo de Gestão Normativa e de Processos (aba NUGEP do painel)."),
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
    ("Registros do diário", "=COUNTA(Diario_Mapeamento!$A$2:$A$500)"),
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
               "Tarefas", "Documentos", "Riscos", "Indicadores", "Diario_Mapeamento", "Jornada",
               "Repositorio", "NUGEP", "Glossario", "FAQ", "Parametros", "Listas"]
wb._sheets = [wb[n] for n in ordem_final]
wb.active = 0
wb.save("/home/claude/painel-processos/data/painel-processos-dados.xlsx")
print("Planilha gerada com sucesso.")
