# -*- coding: utf-8 -*-
"""
Esquema canônico da planilha do Painel de Processos.

Uma única declaração do que o painel espera encontrar em cada aba: nome de
coluna, largura e um valor de demonstração para quando a célula está vazia.
`atualizar_planilha.py` usa isto para acrescentar à planilha EXISTENTE as
colunas que faltam, sem recriar nada.

Como manter em dia: os nomes aqui têm de ser os mesmos das chamadas de
`cabecalho(...)` em `gerar_planilha.py` (que continua sendo o gerador da
planilha de exemplo). `conferir()` compara os dois e aponta a diferença.

Os valores em DEMO são FICTÍCIOS, para o painel não abrir com campo vazio
enquanto os dados reais não chegam — troque à vontade.
"""

# aba → [(coluna, largura), ...]  na ordem em que devem aparecer
ESQUEMA = {
    "Macroprocessos": [
        ("Codigo", 9), ("Nome", 34), ("Categoria", 12), ("Ordem", 7),
        ("Descricao", 46), ("Objetivo", 46), ("Unidade_Responsavel", 16),
        ("Dono_Processo", 26), ("Entregas", 40), ("Clientes_Beneficiarios", 34),
        ("Partes_Interessadas", 34), ("Sistemas", 30),
        ("Normativos_Aplicaveis", 40), ("Imagem_Bizagi", 30),
        ("Observacoes", 26), ("Unidades_Corresponsaveis", 30),
    ],
    "Processos": [
        ("Codigo", 10), ("Macroprocesso", 14), ("Nome", 34), ("Descricao", 44),
        ("Objetivo", 40), ("Area_Responsavel", 16), ("Dono_Processo", 24),
        ("Interlocutor", 24), ("Prioridade", 10), ("Complexidade", 12),
        ("Status_Mapeamento", 14), ("Percentual", 10), ("Fase_Ciclo_BPM", 30),
        ("Inicio_Mapeamento", 13), ("Prazo_Previsto", 13), ("Data_Conclusao", 13),
        ("Fornecedores", 30), ("Entradas", 34), ("Saidas", 34),
        ("Beneficiarios", 28), ("Sistemas", 28), ("Normativos_Relacionados", 40),
        ("Processo_ECodevasf", 20), ("Imagem_Bizagi", 30),
        ("M1_Reuniao_Contextualizacao", 9), ("M2_Macro_Processo_Modelados", 9),
        ("M3_Subprocessos_Modelados", 9), ("M4_ASIS_Modelado", 9),
        ("M5_ASIS_Validado", 9), ("M6_Procedimento_Validado", 9),
        ("M7_Procedimento_Aprovado", 9), ("M8_TOBE_Elaborado", 9),
        ("M9_TOBE_Validado", 9), ("M10_Publicado_Repositorio", 34),
        ("Proxima_Acao", 28), ("Pendencia", 13), ("Ultima_Atualizacao", 13),
        ("Unidades_Corresponsaveis", 30),
    ],
    # Reutilizavel/Reutilizado_Em modelam o "Subprocesso Reutilizável" do
    # BPMN 2.0 (Call Activity) — no Bizagi, que o painel usa para os
    # diagramas, o próprio elemento já se chama assim. Reutilizavel = Sim
    # marca que ESTE subprocesso (que mora nativamente sob Vinculo_Pai) é
    # chamado também de outros pontos do portfólio; Reutilizado_Em lista os
    # códigos desses outros Processos/Subprocessos (';'), de qualquer
    # macroprocesso e em qualquer nível de aninhamento.
    "Subprocessos": [
        ("Codigo", 16), ("Vinculo_Pai", 16), ("Ordem", 7), ("Nome", 34),
        ("Descricao", 46), ("Objetivo", 40), ("Unidade_Responsavel", 16),
        ("Dono", 24), ("Entradas", 40), ("Saidas", 40), ("Sistemas", 28),
        ("Imagem_Bizagi", 30), ("Unidades_Corresponsaveis", 30),
        ("Reutilizavel", 13), ("Reutilizado_Em", 34),
    ],
    # Prazo_Padrao saiu: a duração de uma atividade não é mais digitada, é o
    # somatório de Duracao_Estimada das suas tarefas (painel calcula em JS).
    "Atividades": [
        ("Codigo", 16), ("Vinculo_Pai", 13), ("Ordem", 7), ("Nome", 36),
        ("Descricao", 46), ("Responsavel_Ator", 24), ("Entradas", 38),
        ("Saidas", 38), ("Sistemas", 24),
        ("Base_Normativa", 30), ("Imagem_Bizagi", 44), ("Executor", 30),
        ("Unidades_Corresponsaveis", 30),
    ],
    # Duracao_Estimada é NÚMERO em horas úteis (1 dia útil = 8h) — não mais
    # texto livre. Some-se por atividade, depois recursivamente por
    # subprocesso/processo (caminho crítico, não caminho feliz).
    "Tarefas": [
        ("Codigo", 20), ("Atividade", 16), ("Ordem", 7), ("Nome", 38),
        ("Descricao", 46), ("Tipo_Tarefa", 16), ("Responsavel", 24),
        ("Sistema", 22), ("Duracao_Estimada", 14), ("Observacoes", 28),
        ("Imagem_Bizagi", 44), ("Unidades_Corresponsaveis", 30),
    ],
    "Documentos": [
        ("ID", 10), ("Vinculo_Nivel", 15), ("Vinculo_Codigo", 14),
        ("Tipo_Documento", 26), ("Titulo", 52), ("Versao", 8), ("Data", 12),
        ("Situacao", 14), ("Link", 44), ("Observacoes", 30),
    ],
    "Riscos": [
        ("ID", 9), ("Vinculo_Nivel", 15), ("Vinculo_Codigo", 14),
        ("Descricao_Risco", 52), ("Categoria", 22), ("Probabilidade_1a5", 12),
        ("Impacto_1a5", 10), ("Nivel_PxI", 9), ("Classificacao", 13),
        ("Resposta", 11), ("Controles_Tratamento", 46), ("Responsavel", 22),
        ("Status", 14),
    ],
    "Indicadores": [
        ("ID", 10), ("Vinculo_Nivel", 15), ("Vinculo_Codigo", 14), ("Nome", 40),
        ("Descricao_Formula", 46), ("Unidade", 9), ("Polaridade", 14),
        ("Meta", 8), ("Resultado_Atual", 12), ("Situacao", 15),
        ("Periodicidade", 13), ("Fonte", 24), ("Ultima_Medicao", 13),
    ],
    "Jornada": [
        ("Ordem", 7), ("Fase", 13), ("Nome", 30), ("Duracao", 12),
        ("Objetivo", 44), ("Atividades_Chave", 60), ("Quem_Faz", 30),
        ("Entregaveis", 40), ("Sentimento_Usuario", 40),
    ],
    "Repositorio": [
        ("ID", 10), ("Categoria", 18), ("Fase_Ciclo", 15), ("Codigo", 12),
        ("Titulo", 42), ("Descricao", 56), ("Fonte", 12), ("Link", 50),
        ("Ordem", 7),
    ],
    "NUGEP": [
        ("Ordem", 7), ("Nome", 24), ("Papel", 34), ("Unidade_Sigla", 14),
        ("Unidade_Nome", 38), ("Email", 32), ("Telefone", 16), ("Foto", 46),
        ("Hierarquia", 11),
    ],
    "Glossario": [
        ("Termo", 34), ("Categoria", 22), ("Definicao", 80), ("Fonte", 18),
        ("Termos_Relacionados", 40),
    ],
    "FAQ": [
        ("Ordem", 7), ("Categoria", 26), ("Pergunta", 50), ("Resposta", 90),
    ],
    "Siglas": [("Sigla", 20), ("Nome", 60)],
    "Parametros": [("Chave", 24), ("Valor", 90)],
}

# Colunas que a planilha calcula sozinha (fórmula na célula, fundo cinza).
# Não recebem dado de demonstração nem devem ser preenchidas à mão: cada
# uma repetia, digitada, algo que já estava em outra coluna.
CALCULADAS = {
    "Processos": ["Macroprocesso", "Percentual"],
    "Documentos": ["Vinculo_Nivel"],
    "Riscos": ["Vinculo_Nivel", "Nivel_PxI", "Classificacao"],
    "Indicadores": ["Vinculo_Nivel", "Situacao"],
    "NUGEP": ["Unidade_Nome"],
}


# ── Valores de demonstração ────────────────────────────────────────────
# Preenchem apenas células VAZIAS de colunas recém-criadas, para o painel
# não abrir com buraco enquanto os dados reais não chegam.
#
# Um valor pode ser:
#   str            → o mesmo texto em todas as linhas
#   list           → percorrido em rodízio, linha a linha
#   callable(i, ws)→ recebe o número da linha e a aba
_RETRATOS = [
    "https://randomuser.me/api/portraits/women/44.jpg",
    "https://randomuser.me/api/portraits/men/45.jpg",
    "https://randomuser.me/api/portraits/women/12.jpg",
    "https://randomuser.me/api/portraits/men/76.jpg",
    "https://randomuser.me/api/portraits/women/29.jpg",
    "https://randomuser.me/api/portraits/men/22.jpg",
    "https://randomuser.me/api/portraits/women/8.jpg",
    "https://randomuser.me/api/portraits/men/54.jpg",
    "https://randomuser.me/api/portraits/women/90.jpg",
    "https://randomuser.me/api/portraits/men/60.jpg",
]
_UNIDADES = ["AE/GPE/UNP", "AE/GPE", "AT/GTI", "AR/GPR", "AE/GAG",
             "AA/GLC", "AG/GGP", "AI/GOM", "AR/GRB", "AR/GDT"]

DEMO = {
    "NUGEP": {
        "Papel": ["Analista de Processos", "Especialista em BPM",
                  "Interlocutor de Processos", "Apoio Metodológico",
                  "Analista de Riscos e Conformidade"],
        "Unidade_Sigla": _UNIDADES,
        "Unidade_Nome": "Unidade de Gestão Normativa e de Processos",
        "Foto": _RETRATOS,
        # 3 = equipe da Unidade. Os níveis 1 (Gerente-Executivo) e 2
        # (Gerente) entram por linha própria, não por preenchimento.
        "Hierarquia": 3,
    },
    "Macroprocessos": {
        "Unidades_Corresponsaveis": _UNIDADES,
        "Normativos_Aplicaveis": "A definir",
        "Observacoes": "",
    },
    "Processos": {
        "Unidades_Corresponsaveis": _UNIDADES,
        "Proxima_Acao": "Agendar reunião de contextualização com a área.",
        "Pendencia": "",
        "Normativos_Relacionados": "A definir",
    },
    "Subprocessos": {"Unidades_Corresponsaveis": _UNIDADES},
    "Atividades": {
        "Unidades_Corresponsaveis": _UNIDADES,
        "Executor": ["Analista em Desenvolvimento Regional",
                     "Técnico em Desenvolvimento Regional",
                     "Assistente em Desenvolvimento Regional"],
        "Base_Normativa": "A definir",
    },
    "Tarefas": {
        "Unidades_Corresponsaveis": _UNIDADES,
        "Tipo_Tarefa": "Manual",
    },
    "Documentos": {"Observacoes": ""},
    "Riscos": {"Responsavel": "A definir", "Status": "Aberto"},
    "Indicadores": {"Fonte": "A definir", "Periodicidade": "Mensal"},
    "Repositorio": {"Fonte": "Codevasf"},
}

# Linhas que o painel espera encontrar e que talvez não existam ainda.
# Casadas pela PRIMEIRA coluna da aba; nada é sobrescrito.
LINHAS_NOVAS = {
    # O bloco "Contato institucional" da aba NUGEP mostra a hierarquia
    # acima da Unidade. Sem estas duas linhas, ele aparece só com a equipe.
    "NUGEP": [
        {"Ordem": 1, "Nome": "Ricardo Nunes Vasconcelos",
         "Papel": "Gerente-Executivo da Área de Gestão Estratégica",
         "Unidade_Sigla": "AE", "Unidade_Nome": "Área de Gestão Estratégica",
         "Email": "ricardo.vasconcelos@codevasf.gov.br",
         "Telefone": "(61) 2028-4400",
         "Foto": "https://randomuser.me/api/portraits/men/32.jpg",
         "Hierarquia": 1},
        {"Ordem": 2, "Nome": "Patrícia Moreira Lopes",
         "Papel": "Gerente de Planejamento Estratégico",
         "Unidade_Sigla": "AE/GPE",
         "Unidade_Nome": "Gerência de Planejamento Estratégico",
         "Email": "patricia.lopes@codevasf.gov.br",
         "Telefone": "(61) 2028-4430",
         "Foto": "https://randomuser.me/api/portraits/women/68.jpg",
         "Hierarquia": 2},
    ],
}

# Itens que precisam existir nas listas suspensas da aba Listas.
ITENS_LISTAS = {
    "Sim_Nao": ["Não se aplica"],
}


def colunas(aba):
    return [c for c, _ in ESQUEMA.get(aba, [])]


def largura(aba, coluna):
    for c, w in ESQUEMA.get(aba, []):
        if c == coluna:
            return w
    return 22


def valor_demo(aba, coluna, i, ws=None):
    """Valor de demonstração para (aba, coluna) na i-ésima linha de dados."""
    if coluna in CALCULADAS.get(aba, ()):
        return None
    v = DEMO.get(aba, {}).get(coluna)
    if v is None:
        return None
    if callable(v):
        return v(i, ws)
    if isinstance(v, (list, tuple)):
        return v[i % len(v)] if v else None
    return v


def conferir(headers_por_aba):
    """Compara um {aba: [colunas]} com o esquema e devolve a diferença.
    Serve para checar o gerador contra este arquivo."""
    problemas = []
    for aba, cols in ESQUEMA.items():
        esperado = [c for c, _ in cols]
        atual = headers_por_aba.get(aba)
        if atual is None:
            problemas.append(f"{aba}: aba ausente")
            continue
        faltam = [c for c in esperado if c not in atual]
        sobram = [c for c in atual if c not in esperado]
        if faltam:
            problemas.append(f"{aba}: faltam {', '.join(faltam)}")
        if sobram:
            problemas.append(f"{aba}: fora do esquema {', '.join(sobram)}")
    return problemas
