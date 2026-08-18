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

Revisão desta versão (auditoria de forma e conteúdo, pedido do usuário):
  - Códigos por nível, sempre com 2 letras + número sequencial DENTRO do
    nível (não mais compostos com o número do pai): MG-/MF-/MS- (macro,
    por categoria), PP- (processo), SP- (subprocesso), AT- (atividade),
    TR- (tarefa). Quem liga um nível a outro são as colunas
    Macroprocesso/Vinculo_Pai/Atividade/Vinculo_Codigo — nunca o texto do
    código. Isso é o que deixa a planilha pronta para virar tabelas SQL.
  - Nomenclatura padronizada: Unidade_Organica_Responsavel (era
    Area_Responsavel/Unidade_Responsavel) e Unidades_Organicas_Corresponsaveis
    (era Unidades_Corresponsaveis) em todas as abas que têm essas colunas;
    Beneficiarios (era Clientes_Beneficiarios em Macroprocessos);
    Ponto_Focal_Nugep (era Interlocutor); Produto_Principal (era Produto).
  - Colunas removidas por não terem uso real ou terem virado outra coisa:
    Dono_Processo (Macroprocessos/Processos) e Dono (Subprocessos) — vira
    registro na aba Papeis; Fase_Ciclo_BPM (Processos) — não aparecia no
    painel; Normativos_Aplicaveis/Normativos_Relacionados — cada normativo
    agora é uma linha própria em Documentos, com Link; Cronograma_Proposto_Dias
    (Subprocessos) e Responsavel/Imagem_Bizagi (Atividades) e Imagem_Bizagi
    (Tarefas) — pedido do usuário; Categoria/Termos_Relacionados (Glossario)
    e Codigo (Repositorio) — pedido do usuário; aba PlanoAcao inteira — o
    5W2H passa a ser tratado dentro do marco de análise do processo, sem
    aba própria no painel nem na planilha.
  - Processos.Percentual e todo Vinculo_Nivel viram fórmula (não digite).
  - Marcos M1–M10 renomeados para bater 1:1 com a Metodologia (RES 031/2025).
  - Colunas novas: Processos.Maturidade (lista fechada), Riscos
    .Data_Identificacao/.Prazo_Tratamento (lacunas do registro de risco).
"""

ESQUEMA = {
    "Macroprocessos": [
        ("Codigo", 10), ("Nome", 34), ("Categoria", 13), ("Ordem", 7),
        ("Unidade_Organica_Responsavel", 16), ("Unidades_Organicas_Corresponsaveis", 30),
        ("Descricao", 46), ("Objetivo", 46), ("Entregas", 40), ("Beneficiarios", 34),
        ("Partes_Interessadas", 34), ("Sistemas", 30), ("Imagem_Bizagi", 30), ("Observacoes", 26),
    ],
    "Processos": [
        ("Codigo", 9), ("Nome", 34), ("Macroprocesso", 13), ("Descricao", 44), ("Objetivo", 40),
        ("Unidade_Organica_Responsavel", 16), ("Ponto_Focal_Nugep", 24), ("Unidades_Organicas_Corresponsaveis", 30),
        ("Prioridade", 10), ("Complexidade", 12), ("Maturidade", 13), ("Status_Mapeamento", 14), ("Percentual", 10),
        ("Inicio_Mapeamento", 13), ("Prazo_Previsto", 13), ("Data_Conclusao", 13), ("Ultima_Atualizacao", 13),
        ("M1_Conhecer_Processo", 9), ("M2_Processo_Modelado", 9), ("M3_Subprocessos_Modelados", 9),
        ("M4_ASIS_Modelado", 9), ("M5_ASIS_Validado", 9), ("M6_Procedimento_Aprovado", 9),
        ("M7_Processo_Publicado", 9), ("M8_TOBE_Elaborado", 9), ("M9_TOBE_Aprovado", 9),
        ("M10_Processo_Transformado", 9),
        ("Fornecedores", 30), ("Entradas", 34), ("Saidas", 34), ("Beneficiarios", 28), ("Sistemas", 28),
        ("Processo_ECodevasf", 20), ("Imagem_Bizagi", 30), ("Competencias_Necessarias", 40), ("Fontes_Dados", 30),
        ("Proxima_Acao", 28), ("Pendencia", 24),
    ],
    "Subprocessos": [
        ("Codigo", 9), ("Nome", 32), ("Vinculo_Pai", 12), ("Ordem", 7), ("Descricao", 44), ("Objetivo", 38),
        ("Unidade_Organica_Responsavel", 16), ("Unidades_Organicas_Corresponsaveis", 28),
        ("Reutilizavel", 12), ("Reutilizado_Em", 24), ("Produto_Principal", 34),
        ("Entradas", 36), ("Saidas", 36), ("Sistemas", 26), ("Fontes_Dados", 26), ("Imagem_Bizagi", 28),
    ],
    "Atividades": [
        ("Codigo", 9), ("Nome", 34), ("Vinculo_Pai", 11), ("Ordem", 7), ("Tipo_Atividade", 17),
        ("Descricao", 44), ("Entradas", 34), ("Saidas", 34), ("Sistemas", 22),
    ],
    "Tarefas": [
        ("Codigo", 9), ("Nome", 34), ("Atividade", 9), ("Ordem", 7), ("Tipo_Tarefa", 15),
        ("Disparador", 32), ("Passos", 44), ("Principios", 28), ("Criterios_Desempenho", 32),
        ("Resultados_Esperados", 30), ("Materiais_Ferramentas", 26), ("Pessoas_Consultar", 26),
        ("Duracao_Estimada", 13), ("Observacoes", 26), ("Unidades_Organicas_Corresponsaveis", 28),
    ],
    "Documentos": [
        ("ID", 9), ("Titulo", 48), ("Tipo_Documento", 22), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16),
        ("Versao", 8), ("Data", 12), ("Situacao", 14), ("Ato_Aprovacao", 30), ("Link", 44), ("Observacoes", 28),
    ],
    "Riscos": [
        ("ID", 9), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16), ("Descricao_Risco", 46), ("Categoria", 20),
        ("Fatores", 34), ("Cronograma_Risco", 24), ("Data_Identificacao", 13),
        ("Probabilidade_1a5", 9), ("Impacto_1a5", 8), ("Nivel_PxI", 8), ("Classificacao", 12),
        ("Resposta", 11), ("Controles_Tratamento", 40), ("Prazo_Tratamento", 13),
        ("Responsavel", 20), ("Status", 13),
    ],
    "Metricas": [
        ("ID", 9), ("Nome", 38), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16), ("Categoria", 18),
        ("Descricao_Formula", 44), ("Unidade", 9), ("Polaridade", 13), ("Meta", 8), ("Periodicidade", 13),
        ("Fonte", 22), ("Observacoes", 26),
    ],
    "Medicoes": [
        ("ID", 9), ("Metrica_ID", 10), ("Data_Medicao", 13), ("Valor", 10), ("Observacao", 28),
    ],
    "Papeis": [
        ("ID", 9), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16), ("Papel", 30),
        ("Envolvimento", 16), ("Unidade_Pessoa", 22),
    ],
    "Regras": [
        ("ID", 9), ("Nome", 34), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16),
        ("Tipo_Regra", 16), ("Descricao", 46), ("Fonte_Normativa", 24),
    ],
    "Cultura_Processos": [
        ("ID", 9), ("Ordem", 7), ("Caracteristica", 56), ("Situacao", 14), ("Observacao", 28),
    ],
    "Iniciativas": [
        ("ID", 9), ("Data", 12), ("Titulo", 42), ("Tipo", 22), ("Descricao", 52), ("Processos_Relacionados", 22),
    ],
    "Competencias": [
        ("Ordem", 7), ("Instancia", 46), ("Item_Normativo", 14), ("Atribuicoes", 90),
    ],
    "Jornada": [
        ("Ordem", 7), ("Fase", 13), ("Nome", 30), ("Duracao", 12), ("Objetivo", 44),
        ("Atividades_Chave", 60), ("Quem_Faz", 30), ("Entregaveis", 40), ("Sentimento_Usuario", 40),
    ],
    "Repositorio": [
        ("ID", 9), ("Categoria", 18), ("Fase_Ciclo", 15), ("Titulo", 42), ("Descricao", 56),
        ("Fonte", 12), ("Link", 50), ("Ordem", 7),
    ],
    "NUGEP": [
        ("Ordem", 7), ("Nome", 24), ("Papel", 34), ("Unidade_Sigla", 14), ("Unidade_Nome", 38),
        ("Email", 32), ("Telefone", 16), ("Foto", 46), ("Hierarquia", 11),
    ],
    "Glossario": [("Termo", 34), ("Definicao", 84), ("Fonte", 20)],
    "FAQ": [("Ordem", 7), ("Categoria", 28), ("Pergunta", 50), ("Resposta", 90)],
    "Siglas": [("Sigla", 20), ("Nome", 60)],
    "Parametros": [("Chave", 24), ("Valor", 90)],
    "Dicionario": [("Aba", 18), ("Coluna", 32), ("Descricao", 62), ("Exemplo", 46)],
}

# Colunas que a planilha calcula sozinha (fórmula na célula, fundo cinza).
# Não recebem dado de demonstração nem devem ser preenchidas à mão.
CALCULADAS = {
    "Processos": ["Percentual"],
    "Documentos": ["Vinculo_Nivel"],
    "Riscos": ["Vinculo_Nivel", "Nivel_PxI", "Classificacao"],
    "Metricas": ["Vinculo_Nivel"],
    "Papeis": ["Vinculo_Nivel"],
    "Regras": ["Vinculo_Nivel"],
    "NUGEP": ["Unidade_Nome"],
}

# Listas suspensas (aba Listas) — nome da coluna → opções fechadas.
LISTAS_OPCOES = {
    "Categoria_Macroprocesso": ["Gerencial", "Finalístico", "Suporte"],
    "Status_Mapeamento": ["Não iniciado", "Em andamento", "Concluído", "Suspenso"],
    "Prioridade": ["Alta", "Média", "Baixa"],
    "Complexidade": ["Alta", "Média", "Baixa"],
    "Maturidade_Processo": ["Inicial", "Repetível", "Definido", "Gerenciado", "Otimizado"],
    "Sim_Nao": ["Sim", "Não", "Não se aplica"],
    "Nivel_Vinculo": ["Macroprocesso", "Processo", "Subprocesso", "Atividade", "Tarefa"],
    "Tipo_Atividade": ["Agregação de Valor", "Transferência", "Controle"],
    "Tipo_Tarefa": ["Manual", "Automatizada", "Regra de negócio"],
    "Tipo_Documento": ["Procedimento (PRO)", "Manual", "Norma interna", "Legislação externa",
                       "Formulário/Modelo", "Ata de reunião", "Diagrama BPMN", "Checklist",
                       "Relatório", "Plano", "Outro"],
    "Situacao_Documento": ["Vigente", "Em elaboração", "Em revisão", "Revogado"],
    "Categoria_Risco": ["Operacional", "Legal/Conformidade", "Pessoas",
                        "Tecnologia da Informação", "Financeiro/Orçamentário", "Imagem/Reputação"],
    "Resposta_Risco": ["Mitigar", "Aceitar", "Transferir", "Evitar"],
    "Status_Risco": ["Aberto", "Em tratamento", "Encerrado"],
    "Envolvimento_RACI": ["Executa (R)", "Aprova (A)", "Consultado (C)", "Informado (I)"],
    "Tipo_Regra": ["Operação", "Regulatória", "Cálculo", "Restrição"],
    "Polaridade": ["Maior melhor", "Menor melhor"],
    "Periodicidade": ["Mensal", "Bimestral", "Trimestral", "Semestral", "Anual"],
    "Situacao_Cultura": ["Sim", "Não", "Não avaliado"],
    "Tipo_Iniciativa": ["Automação", "Reaproveitamento de modelo", "Decisão de governança", "Outro"],
    "Categoria_Repositorio": ["Documento oficial", "Template", "Instrumento", "Ferramenta", "Referência"],
    "Fase_Instrumento": ["Planejamento", "Análise", "Desenho", "Implementação", "Monitoramento", "Refinamento"],
    "Categoria_FAQ": ["Conceitos básicos", "Modelagem e SIPOC", "Cadeia de Valor e governança",
                      "Indicadores, metas e riscos", "Plano de Ações AE/GPE", "Como usar o painel"],
}

# Prefixo de código por nível — 2 letras + número sequencial DENTRO do nível.
PREFIXO_NIVEL = {
    "Macroprocesso": ("MG", "MF", "MS"),  # Gerencial / Finalístico / Suporte
    "Processo": "PP", "Subprocesso": "SP", "Atividade": "AT", "Tarefa": "TR",
}


def colunas(aba):
    return [c for c, _ in ESQUEMA.get(aba, [])]


def largura(aba, coluna):
    for c, w in ESQUEMA.get(aba, []):
        if c == coluna:
            return w
    return 22


def conferir(headers_por_aba):
    """Compara um {aba: [colunas]} com o esquema e devolve a diferença."""
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
