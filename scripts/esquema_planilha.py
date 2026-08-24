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

Revisão mais recente (rastreabilidade e volume real de itens, pedido do
usuário):
  - **Códigos (PP-/SP-/AT-/TR-) reiniciam em 01 a cada novo vínculo**
    (Macroprocesso→Processo, Processo→Subprocesso, Subprocesso→Subprocesso
    quando aninhado, Atividade→Tarefa) — reflete que a Companhia terá muitos
    "PP-01", um por Macroprocesso, e não caberia mais numerar tudo de forma
    globalmente única. Isso significa que o Código SOZINHO deixou de ser
    chave única: a chave real é (pai, Código) — ou, de forma prática, a
    coluna Trilha.
  - **Trilha vira a 1ª coluna** de Processos/Subprocessos/Atividades/Tarefas
    (antes vinha por último), seguida das colunas de rollup calculadas
    **Macroprocesso/Processo/Subprocesso/Atividade** (conforme o nível) —
    todas ANTES do Código, para a hierarquia aparecer de cara ao abrir a
    aba. Como o Código não é mais único, essas colunas não são um VLOOKUP
    vivo por código (ficaria ambíguo: existem várias linhas com o mesmo
    código) — são calculadas no momento em que a planilha é gerada a partir
    da posição real de cada item; reordenar ou reclassificar um item exige
    rodar o gerador de novo.
  - Qualquer referência cruzada a um código que deixou de ser único
    (`Vinculo_Pai`, `Atividade`, `Vinculo_Codigo`, `Processos_Relacionados`,
    `Reutilizado_Em`) precisa do CAMINHO COMPLETO quando há ambiguidade —
    ex. `Reutilizado_Em = "MS-01/PP-02; MF-02/PP-01"`, não `"SP-01"` sozinho.
  - **Marco "Em andamento"**: M1–M10 passam a aceitar Sim / Em andamento /
    Não / Não se aplica (lista `Status_Marco`, não mais `Sim_Nao`). Em
    andamento vale meia nota no `Processos.Percentual` (Sim=1, Em
    andamento=0,5, Não=0; Não se aplica fica fora do total).
  - Novo campo `Processos.Processo_ECodevasf_Link` (URL) — o painel usa
    esse link para tornar o número do processo (`Processo_ECodevasf`)
    clicável.
  - `Tarefas.Criterios_Desempenho` migrou para `Metricas.Criterios_Desempenho`
    (tem mais a ver com o indicador do que com a tarefa em si).
  - Campos removidos: `Disparador`, `Passos`, `Principios`,
    `Resultados_Esperados`, `Materiais_Ferramentas`, `Pessoas_Consultar`
    (Tarefas) e `Produto_Principal` (Subprocessos). `Unidades_Organicas_
    Corresponsaveis` (Macroprocessos/Processos/Subprocessos/Tarefas) tinha
    saído numa revisão anterior e voltou a pedido do usuário — ao final de
    cada aba (mesma posição em que foi restaurada na planilha em uso).

Revisões anteriores (ainda válidas):
  - Nomenclatura padronizada: Unidade_Organica_Responsavel, Beneficiarios,
    Ponto_Focal_Nugep, Produto_Principal (removido depois — ver acima).
  - Colunas removidas: Dono_Processo/Dono (viram registro em Papeis),
    Fase_Ciclo_BPM, Normativos_Aplicaveis/Relacionados (cada normativo é
    uma linha em Documentos, com Link), Categoria/Termos_Relacionados
    (Glossario), Codigo (Repositorio); aba PlanoAcao inteira.
  - Processos.Percentual e todo Vinculo_Nivel são fórmula (não digite).
  - Marcos M1–M10 renomeados para bater 1:1 com a Metodologia (RES
    031/2025).
  - Sem lista suspensa em nenhum campo que aceita MAIS DE UM código/sigla
    separado por ; (o Excel só permite escolher um item por vez de uma
    lista). Ponto_Focal_Nugep tem dropdown (nomes da aba NUGEP), por ser
    campo de valor único.

Revisão desta versão (numeração decimal de Subprocesso dentro de
Subprocesso, pedido do usuário):
  - Quando um Subprocesso é filho de OUTRO Subprocesso (em vez de um
    Processo), seu Código usa numeração decimal: {código do subprocesso
    pai}.{posição do filho, 2 dígitos} — ex.: SP-03 com 3 filhos aninhados
    vira SP-03.01, SP-03.02, SP-03.03. Um filho aninhado, sem essa marca,
    ficaria com o mesmo formato "SP-01" de um subprocesso de 1º nível —
    indistinguível a olho nu sem abrir a Trilha. Mesma lógica se repete em
    qualquer profundidade (SP-03.01.01, SP-03.01.02...) e, no futuro, em
    qualquer outro nível da hierarquia que venha a aninhar dentro de si
    mesmo (hoje só Subprocesso permite isso). Convenção de preenchimento
    manual — não há fórmula calculando o Código.

Revisão desta versão (Cadeia de Valor real e ato normativo):
  - As abas de dado passaram a vir de `data/cadeia_valor.json` — Cadeia de
    Valor Integrada da Codevasf, aprovada pela Resolução nº 1099, de 25 de
    setembro de 2025. São 21 macroprocessos (5 gerenciais, 8 finalísticos,
    8 de suporte) e 70 processos de negócio.
  - **A norma que instituiu cada item não virou coluna.** Ela entra como
    documento na aba `Documentos`, vinculada aos macroprocessos e processos
    pela coluna `Vinculo_Codigo` — é o corpo que o painel já mostra no card
    "Normativos e documentos vinculados". A Resolução nº 1099/2025 e os
    demais documentos de referência (Nota Técnica nº 001/2025-AE/GPE/UNP,
    Resolução nº 31/2025, Guia de Modelagem, N-000, PEI 2025-2030, Plano de
    Ação do Nugep, Decisão nº 1484/2024) estão lá.
  - `Prioridade` dos processos reflete o item 2 da Resolução nº 1099/2025 e o
    item 7 da Nota Técnica nº 001/2025: os 26 processos priorizados estão
    como Alta.
  - Três opções novas na lista `Tipo_Documento`: Resolução, Nota técnica e
    Apresentação.

Revisão desta versão (identificador único de métrica, equipe por processo e
papéis sem o nível Atividade — pedido do usuário):
  - **Metricas.ID usa um único prefixo, `IND-`.** Todas as linhas da aba são
    indicadores; o que distingue um indicador de processo de um SLA ou de um
    ROI é a coluna `Categoria` (lista fechada Processo/SLA/ROI), não o
    prefixo do código. `MET-010`/`MET-011` viraram `IND-010`/`IND-011` (e as
    medições que apontavam para eles, junto).
  - **`Equipe_Gerenciamento_Processos` (lista central de gestores) saiu; no
    lugar entra `Equipe_Processo`** (ID, Processo, Ordem, Nome, Email,
    Telefone, Area), com uma equipe POR processo — é o que a RES 031/2025,
    item 3.7, define. `Processo` guarda a Trilha do processo (lista suspensa,
    valor único por célula). Sem foto nem avatar: só nome, e-mail,
    endereço e área da pessoa. A aba fica ao lado de Papeis, não no fim do
    arquivo.
  - **Papeis não tem mais registro de nível Atividade.** Quem executa,
    aprova, é consultado ou informado se declara no Macroprocesso, no
    Processo e no Subprocesso; a ficha da atividade deixou de mostrar
    "Papéis e envolvidos". `PAP-007` (o único de nível Atividade) saiu e os
    IDs seguintes foram renumerados.

Revisão anterior (Gestor do processo e Equipe de Gerenciamento, pedido
do usuário):
  - Novas colunas em `Processos`, ao final da aba: `Gestor_Nome`,
    `Gestor_Email`, `Gestor_Telefone`, `Gestor_Unidade_Organica` — pessoa
    que responde pelo processo no dia a dia, distinta do
    `Ponto_Focal_Nugep` (o contato durante o mapeamento/análise).
  - Nova aba `Equipe_Gerenciamento_Processos` (Ordem, Nome, Email,
    Telefone, Unidade_Organica): diretório dos gestores de processo,
    independente da lista de integrantes da aba NUGEP — substituída por
    `Equipe_Processo` na revisão acima.
  - Vínculos com caminho completo (`Vinculo_Codigo`, `Reutilizado_Em`,
    `Processos_Relacionados`) usam " › " como separador — não mais "/".
"""

ESQUEMA = {
    "Macroprocessos": [
        ("Codigo", 10), ("Nome", 34), ("Categoria", 13), ("Ordem", 7),
        ("Unidade_Organica_Responsavel", 16), ("Unidades_Organicas_Corresponsaveis", 30),
        ("Descricao", 46), ("Objetivo", 46), ("Entregas", 40), ("Beneficiarios", 34),
        ("Partes_Interessadas", 34), ("Sistemas", 30), ("Imagem_Bizagi", 30), ("Observacoes", 26),
        ("Trilha", 14),
    ],
    "Processos": [
        ("Trilha", 20), ("Macroprocesso", 13), ("Codigo", 9), ("Nome", 34),
        ("Descricao", 44), ("Objetivo", 40),
        ("Unidade_Organica_Responsavel", 16), ("Unidades_Organicas_Corresponsaveis", 30), ("Ponto_Focal_Nugep", 24),
        ("Prioridade", 10), ("Complexidade", 12), ("Maturidade", 13), ("Status_Mapeamento", 14), ("Percentual", 10),
        ("Inicio_Mapeamento", 13), ("Prazo_Previsto", 13), ("Data_Conclusao", 13), ("Ultima_Atualizacao", 13),
        ("M1_Conhecer_Processo", 9), ("M2_Processo_Modelado", 9), ("M3_Subprocessos_Modelados", 9),
        ("M4_ASIS_Modelado", 9), ("M5_ASIS_Validado", 9), ("M6_Procedimento_Aprovado", 9),
        ("M7_Processo_Publicado", 9), ("M8_TOBE_Elaborado", 9), ("M9_TOBE_Aprovado", 9),
        ("M10_Processo_Transformado", 9),
        ("Fornecedores", 30), ("Entradas", 34), ("Saidas", 34), ("Beneficiarios", 28), ("Sistemas", 28),
        ("Processo_ECodevasf", 20), ("Processo_ECodevasf_Link", 30), ("Imagem_Bizagi", 30),
        ("Competencias_Necessarias", 40), ("Fontes_Dados", 30),
        ("Proxima_Acao", 28), ("Pendencia", 24),
        ("Gestor_Nome", 24), ("Gestor_Email", 30), ("Gestor_Telefone", 16),
        ("Gestor_Unidade_Organica", 16),
    ],
    "Subprocessos": [
        ("Trilha", 26), ("Macroprocesso", 13), ("Processo", 10), ("Codigo", 9), ("Nome", 32),
        ("Vinculo_Pai", 12), ("Ordem", 7), ("Descricao", 44), ("Objetivo", 38),
        ("Unidade_Organica_Responsavel", 16), ("Unidades_Organicas_Corresponsaveis", 30),
        ("Reutilizavel", 12), ("Reutilizado_Em", 30),
        ("Entradas", 36), ("Saidas", 36), ("Sistemas", 26), ("Fontes_Dados", 26), ("Imagem_Bizagi", 28),
    ],
    "Atividades": [
        ("Trilha", 32), ("Macroprocesso", 13), ("Processo", 10), ("Subprocesso", 10),
        ("Codigo", 9), ("Nome", 34), ("Vinculo_Pai", 11), ("Ordem", 7), ("Tipo_Atividade", 17),
        ("Descricao", 44), ("Entradas", 34), ("Saidas", 34), ("Sistemas", 22),
    ],
    "Tarefas": [
        ("Trilha", 40), ("Macroprocesso", 13), ("Processo", 10), ("Subprocesso", 10), ("Atividade", 9),
        ("Codigo", 9), ("Nome", 34), ("Ordem", 7), ("Tipo_Tarefa", 15),
        ("Duracao_Estimada", 13), ("Observacoes", 26),
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
        ("ID", 9), ("Nome", 34), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16), ("Categoria", 18),
        ("Descricao_Formula", 40), ("Criterios_Desempenho", 40), ("Unidade", 9), ("Polaridade", 13),
        ("Meta", 8), ("Periodicidade", 13), ("Fonte", 22), ("Observacoes", 26), ("Formula", 46),
    ],
    "Medicoes": [
        ("ID", 9), ("Metrica_ID", 10), ("Data_Medicao", 13), ("Valor", 10), ("Observacao", 28),
    ],
    # Papéis valem para Macroprocesso, Processo e Subprocesso — Atividade e
    # Tarefa não têm papéis próprios (pedido do usuário).
    "Papeis": [
        ("ID", 9), ("Vinculo_Nivel", 14), ("Vinculo_Codigo", 16), ("Papel", 30),
        ("Envolvimento", 16), ("Unidade_Pessoa", 22),
    ],
    "Equipe_Processo": [
        ("ID", 9), ("Processo", 22), ("Ordem", 7), ("Nome", 26),
        ("Email", 30), ("Telefone", 16), ("Area", 14),
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
    "Macroprocessos": ["Trilha"],
    "Processos": ["Trilha", "Percentual"],
    "Subprocessos": ["Trilha", "Macroprocesso", "Processo"],
    "Atividades": ["Trilha", "Macroprocesso", "Processo", "Subprocesso"],
    "Tarefas": ["Trilha", "Macroprocesso", "Processo", "Subprocesso"],
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
    "Status_Marco": ["Sim", "Em andamento", "Não", "Não se aplica"],
    # Documentos, Riscos, Metricas e Papeis se declaram até Subprocesso: risco,
    # indicador, normativo e papel de nível Atividade ou Tarefa repetiriam em
    # nível 4 e 5 o que já está declarado em nível 3 (pedido do usuário). Só
    # Regras desce a Atividade — a regra de negócio é o único vínculo que o
    # painel ainda mostra na ficha da atividade, e usa Nivel_Vinculo_Regras.
    # Vinculo_Nivel é coluna calculada em todas essas abas: quem restringe o
    # nível de fato é a validação de Vinculo_Codigo (ver dv_profundidade).
    # Papéis só existem no Processo: o RACI se pactua com quem executa o
    # processo. No Macroprocesso (agrupador) e no Subprocesso a mesma matriz
    # se repetiria em três níveis — por isso a lista de Papeis para em
    # Processo, e a validação de profundidade barra Subprocesso pra baixo.
    "Nivel_Vinculo": ["Macroprocesso", "Processo", "Subprocesso"],
    "Tipo_Atividade": ["Agregação de Valor", "Transferência", "Controle"],
    "Tipo_Tarefa": ["Manual", "Automatizada", "Regra de negócio"],
    "Tipo_Documento": ["Resolução", "Nota técnica", "Procedimento (PRO)", "Manual",
                       "Norma interna", "Legislação externa", "Plano", "Apresentação",
                       "Formulário/Modelo", "Ata de reunião", "Diagrama BPMN", "Checklist",
                       "Relatório", "Outro"],
    "Situacao_Documento": ["Vigente", "Em elaboração", "Em revisão", "Revogado"],
    # Taxonomia da IN Conjunta MP/CGU nº 01/2016: Estratégico é categoria
    # própria, não um caso de Operacional.
    "Categoria_Risco": ["Estratégico", "Operacional", "Legal/Conformidade", "Pessoas",
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
    # Vocabulário da coluna Repositorio.Fase_Ciclo. "Transversal" é para a
    # referência que não pertence a uma fase do ciclo — metodologia, guia,
    # CBOK, BPMN. NÃO usar "Todas as fases": os combos do painel derivam as
    # opções dos valores do dado e reservam o literal "Todas as/os {plural}"
    # para o placeholder de "sem filtro", então esse texto como valor criaria
    # dois rótulos idênticos com sentidos opostos no mesmo controle.
    # Acrescentar item a uma lista existente é seguro: só lista NOVA desloca a
    # coluna das seguintes (ver nota no fim do dict).
    "Fase_Instrumento": ["Planejamento", "Análise", "Desenho", "Implementação", "Monitoramento",
                         "Refinamento", "Transversal"],
    "Categoria_FAQ": ["Conceitos básicos", "Modelagem e SIPOC", "Cadeia de Valor e governança",
                      "Indicadores, metas e riscos", "Plano de Ações AE/GPE", "Como usar o painel"],
    # Ao final do dicionário: ref() calcula a coluna pela posição no dict, e o
    # hand-edit já em uso na planilha real colocou esta lista na coluna Y
    # (a última) — nova lista sempre entra no fim, nunca no meio, senão as
    # colunas de todas as listas seguintes mudam de letra na próxima geração.
    "Categoria_Metrica": ["Processo", "SLA", "ROI"],
    "Nivel_Vinculo_Regras": ["Macroprocesso", "Processo", "Subprocesso", "Atividade", "Tarefa"],
}

# Até que nível Vinculo_Codigo pode descer em cada aba (2 = Subprocesso,
# 4 = Tarefa). Vira validação de célula na planilha, para que a coluna
# calculada Vinculo_Nivel não possa resultar em Atividade ou Tarefa onde não
# deve. A validação procura os prefixos de PREFIXO_NIVEL, não conta
# separadores: Vinculo_Codigo aceita várias trilhas na mesma célula.

# Invariante de referência cruzada: toda validação que aponta para outra aba
# tem de ser reconferida SEMPRE que a aba-alvo muda de tamanho, junto com a
# contagem de linhas. A regra do alvo depende do tipo de aba:
#   - aba de dados  → última linha = nº de registros + 1 (o cabeçalho)
#   - aba Listas    → última linha = fim REAL daquela coluna, porque cada
#                     coluna é um vocabulário independente; usar a contagem de
#                     linhas da aba Listas dá falso positivo em 24 das 25.
# Conferir as duas coisas juntas: uma faixa curta não quebra o arquivo, ela
# marca como inválido o dado que já está gravado (o dropdown de ponto focal
# chegou a oferecer 12 dos 59 nomes do Nugep).
PROFUNDIDADE_VINCULO = {
    "Documentos": 2, "Riscos": 2, "Metricas": 2, "Papeis": 1, "Regras": 3,
}

# Prefixo de código por nível — 2 letras + número sequencial, reiniciado a
# cada novo vínculo com o pai (ver docstring do módulo).
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
