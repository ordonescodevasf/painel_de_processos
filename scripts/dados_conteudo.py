# -*- coding: utf-8 -*-
"""Conteúdo das abas Jornada, Repositorio, NUGEP, Glossario, Competencias, FAQ e Parametros.
Transcrito/adaptado do painel anterior (Lovable) — edite direto NA PLANILHA depois."""

# (Ordem, Fase, Nome, Duracao, Objetivo, Atividades_Chave;, Quem_Faz;, Entregaveis;, Sentimento_Usuario)
JORNADA = [
 (1, "Descobrir", "Escuta e contextualização", "1–2 semanas",
  "Compreender a dor real, o contexto e as expectativas de quem executa e de quem se beneficia do processo.",
  "Reunião de sensibilização com a unidade demandante; Entrevistas semiestruturadas com executores e gestores; Aplicação do formulário de contextualização; Escuta ativa das partes interessadas (jornada empática)",
  "NUGEP; Ponto focal do Nugep; Dono do processo",
  "Ata de reunião; Formulário preenchido; Mapa de partes interessadas",
  "Curiosidade e ansiedade: “por que estão olhando meu trabalho?”"),
 (2, "Definir", "Delimitação de escopo", "1 semana",
  "Traçar limites claros: onde começa, onde termina, quem participa e o que está fora do mapeamento.",
  "SIPOC preliminar (Fornecedores, Entradas, Processo, Saídas, Clientes); Identificação de sistemas e normativos aplicáveis; Definição de premissas, restrições e critérios de sucesso; Termo de abertura do projeto de mapeamento (PMBOK)",
  "Gerente do projeto; NUGEP",
  "Ficha do mapeamento publicada",
  "Segurança: “agora sei o que esperam de mim”"),
 (3, "Desenvolver", "Modelagem AS-IS", "2–4 semanas",
  "Fotografar o processo como ele é hoje, incluindo variações, filas, retrabalhos e workarounds.",
  "Workshops de modelagem com quem executa; Desenho BPMN no Bizagi (com anexação de evidências); Registro de hipóteses, dores e memória do mapeamento (no processo e-Codevasf); Identificação de riscos, gargalos e handoffs",
  "Equipe de mapeamento; Executores do processo",
  "Diagrama BPMN AS-IS; Lista de dores e riscos",
  "Reconhecimento: “alguém finalmente enxerga o que a gente vive”"),
 (4, "Desenvolver", "Validação AS-IS", "1 semana",
  "Confirmar com quem executa e com quem decide que o retrato do processo é fiel.",
  "Reunião de validação com donos e executores; Ajustes finos no BPMN e nos artefatos; Homologação formal do AS-IS",
  "Dono do processo; NUGEP",
  "AS-IS homologado e publicado no repositório",
  "Orgulho: “é isso que a gente faz”"),
 (5, "Desenvolver", "Análise do processo", "1–2 semanas",
  "Investigar as causas dos problemas do AS-IS validado e preparar o plano de melhorias, antes de redesenhar o processo.",
  "Entrevista estruturada por componente habilitador (pessoas, fluxo de trabalho, TI, regras/políticas, métricas, infraestrutura, colaboração); Identificação de causas e consequências dos problemas; Elaboração de propostas de melhoria; Plano de Ação pela técnica 5W2H",
  "Equipe de gerenciamento de processo; Dono do processo",
  "Lista de problemas e causas; Plano de Ação (5W2H)",
  "Alívio: “agora vamos atacar o que realmente trava o meu trabalho”"),
 (6, "Desenvolver", "Redesenho TO-BE", "2–4 semanas",
  "Desenhar o processo desejado, eliminando desperdícios e melhorando a experiência de quem executa e de quem recebe.",
  "Análise de causas-raiz das dores identificadas; Ideação de melhorias com a equipe da área; Modelagem TO-BE (BPMN) e simulação de cenários; Definição de indicadores de desempenho e riscos residuais",
  "NUGEP; Área; Escritório de Riscos (2ª linha)",
  "Diagrama TO-BE; Indicadores; Plano de transição",
  "Esperança: “dá pra melhorar de verdade”"),
 (7, "Entregar", "Publicação e transição", "2 semanas",
  "Colocar o novo processo em operação com apoio, treinamento e comunicação clara.",
  "Publicação no repositório institucional de processos; Capacitação dos executores e multiplicadores; Atualização de normativos e do e-Codevasf; Lançamento do TO-BE",
  "Área responsável; NUGEP; Comunicação",
  "Processo publicado; Equipe capacitada; Comunicação institucional",
  "Confiança: “fui preparado para o novo jeito”"),
 (8, "Evoluir", "Monitoramento e indicadores", "Contínuo",
  "Acompanhar o desempenho real do processo e a experiência de quem o executa e recebe.",
  "Coleta de indicadores (eficiência, eficácia, qualidade, satisfação); Monitoramento de riscos e controles (2ª linha); Escuta contínua da unidade e dos beneficiários",
  "Dono do processo; NUGEP; Escritório de Riscos",
  "Painel de indicadores; Radar de riscos atualizado",
  "Pertencimento: “meu trabalho é acompanhado e valorizado”"),
 (9, "Evoluir", "Melhoria contínua", "Contínuo",
  "Fechar o ciclo: aprender, ajustar e reiniciar a jornada sempre que necessário (PDCA/CBOK 4.0).",
  "Retrospectivas periódicas com a área; Revisão do processo com base em dados e feedback; Abertura de novo ciclo de redesenho quando necessário",
  "NUGEP; Comunidade de prática BPM",
  "Novo ciclo de mapeamento ou ajustes incrementais",
  "Autonomia: “nosso processo é vivo, evolui com a gente”"),
]

_MET = "https://intraplone.codevasf.gov.br/documentos_normativos_referencia/governanca-e-gestao/copy_of_planos/met_gerenciamento-processos_res-031_2025-01-08.pdf"
_GUIA = "https://intraplone.codevasf.gov.br/documentos_normativos_referencia/governanca-e-gestao/copy_of_planos/guia_gerenciamento-processos_res-031_2025-01-08.pdf"

# (ID, Categoria, Fase_Ciclo, Titulo, Descricao, Fonte, Link, Ordem)
REPOSITORIO = [
 ("REP-001", "Documento oficial", "",
  "Metodologia de Gerenciamento de Processos da Codevasf",
  "Documento normativo que institui a metodologia corporativa de gerenciamento de processos (disponível na intranet/e-Codevasf).",
  "Codevasf", _MET, 1),
 ("REP-002", "Documento oficial", "",
  "Guia de Gerenciamento de Processos da Codevasf",
  "Guia prático que detalha etapas, papéis, artefatos e padrões do mapeamento (disponível na intranet/e-Codevasf).",
  "Codevasf", _GUIA, 2),
 ("REP-005", "Template", "Desenho",
  "Modelo de Diagramas BPMN",
  "Template padrão para modelagem em BPMN 2.0 (pools, lanes, eventos, atividades e gateways), em formato Bizagi e draw.io.",
  "Codevasf", "https://exemplo.codevasf.gov.br/modelos/template-bpmn.bpm", 5),
 ("REP-006", "Template", "Análise",
  "Roteiro de Reunião de Contextualização",
  "Guia para conduzir a reunião de contextualização: perguntas-guia, dinâmicas de elicitação e checklist de tópicos.",
  "Codevasf", "https://exemplo.codevasf.gov.br/modelos/roteiro-contextualizacao.docx", 6),
 ("REP-011", "Instrumento", "Desenho",
  "Checklist BPMN 2.0 (notação)",
  "Lista de verificação dos elementos mínimos para um diagrama BPMN aderente ao padrão.",
  "CBOK 4.0", "https://exemplo.codevasf.gov.br/instrumentos/bpmn-01.docx", 11),
 ("REP-012", "Instrumento", "Desenho",
  "Matriz RACI",
  "Atribui Responsável (R), Aprovador (A), Consultado (C) e Informado (I) por atividade.",
  "PMBOK", "https://exemplo.codevasf.gov.br/instrumentos/raci-01.xlsx", 12),
 ("REP-013", "Instrumento", "Monitoramento",
  "Canvas de Indicador (KPI)",
  "Define nome, fórmula, fonte, meta, periodicidade e responsável de um indicador.",
  "Codevasf", "https://exemplo.codevasf.gov.br/instrumentos/kpi-01.docx", 13),
 ("REP-014", "Instrumento", "Análise",
  "Registro de Riscos (ISO 31000)",
  "Identifica risco, causa, efeito, probabilidade, impacto, resposta e responsável.",
  "ISO 31000", "https://exemplo.codevasf.gov.br/instrumentos/risk-01.xlsx", 14),
 ("REP-015", "Instrumento", "Análise",
  "Roteiro de entrevista AS-IS",
  "Perguntas-guia para levantar o processo como ele é hoje com quem executa.",
  "CBOK 4.0", "https://exemplo.codevasf.gov.br/instrumentos/as-is-01.docx", 15),
 ("REP-016", "Instrumento", "Implementação",
  "Plano de transição AS-IS → TO-BE",
  "Compara o processo atual e o futuro e organiza as ações de migração.",
  "CBOK 4.0", "https://exemplo.codevasf.gov.br/instrumentos/to-be-01.docx", 16),
 ("REP-017", "Instrumento", "Refinamento",
  "Diagnóstico rápido de maturidade (PEMM)",
  "Autoavaliação em 5 dimensões: desenho, executores, dono, infraestrutura e indicadores.",
  "CBOK 4.0", "https://exemplo.codevasf.gov.br/instrumentos/mat-01.xlsx", 17),
 ("REP-018", "Ferramenta", "Desenho",
  "Bizagi Modeler",
  "Ferramenta gratuita de modelagem BPMN 2.0, recomendada pela UNP para os diagramas AS-IS e TO-BE.",
  "Bizagi", "https://www.bizagi.com/pt/plataforma/modeler", 18),
 ("REP-022", "Referência", "",
  "BPM CBOK 4.0 — ABPMP International",
  "Guia de referência em gerenciamento de processos de negócio; padrão metodológico principal da UNP.",
  "ABPMP", "https://www.abpmp-br.org", 22),
 ("REP-033", "Referência", "",
  "BPMN 2.0 — Business Process Model and Notation",
  "Especificação oficial da notação usada nos diagramas (Bizagi e demais ferramentas de modelagem): símbolos, semântica de execução e serialização XML.",
  "OMG (Object Management Group)", "https://www.omg.org/spec/BPMN/2.0/", 23),
 ("REP-025", "Template", "Análise",
  "Mapa de Partes Interessadas (modelo)",
  "Identifica e classifica stakeholders por interesse e influência, com estratégia de engajamento por grupo.",
  "PMBOK", "https://exemplo.codevasf.gov.br/modelos/mapa-partes-interessadas.xlsx", 25),
]



# (Ordem, Nome, Papel, Unidade_Sigla, Unidade_Nome, Email, Telefone, Foto, Hierarquia)
#
# Foto: URL de uma imagem pública (aparece no lugar das iniciais do avatar).
#   As URLs abaixo são retratos de exemplo (randomuser.me) só para testar a
#   exibição — troque pelas fotos oficiais quando houver.
# Hierarquia: 1 = Gerente-Executivo (AE); 2 = Gerente (AE/GPE); 3 = equipe da
#   Unidade (AE/GPE/UNP); 0 = interlocutor de outra área. Os níveis 1 e 2
#   aparecem só no bloco "Contato institucional", para mostrar a hierarquia
#   acima da Unidade — não entram na lista de integrantes do núcleo.
_FOTO = "https://randomuser.me/api/portraits/"
NUGEP = [
 (1, "Ricardo Nunes Vasconcelos", "Gerente-Executivo da Área de Gestão Estratégica", "AE",
  "Área de Gestão Estratégica", "ricardo.vasconcelos@codevasf.gov.br", "(61) 2028-4400",
  _FOTO + "men/32.jpg", 1),
 (2, "Patrícia Moreira Lopes", "Gerente de Planejamento Estratégico", "AE/GPE",
  "Gerência de Planejamento Estratégico", "patricia.lopes@codevasf.gov.br", "(61) 2028-4430",
  _FOTO + "women/68.jpg", 2),
 (3, "Bruna Souza", "Coordenadora do NUGEP", "AE/GPE/UNP",
  "Unidade de Gestão Normativa e de Processos", "bruna.souza@codevasf.gov.br", "(61) 2028-4441",
  _FOTO + "women/44.jpg", 3),
 (4, "Carlos Eduardo Lima", "Analista de Processos", "AE/GPE",
  "Gerência de Planejamento Estratégico", "carlos.lima@codevasf.gov.br", "(61) 2028-4442",
  _FOTO + "men/45.jpg", 0),
 (5, "Daniela Ribeiro", "Especialista em BPM", "AT/GTI",
  "Gerência de Tecnologia da Informação", "daniela.ribeiro@codevasf.gov.br", "(61) 2028-4501",
  _FOTO + "women/12.jpg", 0),
 (6, "Eduardo Martins", "Ponto Focal do Nugep — Processos Finalísticos", "AR/GPR",
  "Gerência de Projetos de Revitalização", "eduardo.martins@codevasf.gov.br", "(61) 2028-4612",
  _FOTO + "men/76.jpg", 0),
 (7, "Fernanda Alves", "Analista de Riscos e Conformidade", "AE/GAG",
  "Gerência de Auditoria e Governança", "fernanda.alves@codevasf.gov.br", "(61) 2028-4380",
  _FOTO + "women/29.jpg", 0),
 (8, "Gustavo Pereira", "Apoio Metodológico — PMBOK/CBOK", "AE/GPE/UNP",
  "Unidade de Gestão Normativa e de Processos", "gustavo.pereira@codevasf.gov.br", "(61) 2028-4443",
  _FOTO + "men/22.jpg", 3),
 (9, "Juliana Tavares", "Analista de Processos", "AE/GPE/UNP",
  "Unidade de Gestão Normativa e de Processos", "juliana.tavares@codevasf.gov.br", "(61) 2028-4444",
  _FOTO + "women/8.jpg", 3),
 (10, "Rafael Menezes", "Analista de Dados e Indicadores", "AE/GPE",
  "Gerência de Planejamento Estratégico", "rafael.menezes@codevasf.gov.br", "(61) 2028-4445",
  _FOTO + "men/54.jpg", 0),
 (11, "Camila Duarte", "Ponto Focal do Nugep — Processos de Suporte", "AA/GLC",
  "Gerência de Licitações e Contratos", "camila.duarte@codevasf.gov.br", "(61) 2028-4520",
  _FOTO + "women/90.jpg", 0),
 (12, "Paulo Henrique Barros", "Apoio de Comunicação Institucional", "PR/ASCOM",
  "Assessoria de Comunicação (fictícia)", "paulo.barros@codevasf.gov.br", "(61) 2028-4630",
  _FOTO + "men/60.jpg", 0),
]

# (Termo, Definicao, Fonte)
GLOSSARIO = [
 ("ABPMP", "Association of Business Process Management Professionals — entidade internacional sem fins lucrativos que mantém o CBOK e certificações como CBPP (Certified Business Process Professional).", "ABPMP"),
 ("Atividade", "Conjunto de tarefas executadas para transformar entradas em saídas dentro de um processo. É o nível operacional mapeado em diagramas BPMN.", "CBOK 4.0"),
 ("Arquitetura de Processos", "Visão estruturada e hierárquica de todos os processos da organização, do macroprocesso à tarefa, mostrando relações, donos e contribuição para a estratégia.", "CBOK 4.0"),
 ("AS-IS", "Modelo do processo no estado atual (como ele realmente é executado hoje), usado como base de diagnóstico antes de propor melhorias.", "CBOK 4.0"),
 ("Análise de Processos", "Estudo do processo AS-IS para identificar gargalos, retrabalho, desperdícios e oportunidades de melhoria. Usa ferramentas como Ishikawa, 5 porquês e análise de valor agregado.", "CBOK 4.0"),
 ("Agile", "Abordagem iterativa e incremental de gerenciamento, baseada no Manifesto Ágil, com entregas frequentes e adaptação contínua.", "Manifesto Ágil"),
 ("AE", "Área de Estratégia e Finanças da Codevasf, instância superior da GPE/UNP.", "Codevasf"),
 ("Accountability", "Obrigação de prestar contas e responder pelos resultados, especialmente no setor público.", "TCU"),
 ("BPM", "Business Process Management — disciplina gerencial integrada que combina técnicas, métodos e ferramentas para identificar, modelar, executar, monitorar e melhorar continuamente os processos de negócio, alinhando-os à estratégia e à entrega de valor ao cliente.", "ABPMP CBOK 4.0"),
 ("BPMN", "Business Process Model and Notation — padrão da OMG para modelagem gráfica de processos de negócio. Define símbolos para atividades, eventos, gateways, fluxos e raias.", "OMG / CBOK 4.0"),
 ("Beneficiário / Cidadão-usuário", "Destinatário da saída do processo no setor público. Pode ser interno (outra unidade da Codevasf) ou externo — cidadão, irrigante, comunidade beneficiária, ente federativo ou órgão parceiro.", "SIPOC"),
 ("Linha de Base (Indicador)", "Valor de referência inicial de um indicador, usado como base para comparar resultados futuros.", "PMBOK"),
 ("Burndown", "Gráfico que mostra a quantidade de trabalho restante ao longo do tempo, comum em métodos ágeis.", "Scrum"),
 ("Burnup", "Gráfico que mostra o trabalho concluído acumulado ao longo do tempo, comparando com o escopo total.", "Agile"),
 ("CBOK", "Guide to the Business Process Management Common Body of Knowledge — referência da ABPMP que consolida conceitos, práticas e competências essenciais de BPM. A versão 4.0 reorganiza o conteúdo em 9 áreas de conhecimento.", "ABPMP CBOK 4.0"),
 ("Cadeia de Valor", "Representação visual dos macroprocessos de uma organização ordenados segundo sua contribuição para a entrega de valor. Permite visão sistêmica e priorização do mapeamento.", "Porter / CBOK 4.0"),
 ("Ciclo de Vida BPM", "Sequência iterativa de fases para gerenciar processos: planejar, analisar, desenhar, implementar, monitorar e refinar (PDCA aplicado a processos).", "CBOK 4.0"),
 ("Cronograma", "Modelo que apresenta a sequência, datas e durações das atividades do projeto.", "PMBOK"),
 ("Caminho Crítico", "Sequência de atividades que determina a menor duração possível do projeto. Atrasos em qualquer atividade do caminho crítico atrasam o projeto.", "PMBOK / CPM"),
 ("CPI", "Cost Performance Index — índice de desempenho de custo. CPI < 1 indica estouro de orçamento.", "EVM"),
 ("Compliance", "Conformidade com leis, regulamentos, normas e políticas internas. Inclui prevenção a fraudes e corrupção.", "Lei 12.846/2013"),
 ("Controle Interno", "Conjunto de atividades, planos e métodos adotados para salvaguardar ativos, conferir exatidão e segurança aos dados e promover a eficiência operacional.", "INTOSAI / IN 03/2017"),
 ("Desperdício (Muda)", "Toda atividade que consome recursos sem agregar valor ao cliente. Os 7+1 desperdícios clássicos: superprodução, espera, transporte, processamento excessivo, estoque, movimento, defeito e talento não utilizado.", "Lean"),
 ("Dono do Processo", "Pessoa formalmente responsável pelo desempenho, integridade e melhoria contínua de um processo de ponta a ponta, independentemente da estrutura hierárquica.", "CBOK 4.0"),
 ("Diagrama de Tartaruga", "Ferramenta de modelagem que detalha um processo por entradas, saídas, recursos, métodos, pessoas e indicadores, formando o desenho de uma tartaruga.", "ISO 9001"),
 ("Evento", "Acontecimento que afeta o fluxo do processo: início, fim, intermediário, de tempo, mensagem ou erro.", "BPMN"),
 ("Escritório de Processos (BPM Office)", "Unidade organizacional responsável por orientar metodologia, manter o repositório de processos e apoiar áreas no mapeamento. Na Codevasf, esse papel é da UNP/GPE.", "CBOK 4.0"),
 ("Entrada", "Insumo, informação, documento ou recurso necessário para iniciar ou executar uma atividade do processo.", "SIPOC"),
 ("Eficiência", "Relação entre o que foi produzido (saída) e os recursos consumidos (entrada). “Fazer certo as coisas”.", "CBOK"),
 ("Eficácia", "Grau em que os resultados pretendidos foram alcançados. “Fazer as coisas certas”.", "CBOK"),
 ("Efetividade", "Impacto real e duradouro do processo no público-alvo ou na sociedade. Combina eficiência e eficácia com sustentabilidade.", "Gestão Pública"),
 ("Escopo", "Soma dos produtos, serviços e resultados a serem entregues no projeto. Inclui escopo do produto e escopo do trabalho.", "PMBOK"),
 ("EAP", "Estrutura Analítica do Projeto (Work Breakdown Structure — WBS) — decomposição hierárquica do escopo em entregas e pacotes de trabalho.", "PMBOK"),
 ("Entrega", "Produto, resultado ou capacidade único e verificável, obtido a partir de uma fase ou de todo o projeto.", "PMBOK"),
 ("EVM", "Earned Value Management — técnica que integra escopo, prazo e custo para medir o desempenho do projeto (SPI, CPI).", "PMBOK"),
 ("Fornecedor", "Pessoa, área ou sistema que fornece insumos (entradas) ao processo.", "SIPOC"),
 ("Fluxograma", "Representação gráfica sequencial das etapas de um processo, usando símbolos padronizados para atividade, decisão, início e fim.", "ASME / ISO"),
 ("Gateway", "Elemento BPMN que controla a divergência ou convergência do fluxo: decisão exclusiva, paralela ou inclusiva.", "BPMN"),
 ("Gargalo", "Ponto do processo onde a capacidade é menor que a demanda, restringindo a vazão de todo o fluxo. Identificado por filas, atrasos e acúmulo de trabalho.", "Teoria das Restrições"),
 ("Governança de Processos", "Estrutura de papéis, responsabilidades, políticas e instâncias decisórias que garante que os processos sejam executados, monitorados e melhorados conforme a estratégia.", "CBOK 4.0"),
 ("Gantt", "Gráfico de barras horizontais que representa o cronograma do projeto ao longo do tempo, com início, duração e fim de cada atividade.", "Henry Gantt"),
 ("GPE", "Gerência de Planejamento Estratégico da Codevasf, à qual a UNP está subordinada.", "Codevasf"),
 ("Governança Pública", "Conjunto de mecanismos de liderança, estratégia e controle postos em prática para avaliar, direcionar e monitorar a gestão pública.", "Decreto 9.203/2017"),
 ("Handoff", "Transferência de responsabilidade entre áreas, sistemas ou pessoas no fluxo do processo. Cada handoff é ponto crítico de falha e atraso.", "CBOK 4.0"),
 ("Indicador", "Medida quantitativa ou qualitativa que reflete o desempenho de um processo, atividade ou objetivo. Deve ter fórmula, fonte, meta, periodicidade e responsável.", "CBOK 4.0"),
 ("Impacto", "Magnitude das consequências de um evento de risco caso ele se concretize, normalmente avaliado em escala.", "ISO 31000"),
 ("Kaizen", "Filosofia japonesa de melhoria contínua que promove pequenas mudanças incrementais com participação de toda a equipe.", "Lean"),
 ("KPI", "Key Performance Indicator — indicador-chave de desempenho. Métrica quantificável usada para avaliar o sucesso de um processo, projeto ou objetivo estratégico.", "BSC / CBOK"),
 ("Kanban", "Método visual de gestão do fluxo de trabalho, com cartões e colunas (To Do, Doing, Done) e limites de WIP.", "Toyota / Kanban Method"),
 ("Lean", "Conjunto de práticas focadas na eliminação de desperdícios (muda) e maximização de valor para o cliente, oriundas do Sistema Toyota de Produção.", "Toyota / Lean"),
 ("Linha de Base (Projeto)", "Versão aprovada do escopo, prazo ou custo, usada como referência para comparação com o desempenho real.", "PMBOK"),
 ("Lições Aprendidas", "Conhecimento adquirido durante o projeto, registrado para uso em projetos futuros. Inclui o que deu certo, o que deu errado e recomendações.", "PMBOK"),
 ("LGPD", "Lei Geral de Proteção de Dados Pessoais (Lei 13.709/2018) — disciplina o tratamento de dados pessoais por pessoas físicas e jurídicas, públicas e privadas.", "Lei 13.709/2018"),
 ("Macroprocesso", "Agrupamento dos grandes processos de uma organização que atravessam várias áreas e contribuem de forma estratégica para a cadeia de valor. Na Codevasf, são classificados em gerenciais, finalísticos e de suporte.", "CBOK 4.0 / Codevasf"),
 ("Modelagem de Processos", "Representação gráfica e textual de um processo, mostrando atividades, papéis, fluxos, decisões e eventos. Geralmente utiliza notação BPMN.", "CBOK 4.0"),
 ("Melhoria Contínua", "Esforço sistemático e incremental para aperfeiçoar processos, fundamentado em ciclos PDCA e Kaizen.", "Deming / CBOK 4.0"),
 ("Maturidade em Processos", "Grau de evolução da gestão por processos em uma organização, geralmente medido em 5 níveis (inicial, gerenciado, definido, mensurado, otimizado).", "CMMI / BPM-MM"),
 ("Meta", "Valor numérico que se deseja alcançar para um indicador em um prazo determinado, geralmente desafiador e mensurável.", "BSC"),
 ("Matriz de Risco", "Ferramenta que cruza probabilidade e impacto para classificar e priorizar riscos (baixo, moderado, alto, extremo).", "ISO 31000"),
 ("Mitigação", "Ação para reduzir a probabilidade ou o impacto de um risco, ou para tratar suas consequências.", "PMBOK"),
 ("Marco", "Ponto significativo no cronograma, normalmente a conclusão de uma entrega importante. Não consome tempo.", "PMBOK"),
 ("Matriz RACI", "Matriz de responsabilidades que define quem é Responsável (executa), Accountable (responde), Consultado e Informado em cada atividade.", "PMBOK"),
 ("Mapeamento de Processos (Codevasf)", "Trabalho conduzido pela UNP em parceria com as unidades gestoras para documentar processos em formato padronizado (SIPOC, BPMN, indicadores e riscos), conforme a RES 031/2025.", "RES 031/2025"),
 ("Metodologia Codevasf", "Conjunto de etapas, artefatos e padrões definidos na Metodologia e no Guia de Gerenciamento de Processos (RES 031/2025) para conduzir o mapeamento e a melhoria dos processos.", "RES 031/2025"),
 ("Notação", "Conjunto padronizado de símbolos para representar processos. BPMN é a mais usada em BPM.", "CBOK"),
 ("Processo de Negócio", "Conjunto definido de atividades, executadas por pessoas ou sistemas, que entrega um resultado de valor a um cliente interno ou externo. Tem início, fim, entradas, saídas e mensuração claros.", "CBOK 4.0"),
 ("Pool", "Contêiner BPMN que representa um participante (organização, sistema ou área) e agrupa raias.", "BPMN"),
 ("PDCA", "Plan-Do-Check-Act — ciclo de melhoria contínua atribuído a Deming/Shewhart: planejar, executar, verificar e agir corretivamente.", "Deming"),
 ("Processo de Trabalho", "Expressão usual na administração pública brasileira para o conjunto de atividades executadas por uma unidade. Não é um nível da hierarquia do CBOK 4.0, que adota processo de negócio → subprocesso → atividade → tarefa — nomenclatura seguida por este painel.", "Adm. Pública / CBOK 4.0"),
 ("Processo Ponta a Ponta", "Visão completa do processo, do primeiro evento desencadeador até a entrega final ao cliente, atravessando todas as áreas envolvidas.", "CBOK 4.0"),
 ("Probabilidade", "Estimativa da chance de ocorrência de um evento de risco, geralmente em escala (muito baixa, baixa, média, alta, muito alta).", "ISO 31000"),
 ("Plano de Mitigação", "Conjunto estruturado de ações, prazos e responsáveis para tratar um risco específico.", "PMBOK"),
 ("PMBOK", "Project Management Body of Knowledge — guia do PMI que consolida boas práticas, princípios e domínios de desempenho em gerenciamento de projetos.", "PMI"),
 ("PMI", "Project Management Institute — instituto internacional que mantém o PMBOK e certificações como PMP.", "PMI"),
 ("Projeto", "Esforço temporário empreendido para criar um produto, serviço ou resultado único. Tem início e fim definidos.", "PMBOK"),
 ("Programa", "Conjunto de projetos relacionados, gerenciados de forma coordenada para obter benefícios que não seriam alcançados se gerenciados individualmente.", "PMBOK"),
 ("Portfólio", "Conjunto de projetos, programas e operações gerenciados em grupo para atingir objetivos estratégicos.", "PMBOK"),
 ("Pacote de Trabalho", "Componente de menor nível da EAP, para o qual se pode estimar custo, prazo e responsável.", "PMBOK"),
 ("Plano de Ações AE/GPE", "Plano de Ações 2026 conduzido pela AE/GPE/UNP para estruturar o escritório de processos, mapear processos prioritários e implantar a governança.", "Codevasf 2026"),
 ("Raia (Swimlane)", "Faixa horizontal ou vertical no diagrama BPMN que indica o papel, área ou sistema responsável pelas atividades nela contidas.", "BPMN"),
 ("Reengenharia", "Repensar e redesenhar radicalmente processos para obter ganhos expressivos em custo, qualidade, serviço e velocidade. Diferente da melhoria contínua incremental.", "Hammer & Champy"),
 ("Risco", "Efeito da incerteza sobre os objetivos. Pode ser ameaça (negativo) ou oportunidade (positivo). Tem causa, evento e consequência.", "ISO 31000 / PMBOK"),
 ("Subprocesso", "Recorte de um processo, com objetivo específico, que pode ser detalhado em atividades e tarefas. Permite gerenciar partes complexas do processo separadamente.", "CBOK 4.0"),
 ("SIPOC", "Ferramenta visual que descreve um processo em alto nível por meio de cinco elementos: Fornecedores, Entradas, Processo, Saídas e Beneficiários.", "Six Sigma / CBOK"),
 ("Saída", "Produto, serviço, informação ou decisão gerada pelo processo e entregue a um cliente.", "SIPOC"),
 ("SLA", "Service Level Agreement — acordo formal entre prestador e cliente que define níveis de serviço, prazos, qualidade e penalidades.", "ITIL"),
 ("Stakeholder", "Parte interessada — pessoa, grupo ou organização que afeta ou é afetada pelo projeto.", "PMBOK"),
 ("Scrum", "Framework ágil baseado em sprints curtas, papéis definidos (Product Owner, Scrum Master, Time) e cerimônias (planning, daily, review, retro).", "Schwaber & Sutherland"),
 ("Sprint", "Iteração de duração fixa (geralmente 1 a 4 semanas) em Scrum, ao final da qual se entrega um incremento potencialmente utilizável.", "Scrum"),
 ("SPI", "Schedule Performance Index — índice de desempenho de prazo do projeto. SPI < 1 indica atraso.", "EVM"),
 ("Status de Mapeamento", "Indicador do estágio em que se encontra o mapeamento de um processo: Não iniciado, Em andamento, Concluído ou Suspenso.", "Codevasf"),
 ("Semáforo do Processo", "Sinal visual (verde, amarelo, vermelho) que combina prazo e percentual de execução do mapeamento.", "Codevasf"),
 ("Tarefa", "Menor unidade de trabalho identificada em uma atividade. Pode ser manual, automatizada ou de regra de negócio.", "CBOK 4.0"),
 ("TO-BE", "Modelo do processo no estado futuro desejado, após análise e proposta de melhorias. Serve de referência para a implementação.", "CBOK 4.0"),
 ("UNP", "Unidade de Gestão Normativa e de Processos da Codevasf — vinculada à GPE/AE. Responsável por conduzir a metodologia de mapeamento e a governança dos processos da Empresa.", "RES 031/2025"),
 ("Unidade Gestora", "Área da Codevasf responsável pelo processo, por seus resultados e pela manutenção da documentação atualizada.", "Metodologia Codevasf"),
 ("Valor Agregado", "Atributo de uma atividade que transforma a entrada de forma percebida como valiosa pelo cliente. Atividades sem valor agregado devem ser eliminadas ou reduzidas.", "Lean / CBOK"),
 ("WIP", "Work in Progress — quantidade de trabalho em andamento. Limitar o WIP reduz multitarefa e acelera entregas.", "Kanban / Lean"),
 # Termos abaixo: auditoria contra a RES 031/2025 (§2) — definidos na
 # Resolução e usados pelo painel, mas ausentes do Glossário até então.
 ("5W2H", "Checklist de atividades, prazos e responsabilidades usado no plano de ação da etapa de análise: o que, por que, onde, quando, por quem, como e quanto vai custar.", "RES 031/2025"),
 ("Ator do Processo", "Empregado que participa, em algum momento, da execução do processo.", "RES 031/2025"),
 ("Diagrama", "Representação gráfica que demonstra os principais elementos do fluxo do processo, ajudando a identificar e entender rapidamente suas atividades.", "RES 031/2025"),
 ("Equipe de Gerenciamento de Processo", "Grupo formado pelo ponto focal do Nugep, pelo gestor do processo e por atores do processo, responsável por gerenciar o processo priorizado.", "RES 031/2025"),
 ("Gestor do Processo", "Titular da unidade orgânica responsável pela execução do processo ou, na sua ausência, o demandante do processo. Aprova os processos mapeados pelo ponto focal, forma a equipe de gerenciamento e responde pelos resultados perante a UNP.", "RES 031/2025"),
 ("Hierarquia de Processos", "Forma de visualizar como os processos se desdobram do nível mais alto — a cadeia de valor — até o mais baixo — as tarefas.", "RES 031/2025"),
 ("Plano de Gerenciamento do Processo (PGP)", "Documento da etapa de planejamento, com o cronograma das etapas de conhecimento, análise, transformação, gerenciamento de desempenho, monitoramento e reavaliação do processo.", "RES 031/2025"),
 ("Ponto Focal do Nugep", "Empregado que conduz e coordena os trabalhos de gerenciamento de processos na própria unidade orgânica, como agente multiplicador do Nugep.", "RES 031/2025"),
 ("Procedimento (PRO)", "Documento que reúne objetivo, diagramas, referências legais, classificação na cadeia de valor, unidade responsável, sistemas, produtos, atores, tarefas, regras de negócio, cronograma e indicadores do processo.", "RES 031/2025"),
 ("Regras de Negócio", "Premissas e restrições que garantem o funcionamento adequado da organização: definem o que, onde, por que e como algo será feito, e como será gerenciado e governado.", "RES 031/2025"),
 ("Unidade Orgânica", "Toda e qualquer unidade com representação formal na estrutura orgânica da Empresa.", "RES 031/2025"),
 # Guia de Modelagem de Processos (RES 031/2025) — completa a distinção de
 # níveis de representação junto de "Diagrama", já no Glossário.
 ("Mapa", "Nível intermediário de detalhamento de um processo: evolução do diagrama que soma atores, eventos, regras e resultados. Usado para representar subprocessos, de preferência pelo “caminho feliz” — as exceções ficam na Tabela de Descrição das Atividades.", "RES 031/2025"),
 ("Modelo", "Nível mais analítico de detalhamento de um processo, usado quando o mapa não é suficiente. Tem alto grau de precisão, mas exige cuidado para não poluir a leitura.", "RES 031/2025"),
]

# Competências e atribuições (RES 031/2025, item 3) — exibidas na aba NUGEP
# do painel, seção "Competências e atribuições" (accordion).
COMPETENCIAS = [
 (1, "Conselho de Administração", "3.1",
  "Garantir o apoio institucional para a gestão de processos."),
 (2, "Diretoria Executiva", "3.2",
  "Definir diretrizes para a gestão de processos; "
  "Aprovar a metodologia de gestão de processos; "
  "Aprovar a priorização para a gestão de processos; "
  "Aprovar os indicadores e metas de desempenho dos processos; "
  "Avaliar a aferição dos indicadores dos processos."),
 (3, "Unidade de Gestão Normativa e de Processos (AE/GPE/UNP)", "3.3",
  "Promover e supervisionar a gestão de processos na Codevasf; "
  "Coordenar o Núcleo de Gestão Normativa e de Processos; "
  "Sistematizar, padronizar e difundir princípios, práticas e padrões de gestão de processos; "
  "Elaborar e manter atualizada a metodologia de gestão de processos; "
  "Fornecer orientação e treinamento sobre o gerenciamento de processos; "
  "Difundir a cultura de gestão de processos; "
  "Gerenciar a arquitetura e o repositório de processos; "
  "Consolidar informações por meio de relatórios gerenciais."),
 (4, "Núcleo de Gestão Normativa e de Processos (Nugep)", "3.4",
  "Propor as diretrizes sobre a gestão de processos; "
  "Auxiliar no acompanhamento dos resultados dos processos e na proposição de correções e melhorias; "
  "Avaliar e propor melhorias na metodologia de gestão de processos; "
  "Avaliar e propor melhorias para elaboração de indicadores e metas de desempenho; "
  "Avaliar e propor melhorias nos artefatos e documentos relacionados à gestão de processos; "
  "Fomentar e promover a gestão de processos em suas unidades organizacionais; "
  "Propor a priorização para o gerenciamento de processos."),
 (5, "Ponto focal do Nugep", "3.5",
  "Conduzir e coordenar os trabalhos de gerenciamento de processos no âmbito de sua unidade orgânica; "
  "Conduzir as oficinas de trabalho para levantamento, análise, coleta de informações e proposição de melhorias; "
  "Planejar a implementação, o monitoramento e a avaliação dos processos mapeados; "
  "Atuar como fornecedor de informações técnicas específicas, mesmo em processo fora de sua unidade de lotação; "
  "Atuar como agente multiplicador e facilitador da AE/GPE/UNP; "
  "Ter perfil de liderança e conhecimento em gestão de processos; "
  "Ter prioridade nas capacitações relacionadas ao tema."),
 (6, "Gestor do processo", "3.6",
  "Aprovar os processos de trabalho mapeados pelos respectivos pontos focais do Nugep; "
  "Formar equipe de gerenciamento de processos para gerenciamento do processo; "
  "Engajar os atores do processo nos trabalhos de gerenciamento do processo; "
  "Gerenciar e monitorar os processos sob sua responsabilidade; "
  "Reportar os resultados dos processos à AE/GPE/UNP; "
  "Acompanhar os trabalhos de gerenciamento dos processos sob sua responsabilidade; "
  "Elaborar, monitorar e prestar informações sobre os indicadores de desempenho dos processos; "
  "Assegurar que o processo atenda às expectativas de desempenho estabelecidas; "
  "Propor melhorias ou inovações, com vistas a tornar os processos/subprocessos eficientes, eficazes e efetivos; "
  "Disseminar os processos/subprocessos mapeados dentro da respectiva unidade orgânica."),
 (7, "Equipe de Gerenciamento de Processo", "3.7",
  "Realizar o gerenciamento dos processos priorizados de sua competência."),
]

_FC, _FM, _FG, _FI, _FP, _FU = ("Conceitos básicos", "Modelagem e SIPOC",
                                "Cadeia de Valor e governança", "Indicadores, metas e riscos",
                                "Plano de Ações AE/GPE", "Como usar o painel")

# (Ordem, Categoria, Pergunta, Resposta)
FAQ = [
 (1, _FC, "O que é Gestão de Processos (BPM)?",
  "Business Process Management é uma disciplina gerencial integrada que combina técnicas, métodos e ferramentas para identificar, modelar, executar, monitorar e melhorar continuamente os processos de negócio, alinhando-os à estratégia e à entrega de valor (CBOK 4.0)."),
 (2, _FC, "Qual a diferença entre Gestão DE Processos e Gestão POR Processos?",
  "Gestão DE processos é o trabalho de melhorar processos individualmente: mapear, analisar, redesenhar e medir cada um. Gestão POR processos é o modelo de gestão da organização orientado a processos ponta a ponta — com donos, indicadores e decisões estruturadas pelos fluxos de valor, e não pelos silos funcionais. Este painel evidencia os dois: o repositório de processos (DE) e a cadeia de valor com governança (POR)."),
 (3, _FC, "Qual a diferença entre macroprocesso, processo, subprocesso, atividade e tarefa?",
  "São os níveis da arquitetura de processos (CBOK 4.0): o macroprocesso agrupa grandes conjuntos de processos da cadeia de valor; o processo entrega um resultado a um cliente interno ou externo; o subprocesso é um recorte gerenciável do processo; a atividade é o nível operacional modelado em BPMN; e a tarefa é a menor unidade de trabalho dentro de uma atividade."),
 (4, _FC, "O que é um processo finalístico, gerencial e de suporte?",
  "Finalísticos (ou primários) entregam valor diretamente ao beneficiário — na Codevasf, irrigação, revitalização e desenvolvimento territorial. De suporte habilitam os demais (licitações, pessoas, TI). Gerenciais medem, controlam e direcionam a organização (estratégia, riscos e governança)."),
 (5, _FM, "O que é SIPOC e como ler a ficha de um processo?",
  "SIPOC é a visão de alto nível do processo em cinco colunas: Fornecedores, Entradas, Processo, Saídas e Beneficiários. Na ficha de cada processo deste painel, o SIPOC aparece logo após os dados gerais, seguido dos marcos M1–M10, indicadores, riscos e documentos vinculados."),
 (6, _FM, "O que é AS-IS e TO-BE?",
  "AS-IS é o retrato do processo como ele é executado hoje, usado para diagnóstico. TO-BE é o desenho do processo futuro, após a análise de melhorias. Os dois modelos são validados formalmente com o dono do processo antes da publicação (marcos M5 e M8)."),
 (7, _FM, "Que notação a Codevasf usa para modelar processos?",
  "BPMN 2.0, modelada preferencialmente no Bizagi Modeler, conforme o Guia de Gerenciamento de Processos (RES 031/2025). Os diagramas publicados aparecem na ficha de cada macroprocesso, processo e subprocesso."),
 (8, _FM, "Como propor uma mudança em um processo já mapeado?",
  "Use o botão flutuante “Sugerir melhoria de processo” (que abre um e-mail para a UNP) ou procure o dono do processo indicado na ficha. Propostas aceitas abrem um novo ciclo de análise e redesenho, registrado no processo e-Codevasf correspondente."),
 (9, _FG, "O que é a Cadeia de Valor Integrada da Codevasf?",
  "É a representação dos macroprocessos da Companhia organizados por contribuição ao valor público: gerenciais no direcionamento, finalísticos na entrega à sociedade e de suporte na sustentação — emoldurados pela missão, visão, propósito e valores. Ela é a porta de entrada de navegação deste painel."),
 (10, _FG, "Quem é o dono do processo?",
  "É a pessoa formalmente responsável pelo desempenho, pela integridade e pela melhoria contínua do processo de ponta a ponta, independentemente da hierarquia (CBOK 4.0). O dono está identificado na ficha de cada processo."),
 (11, _FG, "Qual o papel da UNP, da GPE e da AE?",
  "A UNP é o escritório de processos: conduz a metodologia, mantém o repositório e apoia as áreas. Ela integra a GPE (Gerência de Planejamento Estratégico), que por sua vez se vincula à AE (Área de Estratégia e Finanças), instância superior de governança."),
 (12, _FG, "Quem aprova as mudanças nos processos?",
  "O dono do processo valida os modelos (AS-IS e TO-BE) e as instâncias de governança definidas na RES 031/2025 aprovam a publicação, com apoio metodológico da UNP. Mudanças que alteram normativos seguem também o fluxo do Sistema Normativo."),
 (13, _FI, "O que é um KPI?",
  "Key Performance Indicator: um indicador-chave de desempenho, com fórmula, fonte, meta, periodicidade e responsável definidos. A aba Indicadores lista os KPIs por nível (macroprocesso, processo, subprocesso) com a situação calculada em relação à meta."),
 (14, _FI, "Qual a diferença entre eficiência, eficácia e efetividade?",
  "Eficiência é fazer certo as coisas (relação produto/recursos); eficácia é fazer as coisas certas (alcançar o resultado pretendido); efetividade é o impacto real e duradouro no público-alvo — a combinação das duas com sustentabilidade."),
 (15, _FI, "Como os riscos são classificados?",
  "Pela matriz 5×5 de probabilidade × impacto: o produto P×I gera o nível — Baixo (até 4), Moderado (5 a 11), Alto (12 a 19) e Extremo (20 a 25). A aba Riscos mostra a matriz com os riscos posicionados e a tabela com resposta, controles e status."),
 (16, _FI, "Onde vejo os riscos do projeto de mapeamento?",
  "Na aba Riscos (visão geral, com matriz e tabela) e na ficha de cada item da hierarquia, que lista os riscos vinculados àquele macroprocesso, processo, subprocesso ou atividade."),
 (17, _FP, "O que é o Plano de Ações AE/GPE 2026?",
  "É o plano conduzido pela AE/GPE/UNP para estruturar o escritório de processos, mapear os processos prioritários e implantar a governança de processos na Codevasf. Os projetos de mapeamento deste painel derivam dele. (Nesta demonstração, os dados são fictícios.)"),
 (18, _FP, "Quem é a responsável geral pelo plano?",
  "A coordenação do NUGEP — Núcleo de Gestão Normativa e de Processos (ver aba NUGEP), em articulação com a GPE e a AE."),
 (19, _FP, "Como o avanço das ações é medido?",
  "Pelos marcos M1–M10 de cada mapeamento, pelo percentual de execução e pelos indicadores da aba Indicadores."),
 (20, _FP, "O painel usa PMBOK ou CBOK?",
  "Os dois, com papéis distintos: o CBOK 4.0 orienta o conteúdo de processos (hierarquia, SIPOC, ciclo de vida BPM, indicadores); o PMBOK orienta a gestão de cada mapeamento como projeto (termo de abertura, escopo, marcos, riscos e lições aprendidas)."),
 (21, _FU, "Como pesquiso um processo?",
  "Use a busca no topo da página (digite e pressione Enter) para pesquisar em processos, documentos, glossário e registros — ou navegue pela aba Processos, filtrando por macroprocesso, status e texto."),
 (22, _FU, "Como sugerir uma melhoria ou reportar um erro?",
  "Pelo botão flutuante “Sugerir melhoria de processo”, que abre um e-mail para a UNP (ae.gpe.unp@codevasf.gov.br). Indique o código do processo para agilizar a análise."),
 (23, _FU, "Como o painel é alimentado?",
  "Por uma planilha única (Google Sheets ou o arquivo data/painel-processos-dados.xlsx do repositório): cadeia de valor, fichas, documentos, riscos, indicadores, jornada, repositório, NUGEP, glossário e FAQ vêm todos de lá. Os detalhes estão na aba LEIA-ME da planilha e no README."),
 (24, _FU, "Onde encontro os modelos, a metodologia e o guia oficiais?",
  "Na aba Repositório: a jornada de mapeamento, a Metodologia e o Guia de Gerenciamento de Processos (RES 031/2025, na intranet/e-Codevasf), os instrumentos por fase do ciclo BPM, os templates e as ferramentas de modelagem."),
]

# (Chave, Valor)
# Textos e links institucionais editáveis pela planilha, sem mexer no
# código do painel. Toda chave nova aqui precisa ser lida por par() em
# js/app.js, que cai no texto atual quando a chave não existe.
PARAMETROS = [
 ("Titulo_Inicio", "Mapeamento de processos da Codevasf"),
 ("Subtitulo_Inicio",
  "Cadeia de valor, portfólio de processos e repositório de materiais "
  "da Companhia, num só lugar."),
 ("Titulo_Repositorio", "Repositório de materiais e ferramentas"),
 ("Contato_Unidade", "Unidade de Gestão Normativa e de Processos (AE/GPE/UNP)"),
 ("Contato_Email", "ae.gpe.unp@codevasf.gov.br"),
 ("Contato_Telefone", "(61) 2028-4441"),
 ("Link_Metodologia", _MET),
 ("Link_Guia", _GUIA),
]
