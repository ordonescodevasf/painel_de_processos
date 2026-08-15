# -*- coding: utf-8 -*-
"""
Atualiza data/painel-processos-dados.xlsx NO LUGAR, sem recriar a planilha:
soma a coluna M10_Processo_Transformado, a etapa "Análise do processo" da
Jornada, os termos de RES 031/2025 no Glossário e as abas PlanoAcao (5W2H) e
Competencias da auditoria contra a Metodologia, e renomeia "Interlocutor"
para "Ponto Focal do Nugep" na aba NUGEP — preservando tudo o mais: dados já
preenchidos, formatação, fórmulas e validações.

Rode com:
    python scripts/atualizar_planilha.py

Idempotente: rodar de novo não duplica nada — cada mudança confere se já foi
aplicada antes de agir. Reaplique sempre que uma auditoria de metodologia
adicionar algo novo à planilha (em vez de gerar_planilha.py, que RECRIA a
planilha de exemplo do zero e por isso não serve para a planilha em uso).
"""
import os
import re
import datetime as dt
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

CAMINHO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                       "data", "painel-processos-dados.xlsx")

# Mesmas constantes visuais de gerar_planilha.py — mantenha as duas em sincronia.
AZUL = "0040B4"
CINZA_ZEBRA = "F5F8FC"
FONTE = "Arial"
th = Side(style="thin", color="CCCCCC")
BORDA = Border(left=th, right=th, top=th, bottom=th)
F_HEAD = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
F_CELL = Font(name=FONTE, size=10)
FILL_HEAD = PatternFill("solid", fgColor=AZUL)
FILL_ZEBRA = PatternFill("solid", fgColor=CINZA_ZEBRA)
AL_HEAD = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_WRAP = Alignment(vertical="top", wrap_text=True)
AL_TOP = Alignment(vertical="top")
AL_CENTER = Alignment(horizontal="center", vertical="top")


def cabecalho(ws, headers, widths, tab_color=AZUL):
    for j, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = 30
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.zoomScale = 90


def escreve(ws, linhas, wrap_cols=(), center_cols=()):
    for i, linha in enumerate(linhas, start=2):
        for j, v in enumerate(linha, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = F_CELL; c.border = BORDA
            if i % 2 == 0:
                c.fill = FILL_ZEBRA
            c.alignment = AL_WRAP if j in wrap_cols else (AL_CENTER if j in center_cols else AL_TOP)


def col_by_header(ws, nome, linha=1):
    """Índice (1-based) da coluna cujo cabeçalho é `nome`, ou None."""
    for cell in ws[linha]:
        if cell.value == nome:
            return cell.column
    return None


def listas_ref(wb, nome_lista):
    """Referência de validação para uma coluna nomeada da aba Listas, ou None
    se a aba Listas não tiver essa coluna (planilha mais antiga)."""
    if "Listas" not in wb.sheetnames:
        return None
    ls = wb["Listas"]
    col = col_by_header(ls, nome_lista)
    if not col:
        return None
    n = 1
    while ls.cell(row=n + 1, column=col).value not in (None, ""):
        n += 1
    letra = get_column_letter(col)
    return f"=Listas!${letra}$2:${letra}${n + 1}"


def dv(ws, col_letter, ref, ultima_linha=300):
    if not ref:
        return
    v = DataValidation(type="list", formula1=ref, allow_blank=True,
                        showErrorMessage=True, errorStyle="stop",
                        errorTitle="Valor fora da lista",
                        error="Escolha uma das opções da lista. A lista completa está na aba Listas.",
                        showInputMessage=True, promptTitle="Lista de opções",
                        prompt="Clique na seta ao lado da célula e escolha uma das opções.")
    ws.add_data_validation(v)
    v.add(f"{col_letter}2:{col_letter}{ultima_linha}")


def principal():
    wb = load_workbook(CAMINHO)
    mudou = []

    # 1) Processos.M10_Processo_Transformado — fecha o décimo marco (antes,
    # nenhuma coluna real correspondia a ele: ver DESIGN-SYSTEM.md).
    pr = wb["Processos"]
    if not col_by_header(pr, "M10_Processo_Transformado"):
        col = pr.max_column + 1
        c = pr.cell(row=1, column=col, value="M10_Processo_Transformado")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        pr.column_dimensions[get_column_letter(col)].width = 9
        # M1..M9 têm nomes de coluna variados (M6_Procedimento_Validado,
        # M7_Procedimento_Aprovado...) — acha pelo prefixo "M{n}_", não pelo
        # nome inteiro.
        m_cols = []
        for n in range(1, 10):
            achou = None
            for cell in pr[1]:
                if str(cell.value or "").startswith(f"M{n}_"):
                    achou = cell.column
                    break
            m_cols.append(achou)
        for r in range(2, pr.max_row + 1):
            if not pr.cell(row=r, column=1).value:
                continue
            tem_nao = any(
                mc and str(pr.cell(row=r, column=mc).value or "").strip() == "Não"
                for mc in m_cols
            )
            cc = pr.cell(row=r, column=col, value=("Não" if tem_nao else "Sim"))
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_CENTER
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        dv(pr, get_column_letter(col), listas_ref(wb, "Sim_Nao") or '"Sim,Não,Não se aplica"')
        mudou.append("Processos: coluna M10_Processo_Transformado (valor inicial calculado de M1-M9; confira e ajuste manualmente)")

    # 2) Jornada: etapa "Análise do processo" (RES 031/2025, item 5.5),
    # entre "Validação AS-IS" e "Redesenho TO-BE".
    jn = wb["Jornada"]
    ja_tem = any(jn.cell(row=r, column=3).value == "Análise do processo" for r in range(2, jn.max_row + 1))
    if not ja_tem:
        linha_alvo = None
        for r in range(2, jn.max_row + 1):
            try:
                if int(str(jn.cell(row=r, column=1).value).strip()) == 5:
                    linha_alvo = r
                    break
            except (ValueError, TypeError):
                continue
        if linha_alvo:
            jn.insert_rows(linha_alvo)
            valores = ["5", "Desenvolver", "Análise do processo", "1–2 semanas",
                "Investigar as causas dos problemas do AS-IS validado e preparar o plano de melhorias, antes de redesenhar o processo.",
                "Entrevista estruturada por componente habilitador (pessoas, fluxo de trabalho, TI, regras/políticas, métricas, infraestrutura, colaboração); Identificação de causas e consequências dos problemas; Elaboração de propostas de melhoria; Plano de Ação pela técnica 5W2H",
                "Equipe de gerenciamento de processo; Dono do processo",
                "Lista de problemas e causas; Plano de Ação (5W2H)",
                "Alívio: “agora vamos atacar o que realmente trava o meu trabalho”"]
            for j, v in enumerate(valores, start=1):
                c = jn.cell(row=linha_alvo, column=j, value=v)
                c.font = F_CELL; c.border = BORDA
                c.alignment = AL_WRAP if j in (5, 6, 7, 8, 9) else AL_TOP
            for r in range(linha_alvo + 1, jn.max_row + 1):
                try:
                    n = int(str(jn.cell(row=r, column=1).value).strip())
                except (ValueError, TypeError):
                    continue
                if n >= 5:
                    jn.cell(row=r, column=1).value = str(n + 1)
            for r in range(2, jn.max_row + 1):
                for j in range(1, 10):
                    jn.cell(row=r, column=j).fill = FILL_ZEBRA if r % 2 == 0 else PatternFill(fill_type=None)
            mudou.append("Jornada: nova etapa 'Análise do processo' (Ordem 5; etapas seguintes renumeradas)")

    # 3) Glossario: troca "Interlocutor de Processos" por "Ponto Focal do
    # Nugep" e soma os termos da RES 031/2025 ainda ausentes.
    gl = wb["Glossario"]
    linha_del = None
    for r in range(2, gl.max_row + 1):
        if gl.cell(row=r, column=1).value == "Interlocutor de Processos":
            linha_del = r
            break
    if linha_del:
        gl.delete_rows(linha_del)
        mudou.append("Glossario: removido 'Interlocutor de Processos' (papel passa a se chamar Ponto Focal do Nugep)")
    termos_atuais = {gl.cell(row=r, column=1).value for r in range(2, gl.max_row + 1)}
    novos_termos = [
        ("5W2H", "Projetos (PMBOK)", "Checklist de atividades, prazos e responsabilidades usado no plano de ação da etapa de análise: o que, por que, onde, quando, por quem, como e quanto vai custar.", "RES 031/2025", "Plano de Gerenciamento do Processo (PGP)"),
        ("Ator do Processo", "Metodologia Codevasf", "Empregado que participa, em algum momento, da execução do processo.", "RES 031/2025", "Gestor do Processo; Equipe de Gerenciamento de Processo"),
        ("Diagrama", "BPM (CBOK)", "Representação gráfica que demonstra os principais elementos do fluxo do processo, ajudando a identificar e entender rapidamente suas atividades.", "RES 031/2025", "BPMN; Fluxograma"),
        ("Equipe de Gerenciamento de Processo", "Metodologia Codevasf", "Grupo formado pelo ponto focal do Nugep, pelo gestor do processo e por atores do processo, responsável por gerenciar o processo priorizado.", "RES 031/2025", "Ponto Focal do Nugep; Gestor do Processo; Ator do Processo"),
        ("Gestor do Processo", "Metodologia Codevasf", "Titular da unidade orgânica responsável pela execução do processo ou, na sua ausência, o demandante do processo. Aprova os processos mapeados pelo ponto focal, forma a equipe de gerenciamento e responde pelos resultados perante a UNP.", "RES 031/2025", "Dono do Processo; Unidade Orgânica"),
        ("Hierarquia de Processos", "BPM (CBOK)", "Forma de visualizar como os processos se desdobram do nível mais alto — a cadeia de valor — até o mais baixo — as tarefas.", "RES 031/2025", "Arquitetura de Processos; Cadeia de Valor"),
        ("Plano de Gerenciamento do Processo (PGP)", "Metodologia Codevasf", "Documento da etapa de planejamento, com o cronograma das etapas de conhecimento, análise, transformação, gerenciamento de desempenho, monitoramento e reavaliação do processo.", "RES 031/2025", "Equipe de Gerenciamento de Processo; Cronograma"),
        ("Ponto Focal do Nugep", "Metodologia Codevasf", "Empregado que conduz e coordena os trabalhos de gerenciamento de processos na própria unidade orgânica, como agente multiplicador do Nugep.", "RES 031/2025", "Equipe de Gerenciamento de Processo"),
        ("Procedimento (PRO)", "BPM (CBOK)", "Documento que reúne objetivo, diagramas, referências legais, classificação na cadeia de valor, unidade responsável, sistemas, produtos, atores, tarefas, regras de negócio, cronograma e indicadores do processo.", "RES 031/2025", "Regras de Negócio; Diagrama"),
        ("Regras de Negócio", "BPM (CBOK)", "Premissas e restrições que garantem o funcionamento adequado da organização: definem o que, onde, por que e como algo será feito, e como será gerenciado e governado.", "RES 031/2025", "Procedimento (PRO)"),
        ("Unidade Orgânica", "Metodologia Codevasf", "Toda e qualquer unidade com representação formal na estrutura orgânica da Empresa.", "RES 031/2025", "Gestor do Processo; Unidade Gestora"),
        # Guia de Modelagem de Processos (RES 031/2025) — completa a distinção
        # de níveis de representação junto de "Diagrama", já no Glossário.
        ("Mapa", "BPM (CBOK)", "Nível intermediário de detalhamento de um processo: evolução do diagrama que soma atores, eventos, regras e resultados. Usado para representar subprocessos, de preferência pelo \u201ccaminho feliz\u201d — as exceções ficam na Tabela de Descrição das Atividades.", "RES 031/2025", "Diagrama; Modelo"),
        ("Modelo", "BPM (CBOK)", "Nível mais analítico de detalhamento de um processo, usado quando o mapa não é suficiente. Tem alto grau de precisão, mas exige cuidado para não poluir a leitura.", "RES 031/2025", "Diagrama; Mapa"),
    ]
    adicionados = 0
    for termo in novos_termos:
        if termo[0] in termos_atuais:
            continue
        r = gl.max_row + 1
        for j, v in enumerate(termo, start=1):
            c = gl.cell(row=r, column=j, value=v)
            c.font = F_CELL; c.border = BORDA
            c.alignment = AL_WRAP if j in (3, 5) else AL_TOP
            if r % 2 == 0:
                c.fill = FILL_ZEBRA
        adicionados += 1
    if adicionados:
        mudou.append(f"Glossario: {adicionados} termo(s) novo(s) da RES 031/2025")

    # 4) NUGEP: "Interlocutor(a) de Processos..." -> "Ponto Focal do Nugep —
    # ...", mantendo o recorte Finalísticos/Suporte já usado.
    ng = wb["NUGEP"]
    col_papel = col_by_header(ng, "Papel")
    renomeados = 0
    if col_papel:
        for r in range(2, ng.max_row + 1):
            papel = str(ng.cell(row=r, column=col_papel).value or "")
            if not papel.startswith("Interlocutor"):
                continue
            if "Finalístico" in papel:
                novo = "Ponto Focal do Nugep — Processos Finalísticos"
            elif "Suporte" in papel:
                novo = "Ponto Focal do Nugep — Processos de Suporte"
            else:
                novo = "Ponto Focal do Nugep"
            ng.cell(row=r, column=col_papel).value = novo
            renomeados += 1
    if renomeados:
        mudou.append(f"NUGEP: {renomeados} papel(éis) renomeado(s) para 'Ponto Focal do Nugep'")

    # 4.1) Subprocessos.Produto / Cronograma_Proposto_Dias e
    # Atividades.Responsavel (Guia de Modelagem, Anexos A e B).
    sp = wb["Subprocessos"]
    if not col_by_header(sp, "Produto"):
        col = sp.max_column + 1
        c = sp.cell(row=1, column=col, value="Produto")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        sp.column_dimensions[get_column_letter(col)].width = 34
        for r in range(2, sp.max_row + 1):
            if not sp.cell(row=r, column=1).value:
                continue
            cc = sp.cell(row=r, column=col, value="A definir")
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_WRAP
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        mudou.append("Subprocessos: coluna Produto (Guia, Anexo B) — preenchida com 'A definir', confira o produto principal de cada subprocesso")
    if not col_by_header(sp, "Cronograma_Proposto_Dias"):
        col = sp.max_column + 1
        c = sp.cell(row=1, column=col, value="Cronograma_Proposto_Dias")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        sp.column_dimensions[get_column_letter(col)].width = 15
        for r in range(2, sp.max_row + 1):
            if not sp.cell(row=r, column=1).value:
                continue
            cc = sp.cell(row=r, column=col)
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_CENTER
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        mudou.append("Subprocessos: coluna Cronograma_Proposto_Dias (Guia, Anexo B), em branco — estimativa em dias úteis, distinta da duração em horas")

    at = wb["Atividades"]
    if not col_by_header(at, "Responsavel"):
        col = at.max_column + 1
        c = at.cell(row=1, column=col, value="Responsavel")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        at.column_dimensions[get_column_letter(col)].width = 24
        for r in range(2, at.max_row + 1):
            if not at.cell(row=r, column=1).value:
                continue
            cc = at.cell(row=r, column=col)
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_TOP
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        mudou.append("Atividades: coluna Responsavel (Guia, Anexo A), opcional e em branco — sem valor, o painel continua herdando a unidade do subprocesso/processo")

    # 4.2) Glossario: cruza "Diagrama" com os termos novos "Mapa"/"Modelo".
    col_relacionados = 5
    for r in range(2, gl.max_row + 1):
        if gl.cell(row=r, column=1).value == "Diagrama":
            atual = str(gl.cell(row=r, column=col_relacionados).value or "")
            if "Mapa" not in atual:
                gl.cell(row=r, column=col_relacionados).value = (atual + "; Mapa; Modelo") if atual else "Mapa; Modelo"
                mudou.append("Glossario: 'Diagrama' cruzado com 'Mapa' e 'Modelo'")
            break

    # 4.3) Documentos.Ato_Aprovacao e Atividades.Descricao (N-000 — Norma de
    # Gestão dos Instrumentos Normativos: o cabeçalho-padrão de todo
    # instrumento normativo registra o ato que o aprovou; o Anexo J exige
    # título E descrição detalhada de cada atividade no quadro "Descrição
    # dos Procedimentos").
    dc = wb["Documentos"]
    if not col_by_header(dc, "Ato_Aprovacao"):
        col = dc.max_column + 1
        c = dc.cell(row=1, column=col, value="Ato_Aprovacao")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        dc.column_dimensions[get_column_letter(col)].width = 30
        # Só Norma interna/Procedimento (PRO)/Manual são aprovados por
        # resolução (N-000, item 3.1.1); Formulários e os demais tipos
        # (ata, diagrama, relatório...) ficam em branco.
        atos = {
            "DOC-001": "Resolução nº 812, de 18/05/2026 (Diretoria Executiva)",
            "DOC-007": "Resolução nº 845, de 27/04/2026 (Diretor da Área de Aquisições)",
            "DOC-008": "Resolução nº 621, de 20/11/2024 (Diretoria Executiva)",
            "DOC-009": "Resolução nº 703, de 14/08/2025 (Diretoria Executiva)",
        }
        for r in range(2, dc.max_row + 1):
            id_doc = dc.cell(row=r, column=1).value
            if not id_doc:
                continue
            cc = dc.cell(row=r, column=col, value=atos.get(id_doc))
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_WRAP
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        mudou.append("Documentos: coluna Ato_Aprovacao (N-000) — preenchida para as Normas/Procedimentos/Manuais existentes; em branco nos demais tipos")

    if not col_by_header(at, "Descricao"):
        col = at.max_column + 1
        c = at.cell(row=1, column=col, value="Descricao")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        at.column_dimensions[get_column_letter(col)].width = 46
        descricoes = {
            "A-06.01.01.01": "A área demandante registra a necessidade de contratação no Documento de Formalização da Demanda (DFD), justificando o objeto e vinculando-o ao Plano de Contratações Anual.",
            "A-06.01.01.02": "A equipe de planejamento pesquisa catálogos, contratações similares no PNCP e alternativas de mercado que atendam à necessidade formalizada no DFD.",
            "A-06.01.01.03": "A equipe de planejamento estima quantitativos e resultados pretendidos com base no levantamento de soluções e em séries históricas de consumo.",
            "A-06.01.01.04": "A equipe de planejamento consolida os levantamentos no Estudo Técnico Preliminar (ETP) e o submete à aprovação da autoridade competente.",
            "A-06.01.02.01": "A equipe de planejamento redige a minuta do Termo de Referência a partir do ETP aprovado, utilizando os modelos padronizados da Empresa.",
            "A-06.01.02.02": "A Assessoria Jurídica analisa a minuta do Termo de Referência e emite parecer sobre sua conformidade legal antes da publicação do edital.",
            "A-06.01.03.01": "A equipe de planejamento coleta preços em ao menos três fontes distintas (Painel de Preços, PNCP e contratações similares) para compor a estimativa de valor.",
            "A-06.01.03.02": "A equipe de planejamento consolida os preços coletados e tratados estatisticamente em um valor estimado único para a contratação.",
            "A-06.01.03.01.01": "A equipe de planejamento aplica o critério estatístico definido para excluir da amostra valores inexequíveis ou excessivamente discrepantes, registrando a justificativa.",
            "A-06.02.01.01": "A área responsável publica o edital aprovado e o parecer jurídico no Compras.gov.br e no PNCP, abrindo o certame.",
            "A-06.02.01.02": "O agente de contratação conduz a sessão pública, analisa as propostas recebidas e registra o resultado do julgamento em ata.",
            "A-04.01.01.01": "A equipe de operação do perímetro programa semanalmente a distribuição de água com base no plano de cultivo e na disponibilidade hídrica.",
            "A-04.01.01.02": "A equipe de operação executa a distribuição programada e registra volumes e ocorrências no sistema de gestão da irrigação.",
        }
        for r in range(2, at.max_row + 1):
            codigo = at.cell(row=r, column=1).value
            if not codigo:
                continue
            cc = at.cell(row=r, column=col, value=descricoes.get(codigo))
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_WRAP
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        mudou.append("Atividades: coluna Descricao (N-000, Anexo J) — descrição detalhada de cada atividade existente")

    # 4.4) Documentos: exemplo de norma com revisão vencida, para o selo do
    # item 4.2.3 (revisão trienal) aparecer em algum lugar da planilha de
    # demonstração — nenhum documento real até aqui estava vencido.
    if "DOC-016" not in {dc.cell(row=r, column=1).value for r in range(2, dc.max_row + 1)}:
        r = dc.max_row + 1
        valores = {
            1: "DOC-016", 3: "MP-07", 4: "Norma interna",
            5: "NI-014/2019 — Norma de Administração de Pessoal (fictícia)",
            6: "3.0", 7: dt.date(2019, 6, 10), 8: "Vigente",
            9: "https://exemplo.codevasf.gov.br/normativos/ni-014",
            10: "Exemplo — norma com revisão trienal vencida (N-000, item 4.2.3), ainda não atualizada.",
            11: "Resolução nº 512, de 10/06/2019 (Diretoria Executiva)",
        }
        for j, v in valores.items():
            c = dc.cell(row=r, column=j, value=v)
            c.font = F_CELL; c.border = BORDA
            c.alignment = AL_WRAP if j in (5, 9, 10, 11) else (AL_CENTER if j in (6, 7, 8) else AL_TOP)
            if r % 2 == 0:
                c.fill = FILL_ZEBRA
        formula2 = dc.cell(row=2, column=2).value
        nova_formula = re.sub(r"([A-Z]+)2\b", lambda m: m.group(1) + str(r), formula2) if isinstance(formula2, str) else None
        c2 = dc.cell(row=r, column=2, value=nova_formula)
        c2.font = F_CELL; c2.border = BORDA; c2.alignment = AL_CENTER
        if r % 2 == 0:
            c2.fill = FILL_ZEBRA
        mudou.append("Documentos: exemplo DOC-016 (Norma interna de 2019, revisão vencida) para demonstrar o selo do item 4.2.3")

    # 4.5) Preenche Subprocessos.Produto/.Cronograma_Proposto_Dias com
    # valores reais — estavam "A definir"/em branco desde que as colunas
    # foram criadas (item 4.1 acima).
    produtos_sp = {
        "SP-06.01.01": ("Estudo Técnico Preliminar (ETP) aprovado", 10),
        "SP-06.01.02": ("Termo de Referência (TR) aprovado e validado juridicamente", 8),
        "SP-06.01.03": ("Relatório de pesquisa de preços com valor estimado consolidado", 12),
        "SP-06.01.03.01": ("Amostra de preços tratada estatisticamente, com justificativa de exclusões", 2),
        "SP-06.02.01": ("Ata da sessão pública com resultado do certame por item", 5),
        "SP-04.01.01": ("Programação hídrica semanal executada e registrada", 5),
    }
    col_prod = col_by_header(sp, "Produto")
    col_cron = col_by_header(sp, "Cronograma_Proposto_Dias")
    if col_prod and col_cron:
        atualizados = 0
        for r in range(2, sp.max_row + 1):
            codigo = sp.cell(row=r, column=1).value
            if codigo not in produtos_sp:
                continue
            if sp.cell(row=r, column=col_prod).value in (None, "", "A definir"):
                produto, dias = produtos_sp[codigo]
                sp.cell(row=r, column=col_prod).value = produto
                sp.cell(row=r, column=col_cron).value = dias
                atualizados += 1
        if atualizados:
            mudou.append(f"Subprocessos: Produto/Cronograma_Proposto_Dias preenchidos em {atualizados} subprocesso(s)")

    # 4.6) Processos.Competencias_Necessarias (Guia de Modelagem, Anexo C —
    # Tabela de Catalogação do Processo): competências/habilidades de RH
    # necessárias para executar o processo, distinto da aba Competencias
    # (atribuições de governança, item 3 da Metodologia).
    if not col_by_header(pr, "Competencias_Necessarias"):
        col = pr.max_column + 1
        c = pr.cell(row=1, column=col, value="Competencias_Necessarias")
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = AL_HEAD; c.border = BORDA
        pr.column_dimensions[get_column_letter(col)].width = 40
        competencias_proc = {
            "P-06.01": "Elaboração de estudos técnicos e termos de referência; Conhecimento da Lei nº 14.133/2021; Pesquisa e análise de preços de mercado",
            "P-06.02": "Condução de sessões públicas de licitação; Julgamento de propostas e habilitação; Uso de plataformas de compras governamentais",
            "P-06.03": "Fiscalização técnica e administrativa de contratos; Gestão de medições e pagamentos; Aplicação de sanções administrativas",
            "P-04.01": "Operação de sistemas hidráulicos; Manutenção de infraestrutura de irrigação; Relacionamento com associações de irrigantes",
            "P-05.01": "Diagnóstico e recuperação de áreas degradadas; Noções de engenharia ambiental; Articulação com comunidades e órgãos ambientais",
            "P-01.01": "Planejamento estratégico e gestão por indicadores; Facilitação de oficinas; Análise de dados institucionais",
            "P-07.01": "Rotinas de admissão e legislação trabalhista (CLT); Condução de processos de integração; Uso de sistemas de gestão de pessoas",
        }
        for r in range(2, pr.max_row + 1):
            codigo = pr.cell(row=r, column=1).value
            if not codigo:
                continue
            cc = pr.cell(row=r, column=col, value=competencias_proc.get(codigo))
            cc.font = F_CELL; cc.border = BORDA; cc.alignment = AL_WRAP
            if r % 2 == 0:
                cc.fill = FILL_ZEBRA
        mudou.append("Processos: coluna Competencias_Necessarias (Guia de Modelagem, Anexo C) — competências de RH necessárias para executar cada processo")

    # 5) Nova aba PlanoAcao (5W2H — RES 031/2025, item 5.5.4).
    if "PlanoAcao" not in wb.sheetnames:
        pa = wb.create_sheet("PlanoAcao")
        cabecalho(pa,
            ["ID", "Vinculo_Nivel", "Vinculo_Codigo", "Problema", "O_Que", "Por_Que",
             "Onde", "Quando", "Quem", "Como", "Quanto_Custa", "Status"],
            [9, 15, 14, 42, 40, 34, 26, 16, 26, 40, 20, 14], tab_color="168821")
        plano_acao = [
            ["PA-001", "Processo", "P-06.01",
             "Estimativas de preço divergentes por falta de padronização na pesquisa (mesma causa-raiz do risco R-001).",
             "Padronizar o roteiro de pesquisa de preços com no mínimo três fontes e critério único de tratamento estatístico.",
             "Reduzir sobrepreço e licitações desertas causados por estimativas defasadas.",
             "Gerência de Licitações e Contratos (AA/GLC).", "30/09/2026",
             "Ricardo Nogueira (gestor do processo) e equipe de planejamento da AA/GLC",
             "Atualizar o PRO com o roteiro padronizado (DOC-007) e capacitar a equipe.",
             "Sem custo adicional — carga horária interna.", "Em andamento"],
            ["PA-002", "Processo", "P-06.01",
             "Baixo engajamento das áreas demandantes no preenchimento do DFD, gerando retrabalho.",
             "Criar modelo simplificado de DFD e oficina de capacitação para novas áreas demandantes.",
             "Reduzir devoluções por preenchimento incompleto e acelerar o início do processo.",
             "Todas as unidades demandantes, coordenado pela AA/GLC.", "15/11/2026",
             "Bruna Souza (UNP) e Carlos Eduardo Lima",
             "Oficinas trimestrais e modelo simplificado publicado no e-Codevasf.",
             "R$ 0 — recursos internos.", "Não iniciado"],
            ["PA-003", "Subprocesso", "SP-06.01.03",
             "Uso eventual de fontes de preço não admitidas pela IN SEGES nº 65/2021 (mesma causa do risco R-008).",
             "Implantar checklist de validação da pesquisa antes da aprovação do valor estimado.",
             "Garantir conformidade legal e evitar questionamentos de órgãos de controle.",
             "Equipe de planejamento da contratação (AA/GLC).", "31/08/2026",
             "Carlos Eduardo Lima",
             "Checklist obrigatório (modelo do DOC-015) antes da homologação do ETP.",
             "Sem custo — ajuste de procedimento.", "Concluído"],
        ]
        escreve(pa, plano_acao, wrap_cols={4, 5, 6, 7, 9, 10}, center_cols={8, 12})
        pa.freeze_panes = "D2"
        dv(pa, "B", listas_ref(wb, "Nivel_Vinculo") or '"Macroprocesso,Processo,Subprocesso,Atividade,Tarefa"')
        dv(pa, "L", listas_ref(wb, "Status_Mapeamento") or '"Não iniciado,Em andamento,Concluído,Suspenso"')
        mudou.append("Nova aba PlanoAcao (5W2H), com 3 exemplos ligados a P-06.01/SP-06.01.03")

    # 6) Nova aba Competencias (RES 031/2025, item 3).
    if "Competencias" not in wb.sheetnames:
        ct = wb.create_sheet("Competencias")
        cabecalho(ct, ["Ordem", "Instancia", "Item_Normativo", "Atribuicoes"],
                  [7, 46, 14, 90], tab_color="155BCB")
        competencias = [
            [1, "Conselho de Administração", "3.1", "Garantir o apoio institucional para a gestão de processos."],
            [2, "Diretoria Executiva", "3.2", "Definir diretrizes para a gestão de processos; Aprovar a metodologia de gestão de processos; Aprovar a priorização para a gestão de processos; Aprovar os indicadores e metas de desempenho dos processos; Avaliar a aferição dos indicadores dos processos."],
            [3, "Unidade de Gestão Normativa e de Processos (AE/GPE/UNP)", "3.3", "Promover e supervisionar a gestão de processos na Codevasf; Coordenar o Núcleo de Gestão Normativa e de Processos; Sistematizar, padronizar e difundir princípios, práticas e padrões de gestão de processos; Elaborar e manter atualizada a metodologia de gestão de processos; Fornecer orientação e treinamento sobre o gerenciamento de processos; Difundir a cultura de gestão de processos; Gerenciar a arquitetura e o repositório de processos; Consolidar informações por meio de relatórios gerenciais."],
            [4, "Núcleo de Gestão Normativa e de Processos (Nugep)", "3.4", "Propor as diretrizes sobre a gestão de processos; Auxiliar no acompanhamento dos resultados dos processos e na proposição de correções e melhorias; Avaliar e propor melhorias na metodologia de gestão de processos; Avaliar e propor melhorias para elaboração de indicadores e metas de desempenho; Avaliar e propor melhorias nos artefatos e documentos relacionados à gestão de processos; Fomentar e promover a gestão de processos em suas unidades organizacionais; Propor a priorização para o gerenciamento de processos."],
            [5, "Ponto focal do Nugep", "3.5", "Conduzir e coordenar os trabalhos de gerenciamento de processos no âmbito de sua unidade orgânica; Conduzir as oficinas de trabalho para levantamento, análise, coleta de informações e proposição de melhorias; Planejar a implementação, o monitoramento e a avaliação dos processos mapeados; Atuar como fornecedor de informações técnicas específicas, mesmo em processo fora de sua unidade de lotação; Atuar como agente multiplicador e facilitador da AE/GPE/UNP; Ter perfil de liderança e conhecimento em gestão de processos; Ter prioridade nas capacitações relacionadas ao tema."],
            [6, "Gestor do processo", "3.6", "Aprovar os processos de trabalho mapeados pelos respectivos pontos focais do Nugep; Formar equipe de gerenciamento de processos para gerenciamento do processo; Engajar os atores do processo nos trabalhos de gerenciamento do processo; Gerenciar e monitorar os processos sob sua responsabilidade; Reportar os resultados dos processos à AE/GPE/UNP; Acompanhar os trabalhos de gerenciamento dos processos sob sua responsabilidade; Elaborar, monitorar e prestar informações sobre os indicadores de desempenho dos processos; Assegurar que o processo atenda às expectativas de desempenho estabelecidas; Propor melhorias ou inovações, com vistas a tornar os processos/subprocessos eficientes, eficazes e efetivos; Disseminar os processos/subprocessos mapeados dentro da respectiva unidade orgânica."],
            [7, "Equipe de Gerenciamento de Processo", "3.7", "Realizar o gerenciamento dos processos priorizados de sua competência."],
        ]
        escreve(ct, competencias, wrap_cols={4}, center_cols={1, 3})
        ct.freeze_panes = "B2"
        mudou.append("Nova aba Competencias (RES 031/2025, item 3)")

    # 7) LEIA-ME: descreve as abas novas.
    if "LEIA-ME" in wb.sheetnames:
        leia = wb["LEIA-ME"]
        descricoes_existentes = {leia.cell(row=r, column=1).value for r in range(1, leia.max_row + 1)}
        extras = []
        if "PlanoAcao" not in descricoes_existentes:
            extras.append(("PlanoAcao", "Ações do plano 5W2H (O que/Por que/Onde/Quando/Quem/Como/Quanto) da etapa "
                                          "\"Analisar o processo\" — vincula a um Processo, Subprocesso ou Atividade, como Riscos."))
        if "Competencias" not in descricoes_existentes:
            extras.append(("Competencias", "Competências e atribuições de cada instância na gestão de processos "
                                            "(RES 031/2025, item 3) — exibidas na aba NUGEP do painel."))
        for nome, desc in extras:
            r = leia.max_row + 1
            c1 = leia.cell(row=r, column=1, value=nome); c1.font = Font(name=FONTE, size=10, bold=True)
            c2 = leia.cell(row=r, column=2, value=desc); c2.font = F_CELL; c2.alignment = AL_WRAP
        if extras:
            mudou.append("LEIA-ME: descrição das abas novas")

    if mudou:
        wb.save(CAMINHO)
        print("Planilha atualizada:")
        for m in mudou:
            print(" -", m)
        # wb.save() aqui embaixo relê e regrava o arquivo com openpyxl, que
        # nunca escreve o valor calculado de uma fórmula (só a fórmula em
        # si) — mesmo em fórmulas que este script não tocou. Sem recachear,
        # Processos.Macroprocesso/Percentual, Documentos/Riscos.Vinculo_Nivel,
        # Riscos.Nivel_PxI/Classificacao e NUGEP.Unidade_Nome voltam a sair
        # vazios para qualquer leitor que não seja o Excel (SheetJS incluso).
        print("\nRode em seguida, NESTA ORDEM (a primeira é obrigatória):")
        print("  python scripts/cachear_formulas.py")
        print("  python scripts/planilha_para_js.py")
    else:
        print("Nada para atualizar — a planilha já está em dia.")


if __name__ == "__main__":
    principal()
