# -*- coding: utf-8 -*-
"""
Atualiza data/painel-processos-dados.xlsx **no lugar**, sem recriar a planilha,
para a versão em que uma atividade pode pender direto de um processo — ou seja,
processo sem subprocessos, mas com atividades e tarefas.

Use este script (e não gerar_planilha.py) sempre que a planilha já tiver
conteúdo real ou colunas acrescentadas depois — ex.: "Objetivo" em Atividades e
Tarefas. Aqui nada é recriado: dados, colunas extras, formatação, fórmulas e
validações continuam como estão.

O que faz:
  1. Aba Atividades: renomeia o cabeçalho "Subprocesso" para "Vinculo_Pai".
     A coluna passa a aceitar o código de um Subprocesso (SP-...) OU de um
     Processo (P-...). O painel identifica pelo prefixo do código.
  2. Aba Listas: acrescenta "Não se aplica" à lista Sim_Nao (usada nas colunas
     de marcos M1–M10) e estende as validações que apontam para ela — serve
     para marcar, por exemplo, M3 "Subprocessos modelados" num processo que
     não tem subprocessos.
  3. Aba LEIA-ME: atualiza o texto das regras de vínculo.
  4. Acrescenta as colunas novas das fichas, se ainda não existirem, e
     preenche os valores de demonstração: "Unidades_Corresponsaveis" em
     Macroprocessos, Processos, Subprocessos, Atividades e Tarefas, e
     "Executor" (um cargo) em Atividades.
  5. Troca "Procedimento Operacional Padrão (POP)" por "Procedimento (PRO)"
     em todas as abas.
  6. Com --exemplo: insere um exemplo completo de processo sem subprocessos
     (P-06.03) — 2 atividades ligadas direto ao processo, 3 tarefas de uma
     delas, e M3 marcado como "Não se aplica".

Uso:
    python scripts/atualizar_planilha.py            # só o ajuste de estrutura
    python scripts/atualizar_planilha.py --exemplo   # estrutura + exemplo

Depois de rodar, regenere o fallback offline:
    python scripts/planilha_para_js.py

Requisito: openpyxl (pip install openpyxl)
"""
import os
import re
import sys
from copy import copy

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
XLSX = os.path.join(BASE, "data", "painel-processos-dados.xlsx")

NAO_SE_APLICA = "Não se aplica"

CORRESPONSAVEIS = {
  "Macroprocessos": {
      "MP-01": "AR/GDT; AI/GOM",
      "MP-02": "AI/GOM; AR/GRB",
      "MP-03": "AR/GRB; AA/GLC",
      "MP-04": "AA/GLC; AG/GGP",
      "MP-05": "AG/GGP; AT/GTI",
      "MP-06": "AT/GTI; AE/GPE",
      "MP-07": "AE/GPE; AE/GAG",
      "MP-08": "AE/GAG; AR/GDT",
  },
  "Processos": {
      "P-06.01": "AE/GAG; AR/GDT",
      "P-06.02": "AR/GDT; AI/GOM",
      "P-06.03": "AI/GOM; AR/GRB",
      "P-04.01": "AR/GRB; AA/GLC",
      "P-05.01": "AA/GLC; AG/GGP",
      "P-01.01": "AG/GGP; AT/GTI",
      "P-07.01": "AT/GTI; AE/GPE",
  },
  "Subprocessos": {
      "SP-06.01.01": "AI/GOM",
      "SP-06.01.02": "AR/GRB",
      "SP-06.01.03": "AG/GGP",
      "SP-06.01.03.01": "AG/GGP",
      "SP-06.02.01": "AT/GTI",
      "SP-04.01.01": "AE/GPE",
  },
  "Atividades": {
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
  },
  "Tarefas": {
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
  },
}

EXECUTORES = {
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

# Colunas novas por aba, na ordem em que entram à direita do cabeçalho atual.
COLUNAS_NOVAS = {
    "Macroprocessos": ["Unidades_Corresponsaveis"],
    "Processos": ["Unidades_Corresponsaveis"],
    "Subprocessos": ["Unidades_Corresponsaveis"],
    "Atividades": ["Executor", "Unidades_Corresponsaveis"],
    "Tarefas": ["Unidades_Corresponsaveis"],
}

# "Procedimento Operacional Padrão (POP)" passou a "Procedimento (PRO)" —
# a troca roda em todas as abas, em qualquer célula de texto.
TROCAS_PRO = [
    ("Procedimento Operacional Padrão (POP)", "Procedimento (PRO)"),
    ("Procedimento Operacional (POP)", "Procedimento (PRO)"),
    ("Procedimento Operacional Padrão", "Procedimento (PRO)"),
    ("POP 06.01", "PRO 06.01"),
    ("POP 06.03", "PRO 06.03"),
    ("revisar POP em", "revisar o PRO em"),
    ("Publicar POP,", "Publicar o PRO,"),
    ("dos POPs", "dos PROs"),
    ("POPs", "PROs"),
]


ATIVIDADES_EXEMPLO = [
    {
        "Codigo": "A-06.03.01",
        "Vinculo_Pai": "P-06.03",
        "Ordem": 1,
        "Nome": "Designar fiscais e gestor do contrato",
        "Descricao": "Formalizar a designação do fiscal técnico, do fiscal administrativo e do gestor do contrato, com ciência dos designados.",
        "Objetivo": "Garantir que a fiscalização comece com responsabilidades formalmente atribuídas.",
        "Responsavel_Ator": "AA/GLC",
        "Entradas": "Contrato assinado; Minuta de portaria de designação",
        "Saidas": "Portaria de designação publicada",
        "Sistemas": "SEI",
        "Prazo_Padrao": "5 dias úteis",
        "Base_Normativa": "Lei nº 14.133/2021, art. 117",
    },
    {
        "Codigo": "A-06.03.02",
        "Vinculo_Pai": "P-06.03",
        "Ordem": 2,
        "Nome": "Registrar medição e atestar a execução",
        "Descricao": "Conferir o objeto entregue no período, registrar a medição e atestar a nota fiscal para pagamento.",
        "Objetivo": "Liberar o pagamento apenas do que foi efetivamente entregue e conferido.",
        "Responsavel_Ator": "Fiscal técnico do contrato",
        "Entradas": "Relatório de execução; Nota fiscal",
        "Saidas": "Medição registrada; Nota fiscal atestada",
        "Sistemas": "SEI",
        "Prazo_Padrao": "10 dias úteis",
        "Base_Normativa": "Lei nº 14.133/2021, art. 140",
    },
]

TAREFAS_EXEMPLO = [
    {
        "Codigo": "T-06.03.02.01",
        "Atividade": "A-06.03.02",
        "Ordem": 1,
        "Nome": "Conferir o objeto entregue",
        "Descricao": "Verificar quantidade, qualidade e conformidade da entrega com o Termo de Referência.",
        "Objetivo": "Confirmar que a entrega corresponde ao que foi contratado.",
        "Tipo_Tarefa": "Manual",
        "Responsavel": "Fiscal técnico do contrato",
        "Sistema": "SEI",
        "Duracao_Estimada": "1 dia",
    },
    {
        "Codigo": "T-06.03.02.02",
        "Atividade": "A-06.03.02",
        "Ordem": 2,
        "Nome": "Registrar a medição no processo",
        "Descricao": "Lançar a medição do período no processo do contrato, anexando relatórios e evidências.",
        "Objetivo": "Deixar a medição registrada e rastreável no processo do contrato.",
        "Tipo_Tarefa": "Manual",
        "Responsavel": "Fiscal técnico do contrato",
        "Sistema": "SEI",
        "Duracao_Estimada": "0,5 dia",
    },
    {
        "Codigo": "T-06.03.02.03",
        "Atividade": "A-06.03.02",
        "Ordem": 3,
        "Nome": "Atestar a nota fiscal",
        "Descricao": "Atestar a nota fiscal depois da conferência e encaminhá-la para pagamento.",
        "Objetivo": "Autorizar o pagamento com base na conferência registrada.",
        "Tipo_Tarefa": "Manual",
        "Responsavel": "Gestor do contrato",
        "Sistema": "SEI",
        "Duracao_Estimada": "0,5 dia",
        "Observacoes": "O prazo de pagamento conta a partir do ateste.",
    },
]

TROCAS_LEIAME = [
    ("Atividades→Subprocesso",
     "Atividades→Vinculo_Pai (Subprocesso SP-... OU Processo P-..., quando o processo não tem subprocessos)"),
    ("4º nível (CBOK), com entradas, saídas, ator, sistemas e prazos.",
     "4º nível (CBOK), com entradas, saídas, ator, sistemas e prazos. Vinculo_Pai aceita o código de um "
     "Subprocesso (SP-...) OU de um Processo (P-...): processo que não tem subprocesso se decompõe direto "
     'em atividades, e estas em tarefas. Nesse caso, marque M3 (Subprocessos modelados) como "Não se aplica".'),
]


def cols_por_cabecalho(ws):
    """{cabeçalho: índice da coluna} lendo a linha 1."""
    return {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}


def ultima_linha(ws, coluna=1):
    ult = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=coluna).value not in (None, ""):
            ult = r
    return ult


def copia_estilo(destino, modelo):
    destino.font = copy(modelo.font)
    destino.border = copy(modelo.border)
    destino.alignment = copy(modelo.alignment)
    destino.fill = copy(modelo.fill)
    destino.number_format = modelo.number_format


def acrescenta_linha(ws, valores):
    """Escreve uma linha ao final da aba, casando pelos nomes de cabeçalho e
    herdando o estilo da última linha existente (nada de formatação nova)."""
    cols = cols_por_cabecalho(ws)
    r = ultima_linha(ws) + 1
    modelo = r - 1
    for nome, valor in valores.items():
        if nome not in cols:
            print(f'    AVISO: coluna "{nome}" não existe em {ws.title} — valor ignorado.')
            continue
        j = cols[nome]
        c = ws.cell(row=r, column=j, value=valor)
        if modelo >= 2:
            copia_estilo(c, ws.cell(row=modelo, column=j))
    return r


def acrescenta_colunas(ws, nomes):
    """Cria à direita as colunas que ainda não existem, herdando o estilo do
    cabeçalho vizinho. Devolve a lista das que foram criadas."""
    cols = cols_por_cabecalho(ws)
    criadas = []
    j = max(cols.values()) if cols else 0
    for nome in nomes:
        if nome in cols:
            continue
        j += 1
        c = ws.cell(row=1, column=j, value=nome)
        if j > 1:
            copia_estilo(c, ws.cell(row=1, column=j - 1))
        largura = 30
        ws.column_dimensions[get_column_letter(j)].width = largura
        criadas.append(nome)
    return criadas


def preenche_coluna(ws, nome, valores):
    """Preenche a coluna `nome` linha a linha, casando pelo código da 1ª
    coluna. Não sobrescreve célula que já tenha conteúdo."""
    cols = cols_por_cabecalho(ws)
    if nome not in cols:
        return 0
    j = cols[nome]
    escritas = 0
    for r in range(2, ws.max_row + 1):
        cod = str(ws.cell(row=r, column=1).value or "").strip()
        if not cod or cod not in valores:
            continue
        alvo = ws.cell(row=r, column=j)
        if alvo.value not in (None, ""):
            continue
        alvo.value = valores[cod]
        copia_estilo(alvo, ws.cell(row=r, column=max(1, j - 1)))
        escritas += 1
    return escritas


def existe_codigo(ws, codigo):
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or "").strip() == codigo:
            return True
    return False


def renomeia_cabecalho(ws, de, para):
    for c in ws[1]:
        if str(c.value or "").strip() == de:
            c.value = para
            return True
    return False


def acrescenta_item_lista(ws_listas, nome_lista, item):
    """Acrescenta um item à coluna nomeada da aba Listas. Devolve
    (letra_da_coluna, nova_ultima_linha) ou None se não achou a lista."""
    cols = cols_por_cabecalho(ws_listas)
    if nome_lista not in cols:
        return None
    j = cols[nome_lista]
    r = 1
    for rr in range(2, ws_listas.max_row + 1):
        if ws_listas.cell(row=rr, column=j).value not in (None, ""):
            r = rr
            if str(ws_listas.cell(row=rr, column=j).value).strip() == item:
                return (get_column_letter(j), r)  # já existe
    destino = ws_listas.cell(row=r + 1, column=j, value=item)
    if r >= 2:
        copia_estilo(destino, ws_listas.cell(row=r, column=j))
    return (get_column_letter(j), r + 1)


def estende_validacoes(wb, col_letra, ultima):
    """Aponta para o novo intervalo toda validação de lista que já usava a
    coluna col_letra da aba Listas."""
    nova = f"=Listas!${col_letra}$2:${col_letra}${ultima}"
    total = 0
    for ws in wb.worksheets:
        for v in ws.data_validations.dataValidation:
            f = str(v.formula1 or "")
            if "Listas" not in f:
                continue
            letras = re.findall(r"\$([A-Z]+)\$", f)
            if letras and letras[0] == col_letra and f != nova:
                v.formula1 = nova
                total += 1
    return total


def main():
    com_exemplo = "--exemplo" in sys.argv
    if not os.path.exists(XLSX):
        print(f"ERRO: não encontrei {XLSX}")
        return 1
    wb = load_workbook(XLSX)
    mudou = []

    # 1. Atividades: Subprocesso → Vinculo_Pai
    if "Atividades" in wb.sheetnames:
        at = wb["Atividades"]
        if "Vinculo_Pai" in cols_por_cabecalho(at):
            print("1. Atividades: cabeçalho já é Vinculo_Pai — nada a fazer.")
        elif renomeia_cabecalho(at, "Subprocesso", "Vinculo_Pai"):
            mudou.append('Atividades: "Subprocesso" → "Vinculo_Pai"')
            print('1. Atividades: cabeçalho "Subprocesso" renomeado para "Vinculo_Pai".')
        else:
            print("1. Atividades: não achei a coluna Subprocesso nem Vinculo_Pai.")
    else:
        print("1. AVISO: aba Atividades não encontrada.")

    # 2. Listas: "Não se aplica" em Sim_Nao + validações
    if "Listas" in wb.sheetnames:
        res = acrescenta_item_lista(wb["Listas"], "Sim_Nao", NAO_SE_APLICA)
        if res:
            col, ult = res
            qtd = estende_validacoes(wb, col, ult)
            mudou.append(f'Listas: Sim_Nao com "{NAO_SE_APLICA}"')
            print(f'2. Listas: "{NAO_SE_APLICA}" disponível em Sim_Nao '
                  f"(coluna {col}, até a linha {ult}); {qtd} validação(ões) estendida(s).")
        else:
            print("2. AVISO: lista Sim_Nao não encontrada na aba Listas.")
    else:
        print("2. AVISO: aba Listas não encontrada.")

    # 3. LEIA-ME
    if "LEIA-ME" in wb.sheetnames:
        lm = wb["LEIA-ME"]
        trocas = 0
        for linha in lm.iter_rows():
            for c in linha:
                if not isinstance(c.value, str):
                    continue
                novo = c.value
                for de, para in TROCAS_LEIAME:
                    if de in novo and para not in novo:
                        novo = novo.replace(de, para)
                if novo != c.value:
                    c.value = novo
                    trocas += 1
        if trocas:
            mudou.append("LEIA-ME: regras de vínculo atualizadas")
        print(f"3. LEIA-ME: {trocas} célula(s) atualizada(s).")
    else:
        print("3. AVISO: aba LEIA-ME não encontrada.")

    # 4. Colunas novas das fichas + valores de demonstração
    criadas, preenchidas = [], 0
    for aba, nomes in COLUNAS_NOVAS.items():
        if aba not in wb.sheetnames:
            print(f"4. AVISO: aba {aba} não encontrada.")
            continue
        ws = wb[aba]
        novas = acrescenta_colunas(ws, nomes)
        criadas += [f"{aba}.{n}" for n in novas]
        for nome in nomes:
            valores = EXECUTORES if nome == "Executor" else CORRESPONSAVEIS.get(aba, {})
            preenchidas += preenche_coluna(ws, nome, valores)
    if criadas or preenchidas:
        mudou.append(f"{len(criadas)} coluna(s) nova(s), {preenchidas} célula(s) preenchida(s)")
    print(f"4. Colunas das fichas: {len(criadas)} criada(s), {preenchidas} célula(s) preenchida(s).")

    # 5. POP → Procedimento (PRO), em todas as abas
    trocas_pro = 0
    for ws in wb.worksheets:
        for linha in ws.iter_rows():
            for c in linha:
                if not isinstance(c.value, str):
                    continue
                novo = c.value
                for de, para in TROCAS_PRO:
                    novo = novo.replace(de, para)
                if novo != c.value:
                    c.value = novo
                    trocas_pro += 1
    if trocas_pro:
        mudou.append(f"POP → Procedimento (PRO) em {trocas_pro} célula(s)")
    print(f"5. Procedimento (PRO): {trocas_pro} célula(s) atualizada(s).")

    # 6. Exemplo de processo sem subprocessos
    if com_exemplo:
        at = wb["Atividades"] if "Atividades" in wb.sheetnames else None
        tf = wb["Tarefas"] if "Tarefas" in wb.sheetnames else None
        novas = 0
        if at is not None:
            for reg in ATIVIDADES_EXEMPLO:
                if existe_codigo(at, reg["Codigo"]):
                    print(f'6. {reg["Codigo"]} já existe — não inserido.')
                    continue
                acrescenta_linha(at, reg)
                novas += 1
        if tf is not None:
            for reg in TAREFAS_EXEMPLO:
                if existe_codigo(tf, reg["Codigo"]):
                    print(f'6. {reg["Codigo"]} já existe — não inserido.')
                    continue
                acrescenta_linha(tf, reg)
                novas += 1
        # M3 do P-06.03 passa a "Não se aplica"
        if "Processos" in wb.sheetnames:
            pr = wb["Processos"]
            cols = cols_por_cabecalho(pr)
            j = cols.get("M3_Subprocessos_Modelados")
            if j:
                for r in range(2, pr.max_row + 1):
                    if str(pr.cell(row=r, column=1).value or "").strip() == "P-06.03":
                        pr.cell(row=r, column=j, value=NAO_SE_APLICA)
                        print(f'6. P-06.03: M3 marcado como "{NAO_SE_APLICA}".')
                        break
        if novas:
            mudou.append(f"{novas} linha(s) de exemplo (P-06.03 sem subprocessos)")
        print(f"6. Exemplo: {novas} linha(s) inserida(s).")
    else:
        print("6. Exemplo não inserido (rode com --exemplo se quiser o caso demonstrativo).")

    if not mudou:
        print("\nNada mudou — a planilha já estava atualizada.")
        return 0

    wb.save(XLSX)
    print("\nOK → " + os.path.relpath(XLSX, BASE))
    for m in mudou:
        print("   · " + m)
    print("\nAgora regenere o fallback offline:  python scripts/planilha_para_js.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
