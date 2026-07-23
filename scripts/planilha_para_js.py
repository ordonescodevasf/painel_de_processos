# -*- coding: utf-8 -*-
"""
Lê data/painel-processos-dados.xlsx e gera js/dados.js
(window.PAINEL_DADOS), usado como fallback quando o painel é aberto
sem servidor (file://) ou quando o Google Sheets/planilha local falha.

Uso:  python scripts/planilha_para_js.py
Requisito: openpyxl (pip install openpyxl)
"""
import json
import os
import datetime as dt
from openpyxl import load_workbook

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
XLSX = os.path.join(BASE, "data", "painel-processos-dados.xlsx")
SAIDA = os.path.join(BASE, "js", "dados.js")

ABAS = ["Macroprocessos", "Processos", "Subprocessos", "Atividades",
        "Documentos", "Riscos", "Indicadores", "Diario_Mapeamento"]


def valor(v):
    """Normaliza valores de célula para JSON."""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v


def aba_para_linhas(ws):
    """Converte uma aba em lista de objetos {cabecalho: valor} (linhas com Código/ID)."""
    linhas = list(ws.iter_rows(values_only=True))
    if not linhas:
        return []
    cab = [str(c).strip() if c is not None else "" for c in linhas[0]]
    out = []
    for row in linhas[1:]:
        if row[0] is None or str(row[0]).strip() == "":
            continue  # ignora linhas sem código/ID
        obj = {}
        for k, v in zip(cab, row):
            if k:
                obj[k] = valor(v)
        out.append(obj)
    return out


def main():
    # data_only=True lê os valores calculados (a planilha entregue já foi
    # recalculada; se você editar e as fórmulas zerarem, abra e salve no
    # Excel/LibreOffice antes de rodar este script)
    wb = load_workbook(XLSX, data_only=True)
    dados = {"_gerado_em": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
             "_fonte": os.path.basename(XLSX)}
    for aba in ABAS:
        if aba not in wb.sheetnames:
            print(f"AVISO: aba '{aba}' não encontrada — ignorada.")
            dados[aba] = []
            continue
        dados[aba] = aba_para_linhas(wb[aba])
        print(f"{aba}: {len(dados[aba])} linhas")

    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write("/* GERADO AUTOMATICAMENTE por scripts/planilha_para_js.py — não edite à mão.\n")
        f.write("   Fonte: data/painel-processos-dados.xlsx */\n")
        f.write("window.PAINEL_DADOS = ")
        json.dump(dados, f, ensure_ascii=False, indent=1)
        f.write(";\n")
    print(f"OK → {os.path.relpath(SAIDA, BASE)}")


if __name__ == "__main__":
    main()
