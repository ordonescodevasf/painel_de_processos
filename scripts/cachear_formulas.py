#!/usr/bin/env python3
"""Grava o valor calculado ao lado de cada fórmula da planilha do painel.

Por que este passo existe
-------------------------
O openpyxl escreve fórmulas, mas nunca o resultado delas: o cache de valor
(o `<v>` que fica ao lado do `<f>` no XML) só é preenchido por um programa
que sabe calcular — o Excel, ao salvar. Um .xlsx recém-gerado por script
tem, portanto, células de fórmula literalmente vazias para qualquer leitor
que não seja o Excel.

O painel lê a planilha com SheetJS, que devolve o cache. Sem ele, sete
colunas chegavam nulas e o painel perdia o vínculo dos processos com o
macroprocesso, o progresso do mapeamento e os vínculos de documentos,
riscos e indicadores.

Este script roda DEPOIS de gerar_planilha.py / atualizar_planilha.py e
recalcula as mesmas fórmulas em Python, gravando `<f>` e `<v>` juntos —
exatamente o que o Excel faria. As fórmulas continuam na planilha e seguem
recalculando sozinhas quando alguém edita no Excel.

Uso:
    python scripts/cachear_formulas.py [caminho/da/planilha.xlsx]
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

PADRAO = Path(__file__).resolve().parent.parent / "data" / "painel-processos-dados.xlsx"

CELULA = re.compile(rb"<c\b[^>]*>.*?</c>|<c\b[^>]*/>", re.S)
LINHA = re.compile(rb'<row[^>]*r="(\d+)"[^>]*>.*?</row>', re.S)
REF = re.compile(rb'r="([A-Z]+)(\d+)"')
TEXTO = re.compile(rb"<t[^>]*>(.*?)</t>", re.S)
VALOR = re.compile(rb"<v>(.*?)</v>", re.S)
FORMULA = re.compile(rb"<f>.*?</f>", re.S)


def _texto(celula: bytes) -> str:
    achados = TEXTO.findall(celula)
    if achados:
        bruto = b"".join(achados).decode("utf-8")
    else:
        achado = VALOR.search(celula)
        bruto = achado.group(1).decode("utf-8") if achado else ""
    return (
        bruto.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
        .strip()
    )


def _ler_aba(xml: bytes) -> tuple[list[dict[str, str]], dict[int, dict[str, bytes]]]:
    """Devolve (linhas como dicionário coluna→texto, células cruas por linha)."""
    linhas, cruas = [], {}
    for bruto in LINHA.finditer(xml):
        numero = int(bruto.group(1))
        celulas = {}
        for c in CELULA.finditer(bruto.group(0)):
            ref = REF.search(c.group(0))
            if ref:
                celulas[ref.group(1).decode()] = c.group(0)
        cruas[numero] = celulas
        linhas.append({"_r": numero, **{k: _texto(v) for k, v in celulas.items()}})
    return linhas, cruas


def _colunas(ate: str = "AL") -> list[str]:
    nomes = [chr(c) for c in range(ord("A"), ord("Z") + 1)]
    nomes += ["A" + chr(c) for c in range(ord("A"), ord("Z") + 1)]
    return nomes[: nomes.index(ate) + 1]


COLS = _colunas()


def _nivel(codigo: str) -> str:
    codigo = codigo.split(";")[0].strip().upper()
    if codigo.startswith(("MG-", "MF-", "MS-")):
        return "Macroprocesso"
    if codigo.startswith("PP-"):
        return "Processo"
    if codigo.startswith("SP-"):
        return "Subprocesso"
    if codigo.startswith("AT-"):
        return "Atividade"
    if codigo.startswith("TR-"):
        return "Tarefa"
    return ""


def _num(texto: str) -> float | None:
    try:
        return float(texto)
    except (TypeError, ValueError):
        return None


def calcular(aba: str, linhas: list[dict[str, str]]) -> dict[int, dict[str, object]]:
    """Recalcula, por aba, as colunas que a planilha define como fórmula."""
    fora: dict[int, dict[str, object]] = {}

    def por(linha, coluna, valor):
        fora.setdefault(linha, {})[coluna] = valor

    if aba == "Processos":
        # M1..M10 = colunas R..AA (18ª a 27ª); Percentual (col. M) ignora
        # marcos "Não se aplica" — só conta Sim/Não no denominador.
        marcos = COLS[COLS.index("R") : COLS.index("AA") + 1]
        for linha in linhas[1:]:
            sim = sum(1 for c in marcos if linha.get(c, "") == "Sim")
            nao = sum(1 for c in marcos if linha.get(c, "") == "Não")
            por(linha["_r"], "M", round(sim / (sim + nao), 4) if sim + nao else 0)

    elif aba == "Documentos":
        for linha in linhas[1:]:
            por(linha["_r"], "D", _nivel(linha.get("E", "")))

    elif aba == "Riscos":
        for linha in linhas[1:]:
            por(linha["_r"], "B", _nivel(linha.get("C", "")))
            p, i = _num(linha.get("I", "")), _num(linha.get("J", ""))
            nivel = "" if p is None or i is None else int(p * i)
            por(linha["_r"], "K", nivel)
            if nivel == "":
                classe = ""
            elif nivel >= 20:
                classe = "Extremo"
            elif nivel >= 12:
                classe = "Alto"
            elif nivel >= 5:
                classe = "Moderado"
            else:
                classe = "Baixo"
            por(linha["_r"], "L", classe)

    elif aba == "Metricas":
        for linha in linhas[1:]:
            por(linha["_r"], "C", _nivel(linha.get("D", "")))

    elif aba in ("Papeis", "Regras"):
        col_codigo = "C" if aba == "Papeis" else "D"
        col_nivel = "B" if aba == "Papeis" else "C"
        for linha in linhas[1:]:
            por(linha["_r"], col_nivel, _nivel(linha.get(col_codigo, "")))

    return fora


def cachear(caminho: Path) -> int:
    with zipfile.ZipFile(caminho) as z:
        nomes = z.namelist()
        arquivos = {n: z.read(n) for n in nomes}

    livro = arquivos["xl/workbook.xml"].decode("utf-8")
    abas = re.findall(r'<sheet name="([^"]+)"', livro)
    arquivo_da_aba = {a: f"xl/worksheets/sheet{i + 1}.xml" for i, a in enumerate(abas)}

    # Unidade_Nome do NUGEP é um VLOOKUP direto na aba Siglas (A:B).
    siglas, _ = _ler_aba(arquivos[arquivo_da_aba["Siglas"]])
    unidades = {l.get("A", ""): l.get("B", "") for l in siglas[1:] if l.get("A")}

    gravadas = 0
    for aba in ("Processos", "Documentos", "Riscos", "Metricas", "Papeis", "Regras", "NUGEP"):
        if aba not in arquivo_da_aba:
            continue
        alvo = arquivo_da_aba[aba]
        linhas, _ = _ler_aba(arquivos[alvo])
        if aba == "NUGEP":
            calculado = {
                l["_r"]: {"E": unidades.get(l.get("D", ""), "")} for l in linhas[1:]
            }
        else:
            calculado = calcular(aba, linhas)

        def trocar_linha(bruto: re.Match) -> bytes:
            numero = int(bruto.group(1))
            desta = calculado.get(numero)
            if not desta:
                return bruto.group(0)

            def trocar_celula(c: re.Match) -> bytes:
                nonlocal gravadas
                celula = c.group(0)
                ref = REF.search(celula)
                if not ref:
                    return celula
                coluna = ref.group(1).decode()
                if coluna not in desta or b"<f>" not in celula or b"<v>" in celula:
                    return celula
                valor = desta[coluna]
                if valor == "" or valor is None:
                    return celula  # fórmula cujo resultado é vazio
                gravadas += 1
                atributos = celula[: celula.index(b">")]
                formula = FORMULA.search(celula).group(0)
                if isinstance(valor, str):
                    atributos = re.sub(rb'\s+t="[^"]*"', b"", atributos) + b' t="str"'
                return (
                    atributos
                    + b">"
                    + formula
                    + b"<v>"
                    + escape(str(valor)).encode("utf-8")
                    + b"</v></c>"
                )

            return CELULA.sub(trocar_celula, bruto.group(0))

        arquivos[alvo] = LINHA.sub(trocar_linha, arquivos[alvo])

    reserva = caminho.with_suffix(".xlsx.bak")
    shutil.copy2(caminho, reserva)
    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as z:
        for nome in nomes:
            z.writestr(nome, arquivos[nome])
    return gravadas


if __name__ == "__main__":
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else PADRAO
    if not destino.exists():
        sys.exit(f"planilha não encontrada: {destino}")
    total = cachear(destino)
    print(f"{total} fórmulas gravadas com valor calculado em {destino.name}")
    print(f"cópia do arquivo anterior em {destino.with_suffix('.xlsx.bak').name}")
