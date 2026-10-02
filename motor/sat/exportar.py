"""Planilha com as tabelas do diagnóstico por satélite."""
from __future__ import annotations

import io

import pandas as pd

from . import clima as cl
from .diagnostico import Diagnostico


def tabelas(d: Diagnostico) -> dict[str, pd.DataFrame]:
    t: dict[str, pd.DataFrame] = {}
    A = d.area.area_ha
    t["Resumo"] = pd.DataFrame([{"Talhão": d.area.nome, "Área (ha)": round(A, 2),
                                 "Longitude": d.area.centro[0], "Latitude": d.area.centro[1]}])
    H = d.historico
    if H is not None:
        g = H.grupos_ha.round(2)
        g.index.name = "Ano"
        t["Uso do solo (ha)"] = g.reset_index()
        t["Linha do tempo"] = pd.DataFrame(H.eventos, columns=["Evento", "Descrição"])
        if H.prodes is not None and len(H.prodes):
            t["PRODES"] = H.prodes
    L = d.lavoura
    if L is not None:
        s = L.serie.copy()
        s.columns = ["Data", "Ano agrícola", "NDVI médio", "NDVI p10", "NDVI p90", "CV NDVI (%)", "NDVI", "NDRE",
                     "NDMI", "MSAVI2"][:len(s.columns)]
        t["Série Sentinel-2"] = s
        if not L.indices_ano.empty:
            t["Índices por ano"] = L.indices_ano
        if L.estab_area:
            t["Estabilidade (ha)"] = pd.DataFrame([{"Classe": k, "Área (ha)": round(v, 2),
                                                    "% da área": round(100 * v / A, 1)} for k, v in L.estab_area.items()])
        if len(L.pontos):
            t["Pontos prioritários"] = L.pontos[[c for c in L.pontos.columns if not c.startswith("_")]]
    R = d.relevo
    if R is not None:
        t["Declividade (ha)"] = pd.DataFrame([{"Classe": k, "Área (ha)": round(v, 2)} for k, v in R.decliv_ha.items()])
        t["Posição (ha)"] = pd.DataFrame([{"Posição": k, "Área (ha)": round(v, 2)} for k, v in R.posicao_ha.items()])
        t["Erosão (ha)"] = pd.DataFrame([{"Suscetibilidade": k, "Área (ha)": round(v, 2)}
                                         for k, v in R.erosao_ha.items()])
        if len(R.saidas):
            t["Saídas de escoamento"] = R.saidas[[c for c in R.saidas.columns if not c.startswith("_")]]
    C = d.clima
    if C is not None:
        t["Chuva anual"] = C.anual.reset_index(drop=True)
        n = C.normal.copy()
        n.index = [cl.MESES[m - 1] for m in n.index]
        n.index.name = "Mês"
        t["Chuva mensal (normais)"] = n.round(1).reset_index()
        if not C.veranicos.empty:
            t["Veranicos"] = C.veranicos.round(1)
        if not C.estacao.empty:
            e = C.estacao.copy()
            for c in ("Início", "Fim"):
                e[c] = pd.to_datetime(e[c]).dt.strftime("%d/%m/%Y")
            t["Estação chuvosa"] = e
        if not C.extremos.empty:
            t["Temperaturas"] = C.extremos.round(2)
        if getattr(C, "vento", None) is not None:
            from . import vento as vt
            t.update(vt.tabelas(C.vento))
    S = d.solo
    if S is not None:
        t["Solo (SoilGrids)"] = S.tabela[["Atributo", "Unidade", "profundidade", "media", "q05", "q95"]].rename(
            columns={"profundidade": "Profundidade", "media": "Média", "q05": "P5", "q95": "P95"}).round(2)
    T = d.termico
    if T is not None:
        t["Temperatura sup. (Landsat)"] = T.tabela.round(2)
    if d.avisos:
        t["Avisos"] = pd.DataFrame({"Aviso": d.avisos})
    return t


def excel(d: Diagnostico) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, df in tabelas(d).items():
            df.to_excel(w, sheet_name=nome[:31], index=False)
            ws = w.sheets[nome[:31]]
            for col in ws.columns:
                larg = max(len(str(c.value or "")) for c in col[:60])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, larg + 2), 60)
    return buf.getvalue()
