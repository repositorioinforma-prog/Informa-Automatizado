"""
Verificador de divergências do Relatório Descritivo — lógica pura (sem Streamlit).

Lê a planilha do Descritivo (seções "RESULTADOS PELO TOTAL" e "RESULTADOS PELOS
SEGMENTOS"), monta o inventário de perguntas e compara totais x segmentos:

  TITULO_DIVERGENTE  título parecido mas não igual entre total e segmento
  SEM_PAR_SEGMENTO   título só existe nos totais
  SEM_PAR_TOTAL      título só existe nos segmentos
  OPCOES_DIVERGENTES mesma pergunta, rótulos de resposta diferentes
  PERGUNTA_DIVERGENTE texto "Pergunta:" diferente entre total e segmento

Reaproveita `extrair_tabelas` (core/avisos_consistencia_math.py) para título,
rótulos e bases, para não duplicar o parser que o app já usa.
"""
from __future__ import annotations

import re
import unicodedata
from collections import OrderedDict
from difflib import SequenceMatcher

from core.avisos_consistencia_math import extrair_tabelas

LIMIAR_PARECIDO = 0.75


def chave(texto) -> str:
    t = "".join(c for c in unicodedata.normalize("NFKD", str(texto)) if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", t.lower()).strip()


def _marcadores(ws):
    achados = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None:
            continue
        k = chave(v)
        if k in ("resultados pelo total", "resultados pelos segmentos"):
            achados.append((r, k))
    return achados


def _texto_pergunta(ws, linha_base: int) -> str:
    """Texto da linha 'Pergunta: ...' logo abaixo da Base (até 4 linhas)."""
    for r in range(linha_base + 1, linha_base + 5):
        v = ws.cell(r, 1).value
        if v is not None and chave(v).startswith("pergunta"):
            return re.sub(r"\s+", " ", str(v)).strip()
    return ""


def montar_inventario(ws) -> dict:
    """{'totais': [...], 'segmentos': [...], 'tem_secoes': bool}"""
    tabelas = extrair_tabelas(ws)
    ms = _marcadores(ws)
    ult_total = max((r for r, k in ms if k == "resultados pelo total"), default=None)
    ult_seg = max((r for r, k in ms if k == "resultados pelos segmentos"), default=None)
    inv = {"totais": [], "segmentos": [], "tem_secoes": bool(ult_total and ult_seg)}
    for t in tabelas:
        t = dict(t)
        t["pergunta"] = _texto_pergunta(ws, t["linha_base"])
        if ult_seg and t["linha_titulo"] > ult_seg:
            inv["segmentos"].append(t)
        elif ult_total and t["linha_titulo"] > ult_total:
            inv["totais"].append(t)
    return inv


def _agrupar(tabelas):
    g = OrderedDict()
    for t in tabelas:
        g.setdefault(chave(t["titulo"]), []).append(t)
    return g


def _rotulos(t):
    return {chave(x) for x in t["rotulos"] if chave(x) and "continua" not in chave(x)}


def _uniao_rotulos(tabelas):
    u = set()
    for t in tabelas:
        u |= _rotulos(t)
    return u


def comparar(inv: dict) -> list[dict]:
    """Lista de achados: {tipo, severidade, titulo_total, titulo_segmento, detalhe, linhas}."""
    achados = []
    if not inv["tem_secoes"]:
        return achados
    gt, gs = _agrupar(inv["totais"]), _agrupar(inv["segmentos"])
    so_tot = [k for k in gt if k not in gs]
    so_seg = [k for k in gs if k not in gt]

    usados = set()
    for ks in so_seg:
        melhor, r = None, 0.0
        for kt in so_tot:
            s = SequenceMatcher(None, ks, kt).ratio()
            if s > r:
                melhor, r = kt, s
        if melhor and r >= LIMIAR_PARECIDO:
            usados.add(melhor)
            achados.append({
                "tipo": "TITULO_DIVERGENTE", "severidade": "ALTA" if r >= 0.9 else "MÉDIA",
                "titulo_total": gt[melhor][0]["titulo"], "titulo_segmento": gs[ks][0]["titulo"],
                "detalhe": f"parecidos ({r:.0%}), mas não iguais; {len(gs[ks])} tabela(s) de segmento afetadas",
                "linhas": f"total L{gt[melhor][0]['linha_titulo']} / segmento L{gs[ks][0]['linha_titulo']}",
            })
        else:
            achados.append({
                "tipo": "SEM_PAR_TOTAL", "severidade": "MÉDIA",
                "titulo_total": "", "titulo_segmento": gs[ks][0]["titulo"],
                "detalhe": f"{len(gs[ks])} tabela(s) de segmento sem título correspondente nos totais",
                "linhas": f"segmento L{gs[ks][0]['linha_titulo']}",
            })
    for kt in so_tot:
        if kt in usados:
            continue
        achados.append({
            "tipo": "SEM_PAR_SEGMENTO", "severidade": "MÉDIA",
            "titulo_total": gt[kt][0]["titulo"], "titulo_segmento": "",
            "detalhe": "pergunta existe nos totais e não aparece nos segmentos",
            "linhas": f"total L{gt[kt][0]['linha_titulo']}",
        })

    # Mesma pergunta (título igual): compara opções (união dos pedaços) e texto da pergunta
    for k, tabs_t in gt.items():
        if k not in gs:
            continue
        t0, s0 = tabs_t[0], gs[k][0]
        so_t = _uniao_rotulos(tabs_t) - _uniao_rotulos(gs[k])
        so_s = _uniao_rotulos(gs[k]) - _uniao_rotulos(tabs_t)
        if so_t or so_s:
            partes = []
            if so_t:
                partes.append("só no total: " + "; ".join(sorted(so_t))[:120])
            if so_s:
                partes.append("só nos segmentos: " + "; ".join(sorted(so_s))[:120])
            achados.append({
                "tipo": "OPCOES_DIVERGENTES", "severidade": "ALTA",
                "titulo_total": t0["titulo"], "titulo_segmento": s0["titulo"],
                "detalhe": " | ".join(partes),
                "linhas": f"total L{t0['linha_titulo']} / segmento L{s0['linha_titulo']}",
            })
        pergs_t = {chave(t["pergunta"]) for t in tabs_t if t["pergunta"]}
        pergs_s = {chave(t["pergunta"]) for t in gs[k] if t["pergunta"]}
        if pergs_t and pergs_s and not (pergs_t & pergs_s):
            achados.append({
                "tipo": "PERGUNTA_DIVERGENTE", "severidade": "MÉDIA",
                "titulo_total": t0["titulo"], "titulo_segmento": s0["titulo"],
                "detalhe": f"total: {tabs_t[0]['pergunta'][:70]} | segmento: {gs[k][0]['pergunta'][:70]}",
                "linhas": f"total L{t0['linha_titulo']} / segmento L{s0['linha_titulo']}",
            })
    return achados


# ---------------------------------------------------------------------------
# Observações sobre achados conhecidos (editável pela equipe)
# ---------------------------------------------------------------------------
# Cada item: tipo do achado + regex (aplicada sobre títulos + detalhe, já sem
# acento/caixa) + texto mostrado na coluna "Observação". Servem para separar o
# que provavelmente é padrão do relatório do que precisa de conferência.
OBSERVACOES_CONHECIDAS = [
    {
        "tipo": ("TITULO_DIVERGENTE", "SEM_PAR_SEGMENTO"),
        "padrao": r"senador.*(base 100|primeiro candidato|segundo candidato)",
        "texto": "Padrão visto em vários estudos: totais separam 1º e 2º candidato e os "
                 "segmentos usam base 100% (1º + 2º voto). Provavelmente estrutural — confirmar.",
    },
    {
        "tipo": ("OPCOES_DIVERGENTES",),
        "padrao": r"partido federacao",
        "texto": "Provável falso positivo: 'Partido/Federação' é cabeçalho de coluna na "
                 "tabela de segmento, não uma opção de resposta.",
    },
]

_OBS_PADRAO = {
    "OPCOES_DIVERGENTES": "Conferir: rótulo de opção escrito diferente entre total e segmentos.",
    "SEM_PAR_SEGMENTO": "Conferir: pergunta presente nos totais e ausente dos segmentos (intencional?).",
    "SEM_PAR_TOTAL": "Conferir: pergunta presente nos segmentos e ausente dos totais.",
    "TITULO_DIVERGENTE": "Conferir: título escrito diferente entre totais e segmentos.",
    "PERGUNTA_DIVERGENTE": "Conferir: texto da pergunta diferente entre totais e segmentos.",
}

ROTULO_TIPO = {
    "TITULO_DIVERGENTE": "Título divergente",
    "SEM_PAR_SEGMENTO": "Sem par nos segmentos",
    "SEM_PAR_TOTAL": "Sem par nos totais",
    "OPCOES_DIVERGENTES": "Opções divergentes",
    "PERGUNTA_DIVERGENTE": "Pergunta divergente",
}


def observacao(achado: dict) -> str:
    base = chave(f"{achado['titulo_total']} {achado['titulo_segmento']} {achado['detalhe']}")
    for obs in OBSERVACOES_CONHECIDAS:
        if achado["tipo"] in obs["tipo"] and re.search(obs["padrao"], base):
            return obs["texto"]
    return _OBS_PADRAO.get(achado["tipo"], "Conferir.")


def escolher_aba(wb) -> list[str]:
    """Abas candidatas a conter os resultados (ignora o 'Sumário')."""
    nomes = [ws.title for ws in wb.worksheets if chave(ws.title) != "sumario"]
    return nomes or [wb.worksheets[0].title]


def resumo_inventario(inv: dict, achados: list[dict]) -> dict:
    por_tipo = {t: 0 for t in ROTULO_TIPO}
    for a in achados:
        por_tipo[a["tipo"]] = por_tipo.get(a["tipo"], 0) + 1
    return {
        "tabelas_totais": len(inv["totais"]),
        "tabelas_segmentos": len(inv["segmentos"]),
        "tem_secoes": inv["tem_secoes"],
        "achados": len(achados),
        "por_tipo": por_tipo,
    }


def relatorio_para_xlsx(nome_arquivo: str, inv: dict, achados: list[dict]) -> bytes:
    """Excel de conferência (abas Resumo e Achados) como bytes."""
    import io

    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    r = resumo_inventario(inv, achados)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws.append(["Arquivo", nome_arquivo])
    ws.append(["Tabelas nos totais", r["tabelas_totais"]])
    ws.append(["Tabelas nos segmentos", r["tabelas_segmentos"]])
    ws.append(["Tem as duas seções?", "sim" if r["tem_secoes"] else "não (só totais — nada a comparar)"])
    ws.append(["Achados", r["achados"]])
    for tipo, rotulo in ROTULO_TIPO.items():
        ws.append([rotulo, r["por_tipo"].get(tipo, 0)])
    for row in ws.iter_rows(min_col=1, max_col=1):
        row[0].font = Font(bold=True)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 60

    w2 = wb.create_sheet("Achados")
    w2.append(["Tipo", "Severidade", "Título nos totais", "Título nos segmentos",
               "Detalhe", "Onde (linhas)", "Observação"])
    for a in achados:
        w2.append([ROTULO_TIPO.get(a["tipo"], a["tipo"]), a["severidade"], a["titulo_total"],
                   a["titulo_segmento"], a["detalhe"], a["linhas"], observacao(a)])
    fundo = PatternFill("solid", fgColor="1F3864")
    for c in w2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fundo
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, larg in enumerate([22, 11, 48, 48, 60, 24, 60], 1):
        w2.column_dimensions[get_column_letter(i)].width = larg
    for row in w2.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    w2.freeze_panes = "A2"
    if achados:
        w2.auto_filter.ref = w2.dimensions

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()
