"""
Total e Completo → Excel (lógica pura, sem Streamlit).

Escreve as tabelas calculadas por `core.total_completo_math` no layout do Excel que o SPSS exporta para o
Total e o Completo (o mesmo que o Relatório Descritivo tem antes da formatação final):

    Titulo: <título da tabela>                      (negrito, mesclado em toda a largura da tabela)
             | Total | <segmento> ...               (1 ou 2 linhas de cabeçalho; «Total» mesclado em 2 linhas)
    <categoria> | % | % ...                         (percentual da coluna, valor inteiro, formato ###0.0)
    Base        | n | n ...                         (casos sem peso, formato ###0)
    Pergunta: <texto da pergunta>
    (linha em branco)

Fontes e bordas seguem o que o SPSS grava: rótulos Arial 9, números Calibri 10, moldura média em volta do
bloco e linhas finas entre as colunas. As tabelas saem na ordem da programação, numa única aba.

A mesma função de escrita serve à pré-visualização da tela (`planilha_total_completo(..., como_texto=True)`):
o que se vê é o que sai no arquivo. Tabela que não pôde ser calculada NÃO entra no Excel; ela é listada numa
segunda aba («Não geradas») com o motivo, para não sumir em silêncio.
"""

from __future__ import annotations

import io
import re
from typing import Iterable, Sequence

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from core.analise_perguntas_math import _RE_COPIA, _RE_INVALIDOS, _RE_PREFIXO_TOTAL, _altura, _linhas, canon
from core.total_completo_math import ResultadoTotalCompleto, TabelaCalc, TabelaSps

NOME_ABA = "Total e Completo"
NOME_ABA_NAO_GERADAS = "Não geradas"
SECOES = ("Total", "Completo")

FORMATO_PCT = "###0.0"
FORMATO_BASE = "###0"

LARG_ROTULO = 22.7
LARG_TOTAL = 9.5
LARG_SEGMENTO = 13.6
_CHARS_POR_LARGURA = 1.15            # Arial 9 cabe ~15% mais caracteres que a unidade de largura (Calibri 11)

_F_TXT = Font(name="Arial", size=9)
_F_TIT = Font(name="Arial", size=9, bold=True)
_F_NUM = Font(name="Calibri", size=10)


def _lado(estilo: str | None) -> Side:
    return Side(style=estilo) if estilo else Side()


def _borda(esq=None, dir_=None, topo=None, base=None) -> Border:
    return Border(left=_lado(esq), right=_lado(dir_), top=_lado(topo), bottom=_lado(base))


def _larguras(n_colunas: int) -> list[float]:
    """Larguras das colunas A, B (Total), C… da tabela."""
    return [LARG_ROTULO, LARG_TOTAL] + [LARG_SEGMENTO] * max(n_colunas - 1, 0)


def _segmentos_do_cabecalho(calc: TabelaCalc) -> tuple[bool, list[tuple[int, int, str]]]:
    """Decide o cabeçalho. Devolve (tem_2_linhas, grupos) onde cada grupo é (col_ini, col_fim, rótulo) em
    índices 0-based das colunas de valor; só inclui grupos reais (vários sub-rótulos sob um nome)."""
    cols = calc.colunas
    grupos: list[tuple[int, int, str]] = []
    j = 0
    while j < len(cols):
        c = cols[j]
        vertical = c.grupo == c.rotulo
        if vertical:
            j += 1
            continue
        k = j
        while k + 1 < len(cols) and cols[k + 1].grupo == c.grupo and cols[k + 1].grupo != cols[k + 1].rotulo:
            k += 1
        grupos.append((j, k, c.grupo))
        j = k + 1
    return bool(grupos), grupos


def _escrever_tabela(ws, topo: int, calc: TabelaCalc) -> int:
    """Escreve uma tabela a partir da linha `topo`. Devolve a primeira linha livre depois dela
    (já com a linha em branco de separação)."""
    t: TabelaSps = calc.tabela
    cols = calc.colunas
    n = len(cols)
    ult = 1 + n                                   # última coluna (índice 1-based)
    larg = _larguras(n)
    larg_total = sum(larg)

    duas_linhas, grupos = _segmentos_do_cabecalho(calc)
    r_tit = topo
    r_h1 = topo + 1
    r_h2 = r_h1 + 1 if duas_linhas else r_h1
    r_d = r_h2 + 1
    n_lin = len(calc.rotulos)
    r_base = r_d + n_lin
    r_perg = r_base + 1
    chars_total = larg_total * _CHARS_POR_LARGURA

    # ---- título
    ws.merge_cells(start_row=r_tit, start_column=1, end_row=r_tit, end_column=ult)
    c = ws.cell(r_tit, 1, "Titulo: " + t.titulo)
    c.font = _F_TIT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    l_tit = _linhas(c.value, chars_total * 0.95)
    ws.row_dimensions[r_tit].height = 18.0 if l_tit <= 1 else (31.05 if l_tit == 2 else 45.0 if l_tit == 3 else 14.0 * l_tit)

    # ---- cabeçalho: bordas e textos
    a_cab = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
    em_grupo = {j: (ini, fim) for ini, fim, _ in grupos for j in range(ini, fim + 1)}
    for j, col in enumerate(cols):
        cc = 2 + j
        esq = "medium" if j == 0 else "thin"
        dir_ = "medium" if j == n - 1 else "thin"
        if j in em_grupo:
            # linha 1 = nome do grupo (mesclado), linha 2 = rótulo da categoria
            b1 = _borda(esq, dir_, "medium", "thin")
            b2 = _borda(esq, dir_, "thin", "medium")
            ws.cell(r_h1, cc).border = b1
            ws.cell(r_h2, cc).border = b2
            x = ws.cell(r_h2, cc, col.rotulo)
            x.font, x.alignment = _F_TXT, a_cab
        else:
            # sem grupo: o rótulo ocupa as duas linhas (ou a única) do cabeçalho
            ws.cell(r_h1, cc).border = _borda(esq, dir_, "medium", "medium" if r_h2 == r_h1 else None)
            if r_h2 != r_h1:
                ws.cell(r_h2, cc).border = _borda(esq, dir_, None, "medium")
                ws.merge_cells(start_row=r_h1, start_column=cc, end_row=r_h2, end_column=cc)
            x = ws.cell(r_h1, cc, col.rotulo)
            x.font, x.alignment = _F_TXT, a_cab
    for ini, fim, nome in grupos:
        ws.merge_cells(start_row=r_h1, start_column=2 + ini, end_row=r_h1, end_column=2 + fim)
        x = ws.cell(r_h1, 2 + ini, nome)
        x.font, x.alignment = _F_TXT, a_cab

    # alturas do cabeçalho
    def chars(j: int) -> float:
        return larg[1 + j] * _CHARS_POR_LARGURA if j > 0 else larg[1] * _CHARS_POR_LARGURA

    if duas_linhas:
        h1 = max((_linhas(nome, sum(chars(j) for j in range(ini, fim + 1))) for ini, fim, nome in grupos), default=1)
        h1v = max((_linhas(cols[j].rotulo, chars(j)) for j in range(n) if j not in em_grupo), default=1)
        ws.row_dimensions[r_h1].height = _altura(max(h1, 1))
        h2 = max((_linhas(cols[j].rotulo, chars(j)) for j in em_grupo), default=1)
        ws.row_dimensions[r_h2].height = _altura(max(h2, 1))
        # um rótulo sem grupo precisa caber nas duas linhas somadas
        falta = _altura(h1v) - (ws.row_dimensions[r_h1].height + ws.row_dimensions[r_h2].height)
        if falta > 0:
            ws.row_dimensions[r_h1].height += falta
    else:
        ws.row_dimensions[r_h1].height = _altura(max((_linhas(c.rotulo, chars(j)) for j, c in enumerate(cols)), default=1))

    # ---- linhas de dados e Base
    a_rot = Alignment(horizontal="left", vertical="top", wrap_text=True)
    a_num = Alignment(horizontal="right", vertical="top")
    for i, rot in enumerate(calc.rotulos):
        r = r_d + i
        primeira = i == 0
        x = ws.cell(r, 1, rot)
        x.font, x.alignment = _F_TXT, a_rot
        x.border = _borda("medium", "medium", "medium" if primeira else None, None)
        ws.row_dimensions[r].height = _altura(_linhas(rot, larg[0] * _CHARS_POR_LARGURA))
        for j in range(n):
            v = calc.valores[i, j]
            y = ws.cell(r, 2 + j)
            if v is not None and np.isfinite(v):
                y.value = float(v)
            y.font, y.alignment, y.number_format = _F_NUM, a_num, FORMATO_PCT
            y.border = _borda("medium" if j == 0 else "thin", "medium" if j == n - 1 else "thin", "medium" if primeira else None, None)

    x = ws.cell(r_base, 1, "Base")
    x.font, x.alignment = _F_TXT, a_rot
    x.border = _borda("medium", "medium", None, "medium")
    ws.row_dimensions[r_base].height = 15.0
    for j, col in enumerate(cols):
        y = ws.cell(r_base, 2 + j, float(col.base))
        y.font, y.alignment, y.number_format = _F_NUM, a_num, FORMATO_BASE
        y.border = _borda("medium" if j == 0 else "thin", "medium" if j == n - 1 else "thin", None, "medium")
    if n_lin == 0:                      # sem categorias: a Base fecha a moldura em cima também
        for k in range(1, ult + 1):
            b = ws.cell(r_base, k).border
            ws.cell(r_base, k).border = Border(left=b.left, right=b.right, top=_lado("medium"), bottom=b.bottom)

    # ---- pergunta
    proxima = r_base + 1
    if t.pergunta:
        ws.merge_cells(start_row=r_perg, start_column=1, end_row=r_perg, end_column=ult)
        x = ws.cell(r_perg, 1, "Pergunta: " + t.pergunta)
        x.font = _F_TXT
        x.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[r_perg].height = max(13.0, _altura(_linhas(x.value, chars_total)) if _linhas(x.value, chars_total) > 1 else 13.0)
        proxima = r_perg + 1
    return proxima + 1


def _largura_maxima(calcs: Iterable[TabelaCalc]) -> int:
    return max((len(c.colunas) for c in calcs), default=1)


def _ajusta_larguras(ws, n_colunas: int) -> None:
    for k, w in enumerate(_larguras(n_colunas), start=1):
        ws.column_dimensions[get_column_letter(k)].width = w


def _texto_formatado(valor: float, formato: str) -> str:
    """Como o Excel mostra o número (vírgula decimal)."""
    if formato == FORMATO_PCT:
        return f"{valor:.1f}".replace(".", ",")
    return f"{valor:.0f}"


def escolher(res: ResultadoTotalCompleto, secoes: Sequence[str] | None = None) -> list[TabelaCalc]:
    """Tabelas calculadas, na ordem da programação, das seções pedidas (todas por padrão)."""
    pedidas = {s.lower() for s in (secoes or SECOES)}
    return [c for c in res.calculadas if c is not None and c.tabela.secao.lower() in pedidas]


def planilha_tabelas(tabelas: Iterable[TabelaCalc], como_texto: bool = False):
    """Planilha (openpyxl) com as tabelas empilhadas, na mesma aba que vai para o Excel.
    `como_texto=True` troca os números pelo texto já formatado (o renderizador HTML da tela não aplica o
    formato de número das células)."""
    tabelas = list(tabelas)
    wb = Workbook()
    ws = wb.active
    ws.title = NOME_ABA
    topo = 1
    for calc in tabelas:
        topo = _escrever_tabela(ws, topo, calc)
    _ajusta_larguras(ws, _largura_maxima(tabelas))
    if como_texto:
        for linha in ws.iter_rows():
            for c in linha:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    c.value = _texto_formatado(float(c.value), c.number_format)
    return ws


def nao_geradas(res: ResultadoTotalCompleto, secoes: Sequence[str] | None = None) -> list[tuple[TabelaSps, str]]:
    pedidas = {s.lower() for s in (secoes or SECOES)}
    out = []
    for t, c in zip(res.tabelas, res.calculadas):
        if c is None and t.secao.lower() in pedidas:
            out.append((t, res.erros.get(t.ordem) or t.motivo or "não calculada"))
    return out


def montar_excel_total_completo(res: ResultadoTotalCompleto, secoes: Sequence[str] | None = None) -> bytes:
    """Excel final: aba «Total e Completo» (todas as tabelas, na ordem da programação) e, se alguma tabela não
    pôde ser calculada, a aba «Não geradas» com o motivo."""
    tabelas = escolher(res, secoes)
    ws = planilha_tabelas(tabelas)
    wb = ws.parent
    faltam = nao_geradas(res, secoes)
    if faltam:
        w2 = wb.create_sheet(NOME_ABA_NAO_GERADAS)
        w2.append(["#", "Seção", "Variável", "Título", "Motivo"])
        for c in w2[1]:
            c.font = Font(bold=True)
        for t, motivo in faltam:
            w2.append([t.ordem, t.secao, ("$" if t.multipla else "") + t.var_linha, t.titulo, motivo])
        for col, w in zip("ABCDE", (6, 11, 18, 50, 70)):
            w2.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def nome_arquivo_total_completo(nome_sps: str, secoes: Sequence[str] | None = None, extensao: str = ".xlsx") -> str:
    """«Total_e_Completo_Angra_dos_Reis_RJ_Set26.sps» → «Total e Completo_Angra_dos_Reis_RJ_Set26.xlsx».
    Só a seção Total → prefixo «Total»; só o Completo → «Completo»."""
    base = str(nome_sps or "").replace("\\", "/").rsplit("/", 1)[-1]
    base = re.sub(r"\.sps$", "", base.strip(), flags=re.I)
    base = _RE_COPIA.sub("", base)
    base = _RE_PREFIXO_TOTAL.sub("", base)
    if canon(base) in ("total", "total e completo", "completo"):
        base = ""
    base = _RE_INVALIDOS.sub("", base).strip(" _-.")
    pedidas = {s.lower() for s in (secoes or SECOES)}
    prefixo = "Total e Completo" if len(pedidas) != 1 else ("Total" if "total" in pedidas else "Completo")
    return (f"{prefixo}_{base}" if base else prefixo) + extensao
