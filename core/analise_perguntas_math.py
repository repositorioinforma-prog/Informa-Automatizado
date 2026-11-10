"""
Análise das Perguntas — lógica pura (sem Streamlit).

Fluxo:
  1. `ler_perguntas_total`  lê a programação Total (ou Total + Completo; só o TOTAL
                            é usado) e lista as variáveis que viram tabela, com o
                            título/pergunta de cada uma;
  2. `resolver_labels`      confere se cada variável tem label no banco; se não tiver,
                            usa o título da tabela do Total (a pessoa pode editar);
  3. `cruzar` / `calcular_analise`  cruzam cada pergunta pelas variáveis de cruzamento
                            (CRUZAMENTOS_PADRAO + as que a pessoa marcar) com a mesma
                            conta do `CROSSTABS /CELLS=COUNT COLUMN ASRESID /COUNT ROUND CELL`
                            do SPSS: casos com peso ausente/≤ 0 saem, células ponderadas
                            são arredondadas para inteiro, o % da coluna e o resíduo
                            ajustado saem das células já arredondadas;
  4. `montar_excel`         uma aba com todas as tabelas empilhadas (layout da exportação do SPSS já
                            formatada, com o destaque do VBA «DestacarAdjustedResiduals») + aba «Resumo».

Regras do projeto: nunca inventar valor e nunca usar zero como preenchimento — coluna
sem casos fica com % e A.R. vazios, não 0.

Dependências: só pandas/numpy/openpyxl; reaproveita a conta do resíduo ajustado
(`core.residuos_math`) para ficar idêntica à da Análise Residual.
"""
from __future__ import annotations

import io
import math
import re
import unicodedata
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from core.residuos_math import _adjusted_residuals

#: cruzamentos que entram marcados por padrão (só os que existirem no banco são usados)
CRUZAMENTOS_PADRAO: Tuple[str, ...] = ("P1", "P2", "P3", "P4", "P5", "juncaoreligiao", "P7", "P7_C", "VN")

#: limiar do resíduo ajustado para colorir (o mesmo da ferramenta «Análise Residual (A.R)»)
LIMIAR_Z_PADRAO = 1.85

FONTE_BANCO = "banco"
FONTE_TITULO = "titulo_total"
FONTE_SEM_LABEL = "sem_label"
ROTULO_FONTE = {FONTE_BANCO: "label do banco", FONTE_TITULO: "título do Total", FONTE_SEM_LABEL: "sem label (usa o nome)"}


# ---------------------------------------------------------------- utilidades
def canon(texto: Any) -> str:
    """Sem acento, minúsculo, só letras/números separados por espaço (para comparar textos)."""
    s = unicodedata.normalize("NFD", str(texto or "").strip())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _limpa(texto: str) -> str:
    return " ".join(str(texto or "").replace("''", "'").split())


# ---------------------------------------------------------------- 1. Total .sps
@dataclass
class TabelaTotal:
    var: str                 # nome da variável, sem o «$» das múltiplas
    titulo: str = ""         # texto depois de «Titulo:»
    pergunta: str = ""       # texto depois de «Pergunta:» (CAPTION)
    multipla: bool = False   # conjunto de múltipla resposta ($Q4R...): não existe como coluna do banco
    linha: int = 0           # linha do primeiro bloco no arquivo


_RE_INICIO = re.compile(r"^[ \t]*(CTABLES|TABLES)\b", re.I | re.M)
_RE_TITULO = re.compile(r"\bTITLE\s*=?\s*'((?:[^']|'')*)'", re.I)
_RE_CAPTION = re.compile(r"\bCAPTION\s*=?\s*'((?:[^']|'')*)'", re.I)
_RE_VAR_TABLES = re.compile(r"/TABLE\s*=\s*([A-Za-z_@#][\w.@#$]*)", re.I)
_RE_VAR_CTABLES = re.compile(r"/TABLE\s+(\$?[A-Za-z_@#][\w.@#$]*)", re.I)
_RE_PREFIXO_TITULO = re.compile(r"^\s*t[ií]tulo\s*:\s*", re.I)
_RE_PREFIXO_PERGUNTA = re.compile(r"^\s*pergunta\s*:\s*", re.I)


def _fim_comando(texto: str, pos: int) -> int:
    """Índice logo depois do ponto que fecha o comando: ponto fora de aspas e seguido só de
    espaços até o fim da linha (assim «F5.1» e «Dep.» no meio do texto não confundem)."""
    aspas: Optional[str] = None
    n = len(texto)
    i = pos
    while i < n:
        ch = texto[i]
        if aspas:
            if ch == aspas:
                if i + 1 < n and texto[i + 1] == aspas:
                    i += 1
                else:
                    aspas = None
        elif ch in "'\"":
            aspas = ch
        elif ch == ".":
            j = i + 1
            while j < n and texto[j] in " \t\r":
                j += 1
            if j >= n or texto[j] == "\n":
                return i + 1
        i += 1
    return n


def separar_total(texto: str) -> Tuple[str, bool]:
    """Devolve (trecho do TOTAL, tem_completo). Sem cabeçalho «COMPLETO», tudo é Total."""
    m = re.search(r"(?im)^\s*\*+\s*COMPLETO\s*\*+\s*$", texto)
    if m:
        return texto[:m.start()], True
    m = re.search(r"(?i)\*{5,}[^\n]*COMPLETO[^\n]*\*{5,}", texto)
    if m:
        return texto[:m.start()], True
    return texto, False


def ler_perguntas_total(texto: str) -> List[TabelaTotal]:
    """Variáveis que viram tabela no TOTAL, na ordem em que aparecem, sem repetir.

    Uma mesma variável aparece em várias tabelas do Total («Percentual %», «Valor Bruto»,
    «x Federações/Partidos»): vale a primeira, completando título/pergunta se faltarem.
    Tabelas de múltipla resposta (`CTABLES ... $Q4R`) entram marcadas como `multipla`.
    """
    total, _ = separar_total(texto.replace("\r\n", "\n").replace("\r", "\n"))
    achadas: Dict[str, TabelaTotal] = {}
    ordem: List[str] = []
    ultimo_fim = -1
    for m in _RE_INICIO.finditer(total):
        if m.start() < ultimo_fim:
            continue
        fim = _fim_comando(total, m.end())
        ultimo_fim = fim
        bloco = total[m.start():fim]
        tipo = m.group(1).upper()
        mv = (_RE_VAR_CTABLES if tipo == "CTABLES" else _RE_VAR_TABLES).search(bloco)
        if not mv:
            continue
        bruto = mv.group(1)
        if bruto.lower().startswith("$t"):   # variável de total («$tQ1_P»), não é pergunta
            continue
        multipla = bruto.startswith("$")
        var = bruto.lstrip("$")
        mt, mc = _RE_TITULO.search(bloco), _RE_CAPTION.search(bloco)
        titulo = _RE_PREFIXO_TITULO.sub("", _limpa(mt.group(1))) if mt else ""
        pergunta = _RE_PREFIXO_PERGUNTA.sub("", _limpa(mc.group(1))) if mc else ""
        linha = total.count("\n", 0, m.start()) + 1
        if var not in achadas:
            achadas[var] = TabelaTotal(var, titulo, pergunta, multipla, linha)
            ordem.append(var)
        else:
            t = achadas[var]
            t.titulo = t.titulo or titulo
            t.pergunta = t.pergunta or pergunta
    return [achadas[v] for v in ordem]


# ---------------------------------------------------------------- 2. labels
@dataclass
class LabelVar:
    var: str
    label_banco: str = ""
    titulo_total: str = ""
    label_usada: str = ""
    fonte: str = FONTE_SEM_LABEL

    @property
    def difere(self) -> bool:
        """Banco e Total têm texto e eles dizem coisas diferentes."""
        return bool(self.label_banco and self.titulo_total and canon(self.label_banco) != canon(self.titulo_total))


def coluna_do_banco(var: str, colunas: Iterable[str]) -> Optional[str]:
    """Nome real da coluna (o SPSS não diferencia maiúsculas de minúsculas)."""
    alvo = str(var).lstrip("$").lower()
    for c in colunas:
        if str(c).lower() == alvo:
            return str(c)
    return None


def resolver_labels(
    tabelas: Sequence[TabelaTotal],
    labels_banco: Dict[str, str],
    *,
    preferir_titulo: bool = False,
) -> List[LabelVar]:
    """Label de cada pergunta: a do banco; se não tiver, o título da tabela do Total.

    `preferir_titulo=True` inverte a prioridade (usa o título do Total sempre que ele existir).
    """
    por_minuscula = {str(k).lower(): (v or "") for k, v in labels_banco.items()}
    saida: List[LabelVar] = []
    for t in tabelas:
        lb = _limpa(por_minuscula.get(t.var.lower(), ""))
        tt = _limpa(t.titulo)
        if preferir_titulo and tt:
            usada, fonte = tt, FONTE_TITULO
        elif lb:
            usada, fonte = lb, FONTE_BANCO
        elif tt:
            usada, fonte = tt, FONTE_TITULO
        else:
            usada, fonte = t.var, FONTE_SEM_LABEL
        saida.append(LabelVar(t.var, lb, tt, usada, fonte))
    return saida


def cruzamentos_padrao_presentes(colunas: Iterable[str], padrao: Sequence[str] = CRUZAMENTOS_PADRAO) -> Tuple[List[str], List[str]]:
    """(presentes no banco, com o nome real da coluna; ausentes, com o nome pedido)."""
    cols = list(colunas)
    presentes, ausentes = [], []
    for p in padrao:
        real = coluna_do_banco(p, cols)
        (presentes if real else ausentes).append(real or p)
    return presentes, ausentes


# ---------------------------------------------------------------- 3. cálculo
@dataclass
class Cruzamento:
    linha: str
    coluna: str
    rot_linhas: List[str]
    rot_colunas: List[str]
    contagem: np.ndarray          # células ponderadas e arredondadas (r x c)
    pct_coluna: np.ndarray        # % da coluna; NaN se a coluna não tem casos
    z: np.ndarray                 # resíduo ajustado; NaN se não dá para calcular
    base_colunas: np.ndarray      # soma das células de cada coluna (ponderada)
    n_colunas: np.ndarray         # casos sem peso em cada coluna
    total_linhas: np.ndarray      # soma de cada linha (coluna «Total»)
    pct_total_linhas: np.ndarray  # % de cada linha no total geral
    base_total: float
    n_total: int
    casos_sem_peso: int = 0       # casos válidos que saíram por peso ausente ou ≤ 0


def _arredonda(v: float) -> float:
    return float(Decimal(str(float(v))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def mapa_rotulos(value_labels: Optional[Dict[Any, str]]) -> Dict[Any, str]:
    out: Dict[Any, str] = {}
    for k, v in (value_labels or {}).items():
        try:
            out[float(k)] = str(v)
        except (TypeError, ValueError):
            out[str(k)] = str(v)
    return out


def rotulo_categoria(cod: Any, mapa: Dict[Any, str]) -> str:
    try:
        chave: Any = float(cod)
    except (TypeError, ValueError):
        chave = str(cod)
    if chave in mapa:
        return mapa[chave]
    if isinstance(chave, float) and chave.is_integer():
        return str(int(chave))
    return str(cod)


def cruzar(
    df: pd.DataFrame,
    linha: str,
    coluna: str,
    peso: Optional[str] = None,
    value_labels: Optional[Dict[str, Dict[Any, str]]] = None,
) -> Optional[Cruzamento]:
    """Cruza `linha` (a pergunta) por `coluna` (segmento). None se não sobrar nenhum caso."""
    vl = value_labels or {}
    cols = [linha, coluna] + ([peso] if peso else [])
    sub = df[cols].dropna(subset=[linha, coluna])
    sem_peso = 0
    if peso:
        w = pd.to_numeric(sub[peso], errors="coerce")
        valido = w.notna() & (w > 0)
        sem_peso = int((~valido).sum())
        sub = sub[valido]
        w = w[valido]
    else:
        w = pd.Series(1.0, index=sub.index)
    if sub.empty:
        return None

    base = pd.DataFrame({"l": sub[linha].to_numpy(), "c": sub[coluna].to_numpy(), "w": w.to_numpy(dtype=float)})
    ct = base.groupby(["l", "c"])["w"].sum().unstack(fill_value=0.0)
    nn = base.groupby(["l", "c"])["w"].size().unstack(fill_value=0)
    ct = ct.apply(lambda s: s.map(_arredonda))     # /COUNT ROUND CELL
    # Toda categoria que aparece aqui tem ao menos um caso válido: como no SPSS, ela fica na tabela mesmo
    # quando o arredondamento deixa a contagem em 0 (o % sai 0 e o resíduo, que não existe, fica vazio).
    if ct.shape[0] == 0 or ct.shape[1] == 0:
        return None
    nn = nn.reindex(index=ct.index, columns=ct.columns, fill_value=0)

    obs = ct.to_numpy(dtype=float)
    soma_col = obs.sum(axis=0)
    soma_lin = obs.sum(axis=1)
    total = float(obs.sum())
    with np.errstate(divide="ignore", invalid="ignore"):
        pct = np.where(soma_col[None, :] > 0, obs / soma_col[None, :] * 100.0, np.nan)
        pct_tot = soma_lin / total * 100.0 if total > 0 else np.full(len(soma_lin), np.nan)

    m_l, m_c = mapa_rotulos(vl.get(linha)), mapa_rotulos(vl.get(coluna))
    return Cruzamento(
        linha=linha,
        coluna=coluna,
        rot_linhas=[rotulo_categoria(c, m_l) for c in ct.index],
        rot_colunas=[rotulo_categoria(c, m_c) for c in ct.columns],
        contagem=obs,
        pct_coluna=pct,
        z=_adjusted_residuals(obs),
        base_colunas=soma_col,
        n_colunas=nn.to_numpy(dtype=int).sum(axis=0),
        total_linhas=soma_lin,
        pct_total_linhas=pct_tot,
        base_total=total,
        n_total=int(nn.to_numpy().sum()),
        casos_sem_peso=sem_peso,
    )


@dataclass
class ResultadoAnalise:
    por_pergunta: Dict[str, List[Cruzamento]] = field(default_factory=dict)
    sem_dados: List[Tuple[str, str]] = field(default_factory=list)   # (pergunta, cruzamento) sem nenhum caso

    @property
    def n_tabelas(self) -> int:
        return sum(len(v) for v in self.por_pergunta.values())


def calcular_analise(
    df: pd.DataFrame,
    perguntas: Sequence[str],
    cruzamentos: Sequence[str],
    peso: Optional[str] = None,
    value_labels: Optional[Dict[str, Dict[Any, str]]] = None,
) -> ResultadoAnalise:
    res = ResultadoAnalise()
    for p in perguntas:
        lista: List[Cruzamento] = []
        for c in cruzamentos:
            if str(c).lower() == str(p).lower():
                continue
            cz = cruzar(df, p, c, peso, value_labels)
            if cz is None:
                res.sem_dados.append((p, c))
            else:
                lista.append(cz)
        res.por_pergunta[p] = lista
    return res


def tabela_para_tela(cz: Cruzamento) -> pd.DataFrame:
    """Uma linha por (categoria, estatística) para pré-visualizar na tela."""
    colunas = ["Total"] + list(cz.rot_colunas)
    linhas, idx = [], []
    for i, rot in enumerate(cz.rot_linhas):
        linhas.append([cz.total_linhas[i], *cz.contagem[i]])
        idx.append((rot, "Contagem"))
        linhas.append([cz.pct_total_linhas[i], *cz.pct_coluna[i]])
        idx.append((rot, "% da coluna"))
        linhas.append([np.nan, *cz.z[i]])
        idx.append((rot, "Resíduo ajustado"))
    linhas.append([cz.base_total, *cz.base_colunas])
    idx.append(("Base (ponderada)", ""))
    linhas.append([cz.n_total, *cz.n_colunas])
    idx.append(("Base (sem peso)", ""))
    return pd.DataFrame(linhas, index=pd.MultiIndex.from_tuples(idx), columns=colunas)


# ---------------------------------------------------------------- 4. Excel
# Layout igual ao da exportação do SPSS já formatada do Instituto (uma aba só, todas as tabelas
# empilhadas): título «<Pergunta> * <Segmento> Crosstabulation», uma linha «% within <Segmento>» e uma
# «Adjusted Residual» por alternativa, linha «Total» com 100 % e coluna «Total» no fim. Sem linhas
# de contagem. O destaque é o do VBA «DestacarAdjustedResiduals»: resíduo ≤ −limiar em vermelho
# #C00000 e ≥ +limiar em azul #002060, fonte branca em negrito, pintando também o % da mesma
# alternativa (a linha logo acima do resíduo).
COR_NEGATIVO = "FFC00000"
COR_POSITIVO = "FF002060"
NOME_ABA = "Análise das Perguntas"
_LARG_TEXTO = 22.7      # colunas A, B e C
_LARG_SEGMENTO = 13.7   # colunas de segmento e «Total»
_CHARS_TEXTO = 25       # caracteres Arial 9 que cabem por linha em 22,7
_CHARS_SEG = 13


def _lado(estilo: Optional[str]) -> Side:
    return Side(style=estilo) if estilo else Side()


def _b(esq=None, dir_=None, topo=None, base=None) -> Border:
    return Border(left=_lado(esq), right=_lado(dir_), top=_lado(topo), bottom=_lado(base))


def _linhas(texto: Any, chars_por_linha: float) -> int:
    t = str(texto or "")
    return max(1, sum(max(1, math.ceil(len(parte) / max(chars_por_linha, 1))) for parte in t.split("\n")))


def _altura(linhas: int) -> float:
    return 15.0 if linhas <= 1 else (28.05 if linhas == 2 else (45.0 if linhas == 3 else 14.0 * linhas))


def _num(v: Any) -> Optional[float]:
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    return x if np.isfinite(x) else None


def _cor_destaque(z: Optional[float], limiar: float) -> Optional[str]:
    if z is None:
        return None
    if z <= -limiar:
        return COR_NEGATIVO
    if z >= limiar:
        return COR_POSITIVO
    return None


def _bloco_excel(ws, topo: int, cz: Cruzamento, rot_pergunta: str, rot_cruz: str, limiar: float,
                 contagem: bool = False) -> int:
    """Escreve um cruzamento a partir da linha `topo`; devolve a primeira linha livre depois dele
    (já com uma linha em branco de separação). `contagem=True` acrescenta a linha «Count» (3 linhas
    por alternativa, como no CROSSTABS do SPSS); o destaque só pinta o % e o resíduo."""
    n = cz.contagem.shape[1]
    n_lin = len(cz.rot_linhas)
    c_seg, c_tot = 4, 4 + n
    r_tit, r_h1, r_h2, r_d = topo, topo + 1, topo + 2, topo + 3
    n_por = 3 if contagem else 2      # linhas por alternativa
    off = 1 if contagem else 0        # deslocamento da linha «% within»
    r_total = r_d + n_por * n_lin
    rot_within = f"% within {rot_cruz}"

    f_txt = Font(name="Arial", size=9)
    f_tit = Font(name="Arial", size=9, bold=True)
    f_num = Font(name="Calibri", size=10)
    f_dest = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    a_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    a_centro_top = Alignment(horizontal="center", wrap_text=True)
    a_texto = Alignment(horizontal="left", vertical="top", wrap_text=True)
    a_num = Alignment(horizontal="right", vertical="top")

    # ---- alturas
    larg_total = 3 * _LARG_TEXTO + (n + 1) * _LARG_SEGMENTO
    l_tit = _linhas(f"{rot_pergunta} * {rot_cruz} Crosstabulation", larg_total * 0.95)
    ws.row_dimensions[r_tit].height = 18.0 if l_tit <= 1 else (31.05 if l_tit == 2 else 45.0)
    ws.row_dimensions[r_h1].height = _altura(_linhas(rot_cruz, n * _CHARS_SEG))
    ws.row_dimensions[r_h2].height = _altura(max((_linhas(r, _CHARS_SEG) for r in cz.rot_colunas), default=1))
    l_within = _linhas(rot_within, _CHARS_TEXTO)
    alturas_pct, alturas_ar = [], []
    for rot in cz.rot_linhas:
        h = max(_altura(l_within), _altura(_linhas(rot, _CHARS_TEXTO)) - 15.0 * (n_por - 1), 15.0)
        alturas_pct.append(h)
        alturas_ar.append(15.0)
    falta = _altura(_linhas(rot_pergunta, _CHARS_TEXTO)) - (sum(alturas_pct) + sum(alturas_ar) + (15.0 * n_lin if contagem else 0.0))
    if falta > 0:
        alturas_ar[-1] += falta
    for i in range(n_lin):
        r0 = r_d + n_por * i
        if contagem:
            ws.row_dimensions[r0].height = 15.0
        ws.row_dimensions[r0 + off].height = alturas_pct[i]
        ws.row_dimensions[r0 + off + 1].height = alturas_ar[i]
    ws.row_dimensions[r_total].height = _altura(l_within)

    # ---- título (A:Total)
    ws.merge_cells(start_row=r_tit, start_column=1, end_row=r_tit, end_column=c_tot)
    t = ws.cell(r_tit, 1, f"{rot_pergunta} * {rot_cruz} Crosstabulation")
    t.font, t.alignment = f_tit, a_centro
    for k in range(1, c_tot + 1):
        ws.cell(r_tit, k).border = _b(base="medium")

    # ---- cabeçalho: segmento (mesclado) + Total (mesclado em 2 linhas)
    ws.merge_cells(start_row=r_h1, start_column=c_seg, end_row=r_h1, end_column=c_tot - 1)
    h = ws.cell(r_h1, c_seg, rot_cruz)
    h.font, h.alignment = f_txt, a_centro_top
    ws.cell(r_h1, 3).border = _b(dir_="medium")
    for k in range(c_seg, c_tot):
        ws.cell(r_h1, k).border = _b(dir_="thin" if k in (c_seg, c_tot - 1) else None, base="thin")
    ws.merge_cells(start_row=r_h1, start_column=c_tot, end_row=r_h2, end_column=c_tot)
    h = ws.cell(r_h1, c_tot, "Total")
    h.font, h.alignment = f_txt, a_centro_top
    ws.cell(r_h1, c_tot).border = _b("thin", "medium", "medium", "thin")
    ws.cell(r_h2, c_tot).border = _b("thin", "medium", None, "thin")
    for k in (1, 2):
        ws.cell(r_h2, k).border = _b(base="medium")
    ws.cell(r_h2, 3).border = _b(dir_="medium", base="medium")
    for j, rot in enumerate(cz.rot_colunas):
        c = ws.cell(r_h2, c_seg + j, rot)
        c.font, c.alignment = f_txt, a_centro_top
        c.border = _b("medium" if j == 0 else "thin", "thin", "thin", "medium")

    # ---- alternativas: [Count] + «% within» + «Adjusted Residual»
    ws.merge_cells(start_row=r_d, start_column=1, end_row=r_total - 1, end_column=1)
    a = ws.cell(r_d, 1, rot_pergunta)
    a.font, a.alignment = f_txt, a_texto
    for i, rot in enumerate(cz.rot_linhas):
        r0 = r_d + n_por * i
        rp, ra = r0 + off, r0 + off + 1
        primeira, ultima = i == 0, i == n_lin - 1
        ws.merge_cells(start_row=r0, start_column=2, end_row=ra, end_column=2)
        b = ws.cell(r0, 2, rot)
        b.font, b.alignment = f_txt, a_texto
        for r in range(r0, ra + 1):
            ws.cell(r, 2).border = _b(base="thin")
            ws.cell(r, 1).border = _b("medium", topo="medium" if (primeira and r == r0) else None,
                                      base="thin" if (ultima and r == ra) else None)
        linhas = ([(r0, "Count")] if contagem else []) + [(rp, rot_within), (ra, "Adjusted Residual")]
        for r, texto in linhas:
            c = ws.cell(r, 3, texto)
            c.font, c.alignment = f_txt, a_texto
            c.border = _b(dir_="medium", topo=("medium" if primeira else "thin") if r == r0 else None,
                          base="thin" if r == ra else None)

        valores_cnt = [_num(cz.contagem[i, j]) for j in range(n)] + [_num(cz.total_linhas[i])]
        valores_pct = [_num(cz.pct_coluna[i, j]) for j in range(n)] + [_num(cz.pct_total_linhas[i])]
        valores_z = [_num(cz.z[i, j]) for j in range(n)] + [None]
        for j in range(n + 1):
            col = c_seg + j
            eh_total = j == n
            esq = "thin" if (j > 0 or eh_total) else None
            dir_ = "medium" if eh_total else "thin"
            tipos = ([(r0, valores_cnt[j], "###0")] if contagem else []) + [(rp, valores_pct[j], "###0.0%"), (ra, valores_z[j], "###0.0")]
            for r, v, fmt in tipos:
                c = ws.cell(r, col)
                c.font, c.alignment, c.number_format = f_num, a_num, fmt
                if v is not None:
                    c.value = v / 100.0 if fmt == "###0.0%" else v
                topo_b = ("medium" if primeira else "thin") if (eh_total and r == r0) else None
                c.border = _b(esq, dir_, topo_b, "thin" if r == ra else None)
            cor = _cor_destaque(valores_z[j], limiar)
            if cor:
                for r in (rp, ra):
                    c = ws.cell(r, col)
                    c.fill = PatternFill("solid", fgColor=cor)
                    c.font = f_dest

    # ---- linha Total ([Count] + «% within» = 100 %)
    k_tot = 2 if contagem else 1
    ws.merge_cells(start_row=r_total, start_column=1, end_row=r_total + k_tot - 1, end_column=2)
    a = ws.cell(r_total, 1, "Total")
    a.font, a.alignment = f_txt, a_texto
    for r in range(r_total, r_total + k_tot):
        ws.cell(r, 1).border = _b("medium", topo="thin" if r == r_total else None, base="medium" if r == r_total + k_tot - 1 else None)
        ws.cell(r, 2).border = _b(topo="thin" if r == r_total else None, base="medium" if r == r_total + k_tot - 1 else None)
    linhas_tot = ([(r_total, "Count", [_num(x) for x in cz.base_colunas] + [_num(cz.base_total)], "###0")] if contagem else []) + \
                 [(r_total + k_tot - 1, rot_within, [1.0] * (n + 1), "###0.0%")]
    for r, texto, vals, fmt in linhas_tot:
        c = ws.cell(r, 3, texto)
        c.font, c.alignment = f_txt, a_texto
        c.border = _b(dir_="medium", topo="thin" if r == r_total else None, base="medium" if r == r_total + k_tot - 1 else None)
        for j in range(n + 1):
            c = ws.cell(r, c_seg + j, vals[j])
            c.font, c.alignment, c.number_format = f_num, a_num, fmt
            c.border = _b("medium" if j == 0 else "thin", "medium" if j == n else "thin",
                          "thin" if r == r_total else None, "medium" if r == r_total + k_tot - 1 else None)
    return r_total + k_tot + 1


def _ajusta_larguras(ws, n_segmentos: int) -> None:
    for k in (1, 2, 3):
        ws.column_dimensions[get_column_letter(k)].width = _LARG_TEXTO
    for k in range(4, 4 + n_segmentos + 1):
        ws.column_dimensions[get_column_letter(k)].width = _LARG_SEGMENTO


def _texto_formatado(valor: float, formato: str) -> str:
    """Como o Excel mostra o número (vírgula decimal): «###0» inteiro, «###0.0» 1 casa, «###0.0%» percentual."""
    if formato == "###0.0%":
        return f"{valor * 100:.1f}".replace(".", ",") + "%"
    if formato == "###0.0":
        return f"{valor:.1f}".replace(".", ",")
    return f"{valor:.0f}"


def bloco_como_planilha(cz: Cruzamento, rot_pergunta: str, rot_cruz: str, limiar: float = LIMIAR_Z_PADRAO,
                        contagem: bool = False, como_texto: bool = False):
    """Planilha (openpyxl) só com a tabela de um cruzamento, idêntica ao bloco que vai para o Excel.
    Serve à pré-visualização da tela: o que se vê é o que sai no arquivo.
    `como_texto=True` troca cada número pelo texto já formatado (o renderizador HTML da tela não aplica
    o formato de número das células)."""
    wb = Workbook()
    ws = wb.active
    ws.title = NOME_ABA
    _bloco_excel(ws, 1, cz, rot_pergunta, rot_cruz, limiar, contagem)
    _ajusta_larguras(ws, cz.contagem.shape[1])
    if como_texto:
        for linha in ws.iter_rows():
            for c in linha:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    c.value = _texto_formatado(float(c.value), c.number_format)
    return ws


def montar_excel(
    resultado: ResultadoAnalise,
    labels: Sequence[LabelVar],
    labels_cruzamento: Dict[str, str],
    *,
    peso: Optional[str],
    limiar_z: float = LIMIAR_Z_PADRAO,
    perguntas_total: Sequence[TabelaTotal] = (),
    nome_banco: str = "",
    nome_syntax: str = "",
    avisos: Sequence[str] = (),
    incluir_contagem: bool = False,
) -> bytes:
    """Excel final: aba «Análise das Perguntas» (todas as tabelas empilhadas, já pintadas) + aba «Resumo»."""
    wb = Workbook()
    ws = wb.active
    ws.title = NOME_ABA
    por_var = {lv.var.lower(): lv for lv in labels}
    linha_inicial: Dict[str, int] = {}
    topo = 1
    n_max = 1
    for p, lista in resultado.por_pergunta.items():
        lv = por_var.get(p.lower()) or LabelVar(p, label_usada=p)
        for k, cz in enumerate(lista):
            if k == 0:
                linha_inicial[p] = topo
            n_max = max(n_max, cz.contagem.shape[1])
            topo = _bloco_excel(ws, topo, cz, lv.label_usada, labels_cruzamento.get(cz.coluna) or cz.coluna, limiar_z, incluir_contagem)
    _ajusta_larguras(ws, n_max)

    res = wb.create_sheet("Resumo")
    f_b, f_n = Font(name="Arial", size=10, bold=True), Font(name="Arial", size=10)
    res["A1"] = "Análise das Perguntas"
    res["A1"].font = Font(name="Arial", size=15, bold=True)
    info = [
        ("Banco", nome_banco), ("Programação (Total)", nome_syntax),
        ("Peso", peso or "sem peso"), ("Resíduo ajustado pintado a partir de ±", limiar_z),
        ("Tabelas geradas", resultado.n_tabelas),
        ("Linha Count", "incluída" if incluir_contagem else "não incluída"),
        ("Destaque", "vermelho #C00000 (≤ −limiar) e azul #002060 (≥ +limiar), no % e no resíduo da alternativa"),
    ]
    for i, (a, b) in enumerate(info, start=3):
        res.cell(i, 1, a).font = f_b
        res.cell(i, 2, b).font = f_n
    linha = 3 + len(info) + 1
    cab = ["Variável", "Label usada", "Fonte do label", "Label no banco", "Título no Total", "Cruzamentos", "1ª linha na aba", "Observação"]
    for j, t in enumerate(cab, start=1):
        c = res.cell(linha, j, t)
        c.font, c.fill = Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="002060")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    linha += 1
    for p, lista in resultado.por_pergunta.items():
        lv = por_var.get(p.lower()) or LabelVar(p, label_usada=p)
        obs = ""
        if not lista:
            obs = "Nenhum cruzamento com casos."
        elif lv.fonte == FONTE_TITULO and lv.label_banco == "":
            obs = "Banco sem label: foi usado o título do Total."
        elif lv.difere:
            obs = "Label do banco difere do título do Total."
        valores = [p, lv.label_usada, ROTULO_FONTE.get(lv.fonte, ""), lv.label_banco, lv.titulo_total,
                   ", ".join(c.coluna for c in lista), linha_inicial.get(p), obs]
        for j, v in enumerate(valores, start=1):
            c = res.cell(linha, j, v)
            c.font, c.alignment = f_n, Alignment(wrap_text=True, vertical="top")
        linha += 1
    if resultado.sem_dados or avisos:
        linha += 1
        res.cell(linha, 1, "Avisos").font = Font(name="Arial", size=11, bold=True)
        linha += 1
        for a in avisos:
            res.cell(linha, 1, a).font = f_n
            linha += 1
        for p, c in resultado.sem_dados:
            res.cell(linha, 1, f"{p} × {c}: nenhum caso válido (não gerado).").font = f_n
            linha += 1
    for k, w in zip("ABCDEFGH", (14, 52, 20, 44, 52, 40, 14, 40)):
        res.column_dimensions[k].width = w
    res.freeze_panes = "A12"
    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()
