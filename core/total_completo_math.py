"""
Total e Completo — etapa 2 do "Tratamento do Banco" (lógica pura, sem Streamlit).

Lê a programação do Total e Completo (.sps) e calcula, no banco, as tabelas
que o SPSS calcularia:

- `TABLES ... cpct(var :total seg1 seg2)` + `UNW count`  → % da coluna (com o
  peso ligado) e a linha «Base» (casos sem peso);
- `CTABLES ... $MULTIPLA [C][COLPCT.COUNT ...]`          → % de casos de uma
  pergunta de múltipla resposta e a «Base» (casos com ao menos uma resposta).

Depois confere cada tabela (colunas somando 100%, variáveis que existem,
códigos sem rótulo, bases pequenas…) e, se o usuário mandar o Relatório
Descritivo, compara número a número.

Comando que o interpretador não conhece NÃO é calculado em silêncio: a tabela
fica marcada como «não suportada» e vira um aviso.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Sequence

import numpy as np
import pandas as pd

from core.analise_perguntas_math import (
    _RE_CAPTION,
    _RE_INICIO,
    _RE_PREFIXO_PERGUNTA,
    _RE_PREFIXO_TITULO,
    _RE_TITULO,
    _fim_comando,
    _limpa,
    canon,
    mapa_rotulos,
    rotulo_categoria,
)
from core.revisor_syntax_math import Finding, decode_sps

PESO_PADRAO = "peso"
BASE_PEQUENA = 30
TOLERANCIA_SOMA = 0.15           # colunas de TABLES somam 100 ± isso (arredondamento de 1 casa)
TOLERANCIA_COMPARACAO = 0.051    # diferença aceita contra o Descritivo (1 casa decimal)

SEGMENTO_TOTAL = "total"
LINHAS_SO_DO_DESCRITIVO = {"base reduzida"}      # acrescentada depois, pelo processo de base reduzida


# ---------------------------------------------------------------------------
# 1. Ler a programação
# ---------------------------------------------------------------------------

@dataclass
class TabelaSps:
    ordem: int
    secao: str                       # Total | Completo
    comando: str                     # TABLES | CTABLES
    var_linha: str                   # sem o «$»
    multipla: bool
    colunas: list[str]               # variáveis de coluna (inclui «total»)
    titulo: str
    pergunta: str
    linha: int
    suportada: bool = True
    motivo: str = ""
    bruto: str = ""


_RE_TABLE_ANY = re.compile(r"/TABLE\b\s*=?\s*([^\n]+?)\s+BY\s+([^\n]+)", re.I)
_RE_CPCT = re.compile(r"cpct\(\s*(\$?[\w@#.$]+)\s*\(\s*F\d+\.\d+\s*\)\s*(?:\"[^\"]*\"|'[^']*')?\s*:\s*([^)]*)\)", re.I)
_RE_UNW = re.compile(r"\bUNW\s+count\(\s*(\$[\w@#.$]+)", re.I)


def _sem_colchetes(txt: str) -> str:
    out, prof = [], 0
    for ch in txt:
        if ch == "[":
            prof += 1
        elif ch == "]":
            prof = max(prof - 1, 0)
        elif prof == 0:
            out.append(ch)
    return "".join(out)


def _limpa_coluna(txt: str) -> list[str]:
    return [p.strip() for p in _sem_colchetes(txt).split("+") if p.strip()]


def ler_tabelas(sps_bytes: bytes) -> tuple[list[TabelaSps], list[Finding]]:
    texto, _enc, achados = decode_sps(sps_bytes)
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")
    findings: list[Finding] = list(achados)

    m_comp = re.search(r"(?im)^\s*\*+\s*COMPLETO\s*\*+\s*$", texto) or re.search(r"(?i)\*{5,}[^\n]*COMPLETO[^\n]*\*{5,}", texto)
    pos_completo = m_comp.start() if m_comp else len(texto) + 1

    tabelas: list[TabelaSps] = []
    ultimo_fim = -1
    for m in _RE_INICIO.finditer(texto):
        if m.start() < ultimo_fim:
            continue
        fim = _fim_comando(texto, m.end())
        ultimo_fim = fim
        bloco = texto[m.start():fim]
        comando = m.group(1).upper()
        secao = "Total" if m.start() < pos_completo else "Completo"
        linha = texto.count("\n", 0, m.start()) + 1
        mt, mc = _RE_TITULO.search(bloco), _RE_CAPTION.search(bloco)
        titulo = _RE_PREFIXO_TITULO.sub("", _limpa(mt.group(1))) if mt else ""
        pergunta = _RE_PREFIXO_PERGUNTA.sub("", _limpa(mc.group(1))) if mc else ""

        mtab = _RE_TABLE_ANY.search(bloco)
        if not mtab:
            tabelas.append(TabelaSps(len(tabelas) + 1, secao, comando, "?", False, [], titulo, pergunta, linha,
                                     False, "não achei o /TABLE com «BY»", bloco))
            continue
        esq, dir_ = mtab.group(1), mtab.group(2)
        linha_vars = _limpa_coluna(esq)
        bruto_var = linha_vars[0] if linha_vars else "?"
        multipla = bruto_var.startswith("$") and not bruto_var.lower().startswith("$t")
        var_linha = bruto_var.lstrip("$")
        colunas = _limpa_coluna(dir_)
        t = TabelaSps(len(tabelas) + 1, secao, comando, var_linha, multipla, colunas, titulo, pergunta, linha, True, "", bloco)

        if ">" in esq or ">" in dir_:
            t.suportada, t.motivo = False, "tabela com aninhamento (>) ainda não é calculada"
        elif comando == "TABLES":
            mp = _RE_CPCT.search(bloco)
            if not mp:
                t.suportada, t.motivo = False, "não achei a estatística cpct(...) — só calculo % da coluna"
            else:
                segs = [s for s in re.split(r"\s+", mp.group(2).strip()) if s]
                if [s.lower() for s in segs] != [c.lower() for c in colunas]:
                    t.suportada = False
                    t.motivo = "as colunas do cpct não são as mesmas do BY"
            outras = re.findall(r"(?i)\b(mean|median|stddev|sum|ptile|mode|min|max|spct|rpct|tpct)\s*\(", bloco)
            if outras and t.suportada:
                t.suportada, t.motivo = False, f"estatística não suportada: {outras[0]}"
        else:
            if "COLPCT.COUNT" not in bloco.upper():
                t.suportada, t.motivo = False, "CTABLES sem COLPCT.COUNT — só calculo esse tipo"
            if not multipla and t.suportada:
                t.suportada, t.motivo = False, "CTABLES de pergunta simples ainda não é calculado"
        tabelas.append(t)
        if not t.suportada:
            findings.append(Finding(
                "possivel_erro", "TC_TABELA_NAO_SUPORTADA", f"Tabela não calculada: {titulo or var_linha}",
                f"A tabela da linha {linha} ({var_linha}) não foi calculada: {t.motivo}.",
                linha, var_linha, " ".join(bloco.split())[:300], "Confira essa tabela direto no SPSS.",
            ))
    return tabelas, findings


# ---------------------------------------------------------------------------
# 2. Calcular
# ---------------------------------------------------------------------------

@dataclass
class ColunaCalc:
    grupo: str            # nome do segmento (rótulo da variável) — "Total" para a coluna total
    rotulo: str           # categoria do segmento
    var: str
    codigo: Any
    base: float           # linha «Base»
    n: int                # casos sem peso


@dataclass
class TabelaCalc:
    tabela: TabelaSps
    colunas: list[ColunaCalc]
    rotulos: list[str]                # categorias (linhas)
    codigos: list[Any]
    valores: np.ndarray               # linhas x colunas, % (NaN = coluna sem casos)
    contagens: np.ndarray             # linhas x colunas, soma dos pesos
    erro: str = ""


class ErroTabela(Exception):
    pass


def _col(df: pd.DataFrame, nome: str) -> str | None:
    n = nome.upper()
    for c in df.columns:
        if str(c).upper() == n:
            return c
    return None


def _pesos(df: pd.DataFrame, peso: str | None) -> pd.Series:
    if not peso:
        return pd.Series(1.0, index=df.index)
    c = _col(df, peso)
    if c is None:
        raise ErroTabela(f"a variável de peso «{peso}» não existe no banco")
    w = pd.to_numeric(df[c], errors="coerce")
    return w.where(w > 0)             # sem peso ou ≤ 0: o SPSS tira o caso de tudo


def _categorias(serie: pd.Series, valido: pd.Series) -> list[Any]:
    vals = serie[valido & serie.notna()].unique()
    try:
        return sorted(vals.tolist())
    except TypeError:
        return sorted(vals.tolist(), key=str)


def _rotulo_var(nome: str, var_labels: dict[str, str]) -> str:
    """Rótulo da variável; vazio se não tiver (nesse caso o grupo da coluna é o próprio rótulo da categoria)."""
    return (var_labels.get(nome.upper(), "") or "").strip()


def _vazias_incluidas(bloco: str) -> set[str]:
    """Variáveis cujas categorias SEM casos também aparecem (CTABLES /CATEGORIES ... EMPTY=INCLUDE)."""
    out: set[str] = set()
    for m in re.finditer(r"/CATEGORIES\s+VARIABLES\s*=\s*([^\n]*?)(?=\s+(?:ORDER|KEY|EMPTY|TOTAL|LABEL|POSITION|MISSING)\b|\n|$)([^\n]*)", bloco, re.I):
        if re.search(r"EMPTY\s*=\s*INCLUDE", m.group(0), re.I):
            out.update(v.upper() for v in re.findall(r"\$?[\w@#.]+", m.group(1)))
    return out


def calcular_tabela(
    tab: TabelaSps,
    df: pd.DataFrame,
    value_labels: dict[str, dict[Any, str]],
    var_labels: dict[str, str],
    mrsets: dict[str, list[str]],
    peso: str | None = PESO_PADRAO,
) -> TabelaCalc:
    if not tab.suportada:
        raise ErroTabela(tab.motivo)
    vl = {str(k).upper(): v for k, v in (value_labels or {}).items()}
    lb = {str(k).upper(): v for k, v in (var_labels or {}).items()}
    w = _pesos(df, peso)
    valido_peso = w.notna()
    vazias = _vazias_incluidas(tab.bruto)

    # ---- variável(is) de linha -----------------------------------------
    if tab.multipla:
        membros = mrsets.get("$" + tab.var_linha.upper()) or mrsets.get(tab.var_linha.upper())
        if not membros:
            raise ErroTabela(f"o conjunto múltiplo ${tab.var_linha} não foi definido (MRSETS)")
        cols_m = []
        for m in membros:
            c = _col(df, m)
            if c is None:
                raise ErroTabela(f"a variável {m} do conjunto ${tab.var_linha} não existe no banco")
            cols_m.append(c)
        mat = df[cols_m].apply(pd.to_numeric, errors="coerce")
        valido_linha = mat.notna().any(axis=1)
        cats_set = {float(v) for v in pd.unique(mat.to_numpy().ravel()) if pd.notna(v)}
        rot_map = {}
        for m in membros:
            for k, v in mapa_rotulos(vl.get(m.upper())).items():
                rot_map.setdefault(k, v)
        if ("$" + tab.var_linha.upper()) in vazias or tab.var_linha.upper() in vazias:
            cats_set |= {k for k in rot_map if isinstance(k, float)}
        cats = sorted(cats_set)
        presente = {c: (mat == c).any(axis=1) for c in cats}
    else:
        c = _col(df, tab.var_linha)
        if c is None:
            raise ErroTabela(f"a variável {tab.var_linha} não existe no banco")
        serie = df[c]
        valido_linha = serie.notna() if pd.api.types.is_numeric_dtype(serie) else serie.astype(str).str.strip().ne("")
        cats = _categorias(serie, valido_peso & valido_linha)
        presente = {k: (serie == k) for k in cats}
        rot_map = mapa_rotulos(vl.get(tab.var_linha.upper()))

    base_linha = valido_peso & valido_linha

    # ---- colunas ---------------------------------------------------------
    colunas: list[ColunaCalc] = []
    mascaras: list[pd.Series] = []
    for seg in tab.colunas:
        if seg.lower() == SEGMENTO_TOTAL and _col(df, seg) is None:
            colunas.append(ColunaCalc("Total", "Total", seg, 1, 0.0, 0))
            mascaras.append(pd.Series(True, index=df.index))
            continue
        if seg.startswith("$"):
            membros_s = mrsets.get(seg.upper())
            if not membros_s:
                raise ErroTabela(f"o conjunto múltiplo {seg} (coluna) não foi definido (MRSETS)")
            cols_s = []
            for m in membros_s:
                cm = _col(df, m)
                if cm is None:
                    raise ErroTabela(f"a variável {m} do conjunto {seg} não existe no banco")
                cols_s.append(cm)
            mat_s = df[cols_s].apply(pd.to_numeric, errors="coerce")
            rot_s = {}
            for m in membros_s:
                for k, v in mapa_rotulos(vl.get(m.upper())).items():
                    rot_s.setdefault(k, v)
            grupo = _rotulo_var(seg, lb)
            vistos_s = {float(v) for v in pd.unique(mat_s[valido_peso.to_numpy()].to_numpy().ravel()) if pd.notna(v)}
            if seg.upper() in vazias:
                vistos_s |= {k for k in rot_s if isinstance(k, float)}
            for k in sorted(vistos_s):
                rot_k = rotulo_categoria(k, rot_s)
                colunas.append(ColunaCalc(grupo or rot_k, rot_k, seg, k, 0.0, 0))
                mascaras.append((mat_s == k).any(axis=1))
            continue
        cs = _col(df, seg)
        if cs is None:
            raise ErroTabela(f"a variável de coluna {seg} não existe no banco")
        sserie = df[cs]
        if seg.lower() == SEGMENTO_TOTAL:
            colunas.append(ColunaCalc("Total", "Total", seg, 1, 0.0, 0))
            mascaras.append(pd.Series(True, index=df.index))
            continue
        m_seg = mapa_rotulos(vl.get(seg.upper()))
        grupo = _rotulo_var(seg, lb)
        ks = set(_categorias(sserie, valido_peso))
        if seg.upper() in vazias:
            ks |= {k for k in m_seg if isinstance(k, float)}
        for k in sorted(ks, key=lambda x: (isinstance(x, str), x)):
            rot_k = rotulo_categoria(k, m_seg)
            colunas.append(ColunaCalc(grupo or rot_k, rot_k, seg, k, 0.0, 0))
            mascaras.append(sserie == k)

    contagens = np.zeros((len(cats), len(colunas)))
    valores = np.full((len(cats), len(colunas)), np.nan)
    for j, (col, mask) in enumerate(zip(colunas, mascaras)):
        casos = base_linha & mask
        wj = w.where(casos, 0.0).fillna(0.0)
        base_w = float(wj.sum())
        col.n = int(casos.sum())
        col.base = float(col.n) if tab.comando == "TABLES" else base_w
        for i, k in enumerate(cats):
            contagens[i, j] = float(wj[presente[k] & casos].sum())
        if base_w > 0:
            valores[:, j] = contagens[:, j] / base_w * 100.0

    rotulos = [rotulo_categoria(k, rot_map) for k in cats]
    return TabelaCalc(tab, colunas, rotulos, cats, valores, contagens)


@dataclass
class ResultadoTotalCompleto:
    tabelas: list[TabelaSps]
    calculadas: list[TabelaCalc | None]       # alinhada com `tabelas`
    erros: dict[int, str]                      # ordem → motivo
    findings: list[Finding]
    peso: str | None
    comparacao: list["ComparacaoTabela"] | None = None

    @property
    def n_calculadas(self) -> int:
        return sum(1 for c in self.calculadas if c is not None)


def calcular_total_e_completo(
    sps_bytes: bytes,
    df: pd.DataFrame,
    value_labels: dict[str, dict[Any, str]],
    var_labels: dict[str, str],
    mrsets: dict[str, list[str]] | None = None,
    peso: str | None = PESO_PADRAO,
) -> ResultadoTotalCompleto:
    tabelas, findings = ler_tabelas(sps_bytes)
    mrsets = {k.upper(): v for k, v in (mrsets or {}).items()}
    if not mrsets:
        mrsets = mrsets_do_texto(sps_bytes)
    calculadas: list[TabelaCalc | None] = []
    erros: dict[int, str] = {}
    for t in tabelas:
        try:
            calc = calcular_tabela(t, df, value_labels, var_labels, mrsets, peso)
            calculadas.append(calc)
            findings.extend(conferir_tabela(calc, value_labels))
        except ErroTabela as e:
            calculadas.append(None)
            if t.suportada:
                erros[t.ordem] = str(e)
                findings.append(Finding(
                    "critico", "TC_TABELA_COM_ERRO", f"Tabela sem cálculo: {t.titulo or t.var_linha}",
                    f"Não consegui calcular a tabela da linha {t.linha}: {e}.",
                    t.linha, t.var_linha, None, "Confira se o banco é o certo e se a Syntax foi rodada antes.",
                ))
    findings.extend(_conferir_repeticoes(tabelas))
    findings.extend(_conferir_segmentos([c for c in calculadas if c is not None]))
    findings.extend(_conferir_pesos(df, peso))
    return ResultadoTotalCompleto(tabelas, calculadas, erros, findings, peso)


def mrsets_do_texto(sps_bytes: bytes) -> dict[str, list[str]]:
    """MRSETS declarados dentro do próprio arquivo (/MCGROUP NAME=$X VARIABLES=...)."""
    from core.revisor_syntax_math import _parse_mcgroups
    texto, _e, _a = decode_sps(sps_bytes)
    return {k.upper(): v for k, v in _parse_mcgroups(texto).items()}


# ---------------------------------------------------------------------------
# 3. Conferências
# ---------------------------------------------------------------------------

def conferir_tabela(calc: TabelaCalc, value_labels: dict[str, dict[Any, str]]) -> list[Finding]:
    t = calc.tabela
    achados: list[Finding] = []
    nome = t.titulo or t.var_linha
    vl = {str(k).upper(): v for k, v in (value_labels or {}).items()}

    if not calc.codigos:
        achados.append(Finding(
            "possivel_erro", "TC_TABELA_VAZIA", f"Tabela vazia: {nome}",
            f"Nenhum caso válido para {t.var_linha}; a tabela da linha {t.linha} sai em branco.",
            t.linha, t.var_linha, None, "Confira a recodificação dessa variável.",
        ))
        return achados

    # código com dados e sem rótulo → a tabela mostraria o número
    mapa_ref = {} if t.multipla else mapa_rotulos(vl.get(t.var_linha.upper()))
    if not t.multipla and mapa_ref:
        sem = [k for k in calc.codigos if k not in mapa_ref and not (isinstance(k, float) and k in mapa_ref)]
        if sem:
            achados.append(Finding(
                "revisar", "TC_CODIGO_SEM_ROTULO", f"{t.var_linha}: código sem rótulo na tabela",
                f"A tabela «{nome}» (linha {t.linha}) tem casos com código(s) "
                f"{', '.join(_fmt(k) for k in sem[:8])} sem rótulo; a linha sairá com o número.",
                t.linha, t.var_linha, None, "Inclua o rótulo no VALUE LABELS.",
            ))
    elif not t.multipla and not mapa_ref:
        achados.append(Finding(
            "revisar", "TC_SEM_ROTULOS", f"{t.var_linha}: sem rótulos de valor",
            f"A tabela «{nome}» (linha {t.linha}) usa {t.var_linha}, que não tem VALUE LABELS.",
            t.linha, t.var_linha, None, "Inclua o VALUE LABELS dessa variável.",
        ))

    # soma das colunas (só TABLES; múltipla passa de 100)
    if not t.multipla:
        for j, col in enumerate(calc.colunas):
            if col.n == 0:
                continue
            soma = float(np.nansum(calc.valores[:, j]))
            if abs(soma - 100.0) > TOLERANCIA_SOMA:
                achados.append(Finding(
                    "critico", "TC_COLUNA_NAO_SOMA_100",
                    f"{t.var_linha} × {col.grupo}/{col.rotulo}: soma {soma:.1f}%",
                    f"Na tabela «{nome}» (linha {t.linha}) a coluna {col.grupo} · {col.rotulo} soma {soma:.1f}%.",
                    t.linha, t.var_linha, None, "Confira o RECODE e o peso.",
                ))

    # base pequena
    pequenas = [c for c in calc.colunas if 0 < c.n < BASE_PEQUENA]
    if pequenas:
        achados.append(Finding(
            "informativo", "TC_BASE_PEQUENA", f"{t.var_linha}: base menor que {BASE_PEQUENA}",
            f"Na tabela «{nome}» (linha {t.linha}) há colunas com base pequena: "
            + ", ".join(f"{c.grupo}/{c.rotulo} (n={c.n})" for c in pequenas[:8]) + ".",
            t.linha, t.var_linha, None, None,
        ))
    return achados


def _conferir_segmentos(calculadas: list[TabelaCalc]) -> list[Finding]:
    """Casos que ficam fora de um segmento (sem resposta na variável de segmento), uma linha por segmento."""
    por_var: dict[str, dict[str, Any]] = {}
    for c in calculadas:
        base_total = next((x.n for x in c.colunas if x.grupo == "Total"), None)
        if base_total is None or c.tabela.multipla:
            continue
        soma: dict[str, int] = {}
        for x in c.colunas:
            if x.grupo != "Total" and not x.var.startswith("$"):
                soma[x.var] = soma.get(x.var, 0) + x.n
        for v, n in soma.items():
            d = por_var.setdefault(v.upper(), {"var": v, "faltas": [], "tabelas": 0, "linha": c.tabela.linha})
            d["tabelas"] += 1
            d["faltas"].append(base_total - n)
    achados: list[Finding] = []
    for d in por_var.values():
        f = d["faltas"]
        if max(f) <= 0:
            continue
        achados.append(Finding(
            "informativo", "TC_SEGMENTO_NAO_FECHA", f"{d['var']}: casos fora do segmento",
            f"Em {sum(1 for x in f if x > 0)} de {d['tabelas']} tabela(s) por {d['var']}, a soma das colunas fica "
            f"{min(x for x in f if x > 0)} a {max(f)} caso(s) abaixo da base do total: são casos sem resposta em {d['var']}.",
            d["linha"], d["var"], None, "Normal se esses casos não responderam; senão, revise a recodificação.",
        ))
    return achados


def _conferir_pesos(df: pd.DataFrame, peso: str | None) -> list[Finding]:
    if not peso:
        return []
    c = _col(df, peso)
    if c is None:
        return []
    w = pd.to_numeric(df[c], errors="coerce")
    achados: list[Finding] = []
    sem = int((w.isna() | (w <= 0)).sum())
    if sem:
        achados.append(Finding(
            "revisar", "TC_CASOS_SEM_PESO", f"{sem} caso(s) sem peso",
            f"{sem} caso(s) têm {c} vazio ou ≤ 0 e ficam fora de TODAS as tabelas (o SPSS faz o mesmo).",
            None, c, None, "Confira se esses casos deviam ter peso.",
        ))
    validos = w[w > 0]
    if len(validos) and abs(float(validos.sum()) / len(validos) - 1) > 0.02:
        achados.append(Finding(
            "informativo", "TC_PESO_NAO_NORMALIZADO", f"Os pesos somam {validos.sum():.1f} para {len(validos)} casos",
            f"A soma de {c} é {validos.sum():.1f} e há {len(validos)} casos com peso: o peso médio é "
            f"{validos.mean():.3f}, não 1. As bases «Base» (sem peso) continuam as mesmas, mas o CTABLES usa a base ponderada.",
            None, c, None, None,
        ))
    return achados


def _conferir_repeticoes(tabelas: list[TabelaSps]) -> list[Finding]:
    achados: list[Finding] = []
    vistos: dict[tuple, TabelaSps] = {}
    for t in tabelas:
        chave = (t.secao, t.var_linha.lower(), tuple(c.lower() for c in t.colunas))
        if chave in vistos:
            achados.append(Finding(
                "revisar", "TC_TABELA_REPETIDA", f"Tabela repetida: {t.titulo or t.var_linha}",
                f"A tabela da linha {t.linha} repete a da linha {vistos[chave].linha} (mesma variável e mesmas colunas).",
                t.linha, t.var_linha, None, "Apague uma das duas.",
            ))
        else:
            vistos[chave] = t
        if not t.titulo:
            achados.append(Finding(
                "revisar", "TC_SEM_TITULO", f"Tabela sem título: {t.var_linha}",
                f"A tabela da linha {t.linha} ({t.var_linha}) não tem TITLE.", t.linha, t.var_linha,
                None, "Inclua o título «Titulo: …».",
            ))
        if not t.pergunta:
            achados.append(Finding(
                "revisar", "TC_SEM_PERGUNTA", f"Tabela sem pergunta: {t.var_linha}",
                f"A tabela da linha {t.linha} ({t.var_linha}) não tem CAPTION.", t.linha, t.var_linha,
                None, "Inclua a legenda «Pergunta: …».",
            ))
    return achados


def _fmt(v: Any) -> str:
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    except (TypeError, ValueError):
        return str(v)


# ---------------------------------------------------------------------------
# 4. Comparação com o Relatório Descritivo
# ---------------------------------------------------------------------------

@dataclass
class TabelaDescritivo:
    titulo: str
    grupos: list[str]                 # linha de cima do cabeçalho (repetida por coluna)
    colunas: list[str]                # linha de baixo
    linhas: list[tuple[str, list[float | None]]]
    base: list[float | None]
    pergunta: str = ""


def _num(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def ler_descritivo(conteudo: bytes) -> list[TabelaDescritivo]:
    from core.planilha_leitura import linhas_da_planilha

    linhas = linhas_da_planilha(conteudo)
    out: list[TabelaDescritivo] = []
    i, n = 0, len(linhas)

    def txt(v):
        return "" if v is None else str(v).strip()

    while i < n:
        a = txt(linhas[i][0]) if linhas[i] else ""
        if not a.lower().startswith("titulo") and not a.lower().startswith("título"):
            i += 1
            continue
        titulo = _RE_PREFIXO_TITULO.sub("", a)
        i += 1
        cab: list[list[Any]] = []
        while i < n and not txt(linhas[i][0] if linhas[i] else ""):
            if any(txt(v) for v in linhas[i]):
                cab.append(list(linhas[i]))
            i += 1
            if len(cab) > 3:
                break
        largura = 2 + max((max((j for j, v in enumerate(r[1:], 0) if txt(v)), default=-1) for r in cab), default=-1) \
            if cab else 0
        grupos, colunas = [], []
        if cab:
            topo, baixo = (cab[0], cab[1]) if len(cab) >= 2 else (cab[0], cab[0])
            g_atual = ""
            for j in range(1, largura):
                t_ = txt(topo[j]) if j < len(topo) else ""
                if t_:
                    g_atual = t_
                grupos.append(g_atual)
                colunas.append(txt(baixo[j]) if j < len(baixo) else "")
            if len(cab) == 1:
                colunas = list(grupos)
            colunas = [c_ or g_ for g_, c_ in zip(grupos, colunas)]
        dados: list[tuple[str, list[float | None]]] = []
        base: list[float | None] = []
        while i < n:
            lin = linhas[i]
            r = txt(lin[0]) if lin else ""
            if r.lower().startswith("pergunta"):
                break
            if r.lower().startswith("titulo") or r.lower().startswith("título"):
                i -= 1
                break
            vals = [_num(lin[j]) if j < len(lin) else None for j in range(1, 1 + len(colunas))]
            if r.lower() == "base":
                base = vals
            elif r:
                dados.append((r, vals))
            i += 1
        pergunta = ""
        if i < n and txt(linhas[i][0] if linhas[i] else "").lower().startswith("pergunta"):
            pergunta = _RE_PREFIXO_PERGUNTA.sub("", txt(linhas[i][0]))
            i += 1
        out.append(TabelaDescritivo(titulo, grupos, colunas, dados, base, pergunta))
    return out


@dataclass
class ComparacaoTabela:
    ordem: int
    titulo: str
    status: str                      # igual | diferente | so_no_app | so_no_descritivo
    celulas: int = 0
    celulas_diferentes: int = 0
    maior_diferenca: float = 0.0
    detalhe: str = ""
    exemplos: list[str] = field(default_factory=list)


def _chave_coluna(grupo: str, rotulo: str) -> tuple[str, str]:
    g, r = canon(grupo), canon(rotulo)
    if g == "total" and r == "total":
        return ("total", "total")
    if g == r:                       # variável sem rótulo: o grupo é a própria categoria
        return ("*", r)
    return (g, r)


def _acha_coluna(chave: tuple[str, str], por_chave: dict[tuple[str, str], int]) -> int | None:
    if chave in por_chave:
        return por_chave[chave]
    if chave[0] == "*":              # grupo livre: vale qualquer grupo com a mesma categoria
        for (g, r), j in por_chave.items():
            if r == chave[1]:
                return j
    else:                            # o Descritivo pode ter «Total» como grupo de colunas soltas
        for (g, r), j in por_chave.items():
            if g == "*" and r == chave[1]:
                return j
    return None


def comparar_com_descritivo(
    res: ResultadoTotalCompleto,
    descritivo: list[TabelaDescritivo],
    tolerancia: float = TOLERANCIA_COMPARACAO,
) -> list[ComparacaoTabela]:
    """Compara cada tabela do app com a do Descritivo de mesmo título que tenha as
    mesmas colunas (o Descritivo junta, numa só tabela, as várias tabelas do
    Completo; por isso casamos pelas colunas, não pela posição)."""
    out: list[ComparacaoTabela] = []
    usadas: set[int] = set()
    por_titulo: dict[str, list[int]] = {}
    for k, d in enumerate(descritivo):
        por_titulo.setdefault(canon(d.titulo), []).append(k)
    cobertas: dict[int, set[tuple[str, str]]] = {}

    for t, c in zip(res.tabelas, res.calculadas):
        if c is None:
            out.append(ComparacaoTabela(t.ordem, t.titulo or t.var_linha, "so_no_descritivo",
                                        detalhe="tabela não calculada pelo app"))
            continue
        chaves_app = [_chave_coluna(x.grupo, x.rotulo) for x in c.colunas]
        escolhido = None
        for k in por_titulo.get(canon(t.titulo), []):
            d = descritivo[k]
            por_chave_d = {_chave_coluna(g, col): j for j, (g, col) in enumerate(zip(d.grupos, d.colunas))}
            if all(_acha_coluna(ch, por_chave_d) is not None for ch in chaves_app) and (
                len(chaves_app) > 1 or len(d.colunas) == 1 or k not in usadas
            ):
                # tabela só com «Total» casa com a do Descritivo que também é só «Total»
                if len(chaves_app) == 1 and len(d.colunas) != 1 and any(
                    len(descritivo[j].colunas) == 1 for j in por_titulo.get(canon(t.titulo), [])
                ):
                    continue
                if escolhido is None or (len(chaves_app) == 1 and k not in usadas):
                    escolhido = k
                    if k not in usadas:
                        break
        if escolhido is None:
            out.append(ComparacaoTabela(t.ordem, t.titulo or t.var_linha, "so_no_app",
                                        detalhe="não achei essa tabela (com essas colunas) no Descritivo"))
            continue
        usadas.add(escolhido)
        pd_ = {_chave_coluna(g, col): j for j, (g, col) in enumerate(zip(descritivo[escolhido].grupos, descritivo[escolhido].colunas))}
        cobertas.setdefault(escolhido, set()).update(_acha_coluna(ch, pd_) for ch in chaves_app)
        out.append(_comparar_uma(t.ordem, c, descritivo[escolhido], tolerancia))

    for k, d in enumerate(descritivo):
        indices_d = set(range(len(d.colunas)))
        if k not in usadas:
            out.append(ComparacaoTabela(0, d.titulo, "so_no_descritivo", detalhe="tabela do Descritivo sem correspondente na programação"))
        elif indices_d - cobertas.get(k, set()):
            falt = indices_d - cobertas.get(k, set())
            out.append(ComparacaoTabela(0, d.titulo, "so_no_descritivo",
                                        detalhe=f"{len(falt)} coluna(s) do Descritivo não vêm de nenhuma tabela da programação"))
    return out


def _comparar_uma(ordem: int, c: TabelaCalc, d: TabelaDescritivo, tol: float) -> ComparacaoTabela:
    r = ComparacaoTabela(ordem, c.tabela.titulo or c.tabela.var_linha, "igual")
    pos_d = {_chave_coluna(g, col): j for j, (g, col) in enumerate(zip(d.grupos, d.colunas))}
    mapa_cols = [_acha_coluna(_chave_coluna(x.grupo, x.rotulo), pos_d) for x in c.colunas]
    linhas_d = [(rot, vals) for rot, vals in d.linhas if canon(rot) not in LINHAS_SO_DO_DESCRITIVO]
    por_rotulo = {canon(rot): vals for rot, vals in linhas_d}
    app_rot = {canon(x) for x in c.rotulos}
    if len(c.rotulos) != len(linhas_d):
        r.status = "diferente"
        so_app = [x for x in c.rotulos if canon(x) not in {canon(a) for a, _ in linhas_d}]
        so_d = [a for a, _ in linhas_d if canon(a) not in app_rot]
        r.detalhe = f"o app tem {len(c.rotulos)} linha(s) e o Descritivo {len(linhas_d)}"
        if so_app:
            r.detalhe += f"; só no app: {', '.join(so_app[:3])}"
        if so_d:
            r.detalhe += f"; só no Descritivo: {', '.join(so_d[:3])}"
    for i, rot in enumerate(c.rotulos):
        vals = por_rotulo.get(canon(rot))
        if vals is None:
            continue
        for j, jd in enumerate(mapa_cols):
            a = c.valores[i, j]
            b = vals[jd] if jd < len(vals) else None
            r.celulas += 1
            a_v = 0.0 if (a is None or np.isnan(a)) else float(a)
            b_v = 0.0 if b is None else float(b)
            dif = abs(a_v - b_v)
            if dif > tol:
                r.celulas_diferentes += 1
                r.maior_diferenca = max(r.maior_diferenca, dif)
                if len(r.exemplos) < 4:
                    r.exemplos.append(f"{rot} · {c.colunas[j].grupo}/{c.colunas[j].rotulo}: app={a_v:.2f} · Descritivo={b_v:.2f}")
    if d.base:
        for j, jd in enumerate(mapa_cols):
            b = d.base[jd] if jd < len(d.base) else None
            if b is None:
                continue
            col = c.colunas[j]
            r.celulas += 1
            if abs(col.base - b) > 0.5:
                r.celulas_diferentes += 1
                r.maior_diferenca = max(r.maior_diferenca, abs(col.base - b))
                if len(r.exemplos) < 4:
                    r.exemplos.append(f"Base · {col.grupo}/{col.rotulo}: app={col.base:.0f} · Descritivo={b:.0f}")
    if r.celulas_diferentes:
        r.status = "diferente"
    return r


def achados_da_comparacao_tc(comp: list[ComparacaoTabela]) -> list[Finding]:
    achados: list[Finding] = []
    for c in comp:
        if c.status == "diferente":
            achados.append(Finding(
                "critico", "TC_DIFERE_DO_DESCRITIVO",
                f"{c.titulo}: {c.celulas_diferentes} célula(s) diferente(s) do Descritivo",
                (c.detalhe + " " if c.detalhe else "") + "; ".join(c.exemplos),
                None, None, None, "Veja se o banco e o peso são os mesmos usados no SPSS.",
            ))
        elif c.status == "so_no_app":
            achados.append(Finding(
                "revisar", "TC_SO_NO_APP", f"{c.titulo}: não está no Descritivo", c.detalhe,
                None, None, None, None,
            ))
    return achados
