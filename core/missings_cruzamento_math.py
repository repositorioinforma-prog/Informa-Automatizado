"""
Missings das variáveis de cruzamento (segmentos) do Completo — lógica pura, sem Streamlit.

No Total e Completo, cada variável de segmento (P1, P2, Q13S, o conjunto $P6…) vira um conjunto de colunas,
uma por categoria. Categoria com poucos casos dá coluna de base pequena e pouco confiável; o procedimento da
equipe é declará-la como missing para ela sair das colunas (os casos continuam no «Total»).

Este módulo:
- acha todas as variáveis de segmento da programação (colunas de TABLES/CTABLES, exceto «total»);
- lista cada categoria com o código, o rótulo e a quantidade de casos (sem peso, só casos com peso válido —
  o mesmo critério da linha «Base» das tabelas);
- sugere como missing as categorias com menos de `BASE_MINIMA` casos, EXCETO as variáveis de `EXCECOES`
  (P1, P2, P3, P7), que nunca entram na sugestão;
- marca o que já é missing no .sav (esses códigos já chegam vazios ao cálculo e entram na sintaxe);
- escreve a sintaxe `MISSING VALUES` equivalente, para colar no SPSS.

Quem decide é a pessoa: a sugestão é só o ponto de partida do painel.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Sequence

import numpy as np
import pandas as pd

from core.analise_perguntas_math import mapa_rotulos, rotulo_categoria
from core.total_completo_math import SEGMENTO_TOTAL, TabelaSps, _col, _pesos, _rotulo_var

BASE_MINIMA = 25
EXCECOES = ("P1", "P2", "P3", "P7")        # variáveis que nunca entram na sugestão automática
SECOES_PADRAO = ("Completo",)             # o painel olha os segmentos do Completo; o Total pode ser incluído
MAX_DISCRETOS_SPSS = 3                     # o SPSS aceita até 3 valores discretos de missing por variável


@dataclass
class CategoriaSegmento:
    var: str                 # nome na programação (com «$» se for conjunto múltiplo)
    rotulo_var: str
    codigo: Any
    rotulo: str
    n: int                   # casos sem peso (com peso válido)
    ja_missing: bool = False  # já é missing definido no .sav
    sugerido: bool = False    # a regra automática marcaria como missing

    limite: int = BASE_MINIMA

    @property
    def situacao(self) -> str:
        if self.ja_missing:
            return "já é missing no banco"
        if self.sugerido:
            return f"base < {self.limite}"
        if self.n == 0:
            return "sem casos"
        return ""


# ---------------------------------------------------------------------------
# Missings que já vêm no .sav
# ---------------------------------------------------------------------------

def _faixas_numericas(specs: Iterable[Any]) -> list[tuple[float, float]]:
    out: list[tuple[float, float]] = []
    for sp in specs or []:
        try:
            lo, hi = float(sp["lo"]), float(sp["hi"])
        except (KeyError, TypeError, ValueError):
            continue                       # missing de variável texto («-» etc.): fora do escopo
        out.append((lo, hi))
    return out


def aplicar_missings_do_banco(df: pd.DataFrame, missing_ranges: dict[str, list[Any]] | None) -> pd.DataFrame:
    """Devolve uma cópia com os códigos de missing de usuário do .sav trocados por NaN
    (o que o SPSS faz ao calcular tabelas e o que o `read_sav` padrão já faz)."""
    out = df.copy()
    por_nome = {str(c).upper(): c for c in out.columns}
    for var, specs in (missing_ranges or {}).items():
        c = por_nome.get(str(var).upper())
        if c is None:
            continue
        if pd.api.types.is_numeric_dtype(out[c]):
            for lo, hi in _faixas_numericas(specs):
                out[c] = out[c].mask((out[c] >= lo) & (out[c] <= hi))
        else:                                   # variável texto: missing é um texto exato (ex.: «-»)
            for sp in specs or []:
                try:
                    lo, hi = sp["lo"], sp["hi"]
                except (KeyError, TypeError):
                    continue
                if isinstance(lo, str) and lo == hi:
                    out[c] = out[c].mask(out[c] == lo)
    return out


def _em_faixa(codigo: Any, faixas: Sequence[tuple[float, float]]) -> bool:
    try:
        x = float(codigo)
    except (TypeError, ValueError):
        return False
    return any(lo <= x <= hi for lo, hi in faixas)


# ---------------------------------------------------------------------------
# Listar variáveis de segmento
# ---------------------------------------------------------------------------

def variaveis_de_segmento(tabelas: Iterable[TabelaSps], secoes: Iterable[str] = SECOES_PADRAO) -> list[str]:
    """Nomes das variáveis de segmento das seções pedidas (por padrão só o Completo), na ordem em que
    aparecem, sem «total»."""
    pedidas = {x.lower() for x in secoes}
    vistos: dict[str, str] = {}
    for t in tabelas:
        if not t.suportada or t.secao.lower() not in pedidas:
            continue
        for seg in t.colunas:
            if seg.lower() == SEGMENTO_TOTAL:
                continue
            vistos.setdefault(seg.upper(), seg)
    return list(vistos.values())


def _chave_var(nome: str) -> str:
    return nome.upper().lstrip("$")


def listar_categorias(
    tabelas: Iterable[TabelaSps],
    df_bruto: pd.DataFrame,
    value_labels: dict[str, dict[Any, str]],
    var_labels: dict[str, str],
    mrsets: dict[str, list[str]] | None = None,
    peso: str | None = None,
    missing_ranges: dict[str, list[Any]] | None = None,
    base_minima: int = BASE_MINIMA,
    excecoes: Iterable[str] = EXCECOES,
    secoes: Iterable[str] = SECOES_PADRAO,
) -> list[CategoriaSegmento]:
    """Todas as categorias de todas as variáveis de segmento. `df_bruto` é o banco COM os códigos de missing
    do .sav (para mostrar quantos casos eles tinham); o cálculo usa a versão já sem eles."""
    vl = {str(k).upper(): v for k, v in (value_labels or {}).items()}
    lb = {str(k).upper(): v for k, v in (var_labels or {}).items()}
    mr = {str(k).upper(): v for k, v in (mrsets or {}).items()}
    miss = {str(k).upper(): _faixas_numericas(v) for k, v in (missing_ranges or {}).items()}
    sem_regra = {_chave_var(x) for x in excecoes}
    w = _pesos(df_bruto, peso)
    ok = w.notna().to_numpy()

    out: list[CategoriaSegmento] = []
    for seg in variaveis_de_segmento(tabelas, secoes):
        if seg.startswith("$"):
            membros = mr.get(seg.upper())
            if not membros:
                continue
            cols = [c for c in (_col(df_bruto, m) for m in membros) if c is not None]
            if not cols:
                continue
            mat = df_bruto[cols].apply(pd.to_numeric, errors="coerce")
            rot: dict[Any, str] = {}
            for m in membros:
                for k, v in mapa_rotulos(vl.get(m.upper())).items():
                    rot.setdefault(k, v)
            faixas = [f for m in membros for f in miss.get(m.upper(), [])]
            codigos = {float(v) for v in pd.unique(mat.to_numpy().ravel()) if pd.notna(v)} | {k for k in rot if isinstance(k, float)}
            contar = lambda k: int(((mat == k).any(axis=1).to_numpy() & ok).sum())      # noqa: E731
        else:
            c = _col(df_bruto, seg)
            if c is None:
                continue
            serie = df_bruto[c]
            rot = mapa_rotulos(vl.get(seg.upper()))
            faixas = miss.get(seg.upper(), [])
            codigos = {v for v in serie.dropna().unique().tolist()} | {k for k in rot if isinstance(k, (int, float))}
            contar = lambda k: int(((serie == k).to_numpy() & ok).sum())                 # noqa: E731
        try:
            ordenados = sorted(codigos)
        except TypeError:
            ordenados = sorted(codigos, key=str)
        rotulo_v = _rotulo_var(seg.lstrip("$") if not seg.startswith("$") else seg, lb)
        for k in ordenados:
            n = contar(k)
            ja = _em_faixa(k, faixas)
            if n == 0 and not ja and k not in rot:
                continue                                  # código que nem existe no banco nem tem rótulo
            out.append(CategoriaSegmento(
                var=seg, rotulo_var=rotulo_v, codigo=k, rotulo=rotulo_categoria(k, rot), n=n,
                ja_missing=ja,
                sugerido=(not ja) and 0 < n < base_minima and _chave_var(seg) not in sem_regra,
                limite=base_minima,
            ))
    return out


# ---------------------------------------------------------------------------
# Escolhas → cálculo e sintaxe
# ---------------------------------------------------------------------------

def sugestao(categorias: Iterable[CategoriaSegmento]) -> set[tuple[str, Any]]:
    """Pares (variável, código) marcados como missing por padrão: o que já é missing + a regra automática."""
    return {(c.var.upper(), c.codigo) for c in categorias if c.ja_missing or c.sugerido}


def missings_para_calculo(categorias: Iterable[CategoriaSegmento], marcados: Iterable[tuple[str, Any]]) -> dict[str, set]:
    """{variável: códigos} que o cálculo deve tirar das colunas. O que já é missing no .sav não entra aqui
    (já chega vazio ao cálculo); só o que a pessoa marcou a mais."""
    marc = {(v.upper(), k) for v, k in marcados}
    out: dict[str, set] = {}
    for c in categorias:
        if c.ja_missing:
            continue
        if (c.var.upper(), c.codigo) in marc:
            out.setdefault(c.var.upper(), set()).add(c.codigo)
    return out


def _num(v: Any) -> str:
    f = float(v)
    return str(int(f)) if f == int(f) else str(f)


def sintaxe_missing(
    categorias: Iterable[CategoriaSegmento],
    marcados: Iterable[tuple[str, Any]],
    mrsets: dict[str, list[str]] | None = None,
    missing_ranges: dict[str, list[Any]] | None = None,
) -> str:
    """Comandos `MISSING VALUES` equivalentes ao que está marcado. Inclui o que já é missing no .sav da mesma
    variável, porque o comando SUBSTITUI a definição anterior. Variáveis com o mesmo conjunto de códigos saem
    juntas. Conjunto múltiplo ($P6): o comando vai nas variáveis-membro."""
    marc = {(v.upper(), k) for v, k in marcados}
    mr = {str(k).upper(): v for k, v in (mrsets or {}).items()}
    existentes = {str(k).upper(): _faixas_numericas(v) for k, v in (missing_ranges or {}).items()}

    novos: dict[str, set] = {}
    for c in categorias:
        if (c.var.upper(), c.codigo) in marc and not c.ja_missing:
            novos.setdefault(c.var, set()).add(c.codigo)
    por_var: dict[str, list[str]] = {}            # variável real -> definição (lista de textos)
    avisos: list[str] = []
    for seg, codigos in novos.items():
        alvos = mr.get(seg.upper(), []) if seg.startswith("$") else [seg]
        for a in alvos:
            ex = existentes.get(a.upper(), [])
            itens = [f"{_num(lo)} THRU {_num(hi)}" if lo != hi else _num(lo) for lo, hi in ex]
            itens += [_num(k) for k in sorted(codigos) if not _em_faixa(k, ex)]
            por_var[a] = itens
            discretos = sum(1 for i in itens if "THRU" not in i)
            if discretos > MAX_DISCRETOS_SPSS:
                avisos.append(f"* ATENÇÃO: {a} ficou com {discretos} valores discretos; o SPSS aceita no máximo {MAX_DISCRETOS_SPSS}. "
                              "Recodifique as categorias pequenas para um único código e declare só esse como missing.")
    grupos: dict[tuple[str, ...], list[str]] = {}
    for v, itens in por_var.items():
        grupos.setdefault(tuple(itens), []).append(v)
    linhas = [f"MISSING VALUES {' '.join(vs)} ({', '.join(itens)})." for itens, vs in grupos.items()]
    return "\n".join(linhas + avisos)


def tabela_para_painel(categorias: Sequence[CategoriaSegmento], marcados: Iterable[tuple[str, Any]]) -> pd.DataFrame:
    """DataFrame do painel (uma linha por categoria)."""
    marc = {(v.upper(), k) for v, k in marcados}
    return pd.DataFrame([{
        "Variável": c.var, "Rótulo da variável": c.rotulo_var, "Código": c.codigo, "Rótulo": c.rotulo,
        "Casos": c.n, "Situação": c.situacao,
        "Missing": bool(c.ja_missing or (c.var.upper(), c.codigo) in marc),
    } for c in categorias])


def marcados_do_painel(df: pd.DataFrame) -> set[tuple[str, Any]]:
    """Lê de volta o que está marcado no painel (já-missing sempre conta como marcado)."""
    out: set[tuple[str, Any]] = set()
    for _, r in df.iterrows():
        if bool(r["Missing"]) or r["Situação"] == "já é missing no banco":
            out.add((str(r["Variável"]).upper(), r["Código"]))
    return out
