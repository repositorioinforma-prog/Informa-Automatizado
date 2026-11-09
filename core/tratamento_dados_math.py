"""
Tratamento de Dados — manipulação de bancos .sav "como no SPSS".

Núcleo puro (sem Streamlit), testável isoladamente. Cada operação do SPSS
(COMPUTE, IF, RECODE, SELECT IF, SORT CASES, MISSING VALUES, VALUE LABELS,
ADD FILES, MATCH FILES, AGGREGATE...) vira uma função que recebe um `Banco`
e devolve um `Resultado` com:
  - o novo banco (o original nunca é alterado);
  - a syntax SPSS equivalente (para rodar no SPSS ou guardar como registro);
  - uma descrição curta e avisos para a pessoa.

Um `Passo` guarda (operação, parâmetros); o histórico é a lista de passos e o
"desfazer" é reaplicar os passos restantes sobre o banco original.

Semântica fiel ao SPSS onde importa:
  - missing de usuário viram "ausente" nas expressões (MISSING() e VALUE()
    enxergam o código bruto);
  - comparação com ausente não é verdadeira nem falsa (lógica de 3 valores);
  - IF/SELECT IF só agem onde a condição é VERDADEIRA;
  - RECODE: vale a primeira regra que casar; em variável nova, o que não foi
    citado vira system-missing; em variável existente, fica como estava.

As expressões são lidas por um analisador próprio (sem `eval`).
"""
from __future__ import annotations

import io
import math
import os
import re
import sys
import tempfile
import warnings
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

LOWEST = float(np.nextafter(-sys.float_info.max, 0))   # LO / LOWEST do SPSS
HIGHEST = sys.float_info.max                           # HI / HIGHEST do SPSS
MEDIDAS = ("nominal", "ordinal", "scale")
RESERVADAS = {"ALL", "AND", "BY", "EQ", "GE", "GT", "LE", "LT", "NE", "NOT", "OR", "TO", "WITH"}


class ErroTratamento(ValueError):
    """Erro de uso (expressão inválida, variável inexistente...) com mensagem para a pessoa."""


# =============================================================================
# Estado do banco
# =============================================================================
@dataclass
class Banco:
    df: pd.DataFrame
    rotulos_var: Dict[str, str] = field(default_factory=dict)
    rotulos_valor: Dict[str, Dict[Any, str]] = field(default_factory=dict)
    missing: Dict[str, List[Dict[str, Any]]] = field(default_factory=dict)
    medida: Dict[str, str] = field(default_factory=dict)
    nome: str = "banco.sav"

    def copia(self) -> "Banco":
        return Banco(
            df=self.df.copy(),
            rotulos_var=dict(self.rotulos_var),
            rotulos_valor={k: dict(v) for k, v in self.rotulos_valor.items()},
            missing={k: [dict(f) for f in v] for k, v in self.missing.items()},
            medida=dict(self.medida),
            nome=self.nome,
        )

    @property
    def colunas(self) -> List[str]:
        return [str(c) for c in self.df.columns]


@dataclass
class Resultado:
    banco: Banco
    sintaxe: str
    descricao: str
    avisos: List[str] = field(default_factory=list)


@dataclass
class Passo:
    operacao: str
    params: Dict[str, Any]
    descricao: str = ""
    sintaxe: str = ""
    avisos: List[str] = field(default_factory=list)


def eh_texto(serie: pd.Series) -> bool:
    return not (pd.api.types.is_numeric_dtype(serie) or pd.api.types.is_datetime64_any_dtype(serie))


# ----------------------------------------------------------------------------- utilidades
def _col(banco: Banco, nome: str) -> str:
    """Resolve o nome de variável sem diferenciar maiúsculas de minúsculas (como o SPSS)."""
    nome = str(nome).strip()
    for c in banco.df.columns:
        if str(c).lower() == nome.lower():
            return str(c)
    raise ErroTratamento(f"A variável '{nome}' não existe no banco.")


def _col_ou_none(banco: Banco, nome: str) -> Optional[str]:
    try:
        return _col(banco, nome)
    except ErroTratamento:
        return None


def validar_nome_variavel(nome: str) -> str:
    nome = str(nome).strip()
    if not nome:
        raise ErroTratamento("Informe o nome da variável.")
    if len(nome.encode("utf-8")) > 64:
        raise ErroTratamento("O nome da variável pode ter no máximo 64 bytes.")
    if not (re.match(r"^[^\W\d_][\w.@#$]*$", nome) or re.match(r"^[@#$][\w.@#$]*$", nome)):
        raise ErroTratamento(
            f"Nome inválido: '{nome}'. Use letra (ou @, #, $) no início e só letras, números, _ . @ # $ depois.")
    if nome.endswith("."):
        raise ErroTratamento("O nome da variável não pode terminar com ponto.")
    if nome.upper() in RESERVADAS:
        raise ErroTratamento(f"'{nome}' é palavra reservada do SPSS.")
    return nome


def _lista_vars(banco: Banco, nomes: Sequence[str]) -> List[str]:
    if not nomes:
        raise ErroTratamento("Escolha ao menos uma variável.")
    vistos, saida = set(), []
    for n in nomes:
        c = _col(banco, n)
        if c.lower() not in vistos:
            vistos.add(c.lower())
            saida.append(c)
    return saida


def _num(x: float) -> str:
    if x <= LOWEST:
        return "LO"
    if x >= HIGHEST:
        return "HI"
    if float(x).is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(float(x))


def _txt(s: str) -> str:
    return "'" + str(s).replace("'", "''") + "'"


def _limpa_expr(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "").strip()).rstrip(".").strip()


def _lista_sintaxe(nomes: Sequence[str]) -> str:
    return " ".join(nomes)


def _mascara_missing(arr: np.ndarray, faixas: Sequence[Dict[str, Any]], texto: bool) -> np.ndarray:
    if texto:
        vals = {str(f["lo"]).rstrip() for f in faixas if f.get("lo") == f.get("hi")}
        return np.array([x is not None and str(x).rstrip() in vals for x in arr], dtype=bool)
    m = np.zeros(len(arr), dtype=bool)
    with np.errstate(invalid="ignore"):
        for f in faixas:
            m |= (arr >= float(f["lo"])) & (arr <= float(f["hi"]))
    return m


def valores_numericos(banco: Banco, nome: str, bruto: bool = False) -> np.ndarray:
    c = _col(banco, nome)
    arr = pd.to_numeric(banco.df[c], errors="coerce").to_numpy(dtype=float, na_value=np.nan)
    if not bruto and banco.missing.get(c):
        arr = arr.copy()
        arr[_mascara_missing(arr, banco.missing[c], False)] = np.nan
    return arr


def valores_texto(banco: Banco, nome: str, bruto: bool = False) -> np.ndarray:
    c = _col(banco, nome)
    s = banco.df[c]
    arr = np.array([None if (x is None or (isinstance(x, float) and math.isnan(x)) or x is pd.NA) else str(x)
                    for x in s.tolist()], dtype=object)
    if not bruto and banco.missing.get(c):
        arr[_mascara_missing(arr, banco.missing[c], True)] = None
    return arr


# =============================================================================
# Avaliador de expressões (sintaxe de COMPUTE / IF do SPSS)
# =============================================================================
_TOKEN = re.compile(r"""\s*(?:
    (?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)
   |(?P<str>'(?:[^']|'')*'|"(?:[^"]|"")*")
   |(?P<name>[^\W\d][\w.@#$]*|[@#$][\w.@#$]*)
   |(?P<op>\*\*|<>|~=|!=|<=|>=|[-+*/^=<>&|~(),])
)""", re.X)

_OPS_PALAVRA = {"AND", "OR", "NOT", "EQ", "NE", "LT", "GT", "LE", "GE", "TO"}
_CMP = {"=": "eq", "EQ": "eq", "<>": "ne", "~=": "ne", "!=": "ne", "NE": "ne",
        "<": "lt", "LT": "lt", ">": "gt", "GT": "gt", "<=": "le", "LE": "le", ">=": "ge", "GE": "ge"}


@dataclass
class _V:
    k: str            # "n" (numérico/lógico: 1, 0 ou NaN) | "s" (texto)
    a: np.ndarray


def _tokenizar(texto: str) -> List[Tuple[str, str]]:
    texto = _limpa_expr(texto)
    if not texto:
        raise ErroTratamento("A expressão está vazia.")
    pos, saida = 0, []
    while pos < len(texto):
        if texto[pos:].strip() == "":
            break
        m = _TOKEN.match(texto, pos)
        if not m or m.end() == pos:
            raise ErroTratamento(f"Caractere inesperado em: ...{texto[pos:pos + 15]!r}")
        pos = m.end()
        for tipo in ("num", "str", "name", "op"):
            if m.group(tipo) is not None:
                saida.append((tipo, m.group(tipo)))
                break
    return saida


class _Avaliador:
    def __init__(self, banco: Banco, texto: str):
        self.b = banco
        self.n = len(banco.df)
        self.tok = _tokenizar(texto)
        self.i = 0
        self.texto = texto

    # ----- tokens
    def _peek(self, k: int = 0) -> Optional[Tuple[str, str]]:
        j = self.i + k
        return self.tok[j] if j < len(self.tok) else None

    def _eh(self, valor: str, k: int = 0) -> bool:
        t = self._peek(k)
        return bool(t) and t[1].upper() == valor.upper() and t[0] in ("op", "name")

    def _aceita(self, valor: str) -> bool:
        if self._eh(valor):
            self.i += 1
            return True
        return False

    def _espera(self, valor: str) -> None:
        if not self._aceita(valor):
            raise ErroTratamento(f"Esperava '{valor}' na expressão: {self.texto}")

    def avaliar(self) -> _V:
        v = self._or()
        if self.i < len(self.tok):
            raise ErroTratamento(f"Sobrou texto na expressão a partir de '{self.tok[self.i][1]}'.")
        return v

    # ----- lógica de 3 valores
    @staticmethod
    def _tri(v: _V) -> np.ndarray:
        if v.k != "n":
            raise ErroTratamento("Operador lógico aplicado a texto.")
        return np.where(np.isnan(v.a), np.nan, (v.a != 0).astype(float))

    def _or(self) -> _V:
        v = self._and()
        while self._eh("OR") or self._eh("|"):
            self.i += 1
            w = self._and()
            a, b = self._tri(v), self._tri(w)
            r = np.where((a == 1) | (b == 1), 1.0, np.where(np.isnan(a) | np.isnan(b), np.nan, 0.0))
            v = _V("n", r)
        return v

    def _and(self) -> _V:
        v = self._not()
        while self._eh("AND") or self._eh("&"):
            self.i += 1
            w = self._not()
            a, b = self._tri(v), self._tri(w)
            r = np.where((a == 0) | (b == 0), 0.0, np.where(np.isnan(a) | np.isnan(b), np.nan, 1.0))
            v = _V("n", r)
        return v

    def _not(self) -> _V:
        if self._eh("NOT") or self._eh("~"):
            self.i += 1
            return _V("n", 1.0 - self._tri(self._not()))
        return self._comp()

    def _comp(self) -> _V:
        v = self._soma()
        t = self._peek()
        if t and t[1].upper() in _CMP and t[0] in ("op", "name"):
            self.i += 1
            w = self._soma()
            return self._comparar(_CMP[t[1].upper()], v, w)
        return v

    def _comparar(self, op: str, v: _V, w: _V) -> _V:
        if v.k != w.k:
            raise ErroTratamento("Não dá para comparar número com texto.")
        if v.k == "n":
            a, b = v.a, w.a
            ausente = np.isnan(a) | np.isnan(b)
            with np.errstate(invalid="ignore"):
                r = {"eq": a == b, "ne": a != b, "lt": a < b, "gt": a > b, "le": a <= b, "ge": a >= b}[op]
        else:
            a = np.array([None if x is None else x.rstrip() for x in v.a], dtype=object)
            b = np.array([None if x is None else x.rstrip() for x in w.a], dtype=object)
            ausente = np.array([x is None or y is None for x, y in zip(a, b)])
            r = np.array([False if (x is None or y is None) else
                          {"eq": x == y, "ne": x != y, "lt": x < y, "gt": x > y, "le": x <= y, "ge": x >= y}[op]
                          for x, y in zip(a, b)])
        return _V("n", np.where(ausente, np.nan, r.astype(float)))

    # ----- aritmética
    def _num2(self, v: _V, w: _V, f) -> _V:
        if v.k != "n" or w.k != "n":
            raise ErroTratamento("Operação aritmética com texto.")
        with np.errstate(all="ignore"):
            r = f(v.a, w.a)
        r = np.where(np.isfinite(r), r, np.nan)
        return _V("n", r)

    def _soma(self) -> _V:
        v = self._termo()
        while True:
            if self._eh("+"):
                self.i += 1
                v = self._num2(v, self._termo(), np.add)
            elif self._eh("-"):
                self.i += 1
                v = self._num2(v, self._termo(), np.subtract)
            else:
                return v

    def _termo(self) -> _V:
        v = self._unario()
        while True:
            if self._eh("*"):
                self.i += 1
                v = self._num2(v, self._unario(), np.multiply)
            elif self._eh("/"):
                self.i += 1
                v = self._num2(v, self._unario(), np.divide)
            else:
                return v

    def _unario(self) -> _V:
        if self._eh("-"):
            self.i += 1
            v = self._unario()
            return self._num2(_V("n", np.zeros(self.n)), v, np.subtract)
        if self._eh("+"):
            self.i += 1
            return self._unario()
        return self._potencia()

    def _potencia(self) -> _V:
        v = self._atomo()
        if self._eh("**") or self._eh("^"):
            self.i += 1
            return self._num2(v, self._unario(), np.power)
        return v

    # ----- átomos
    def _cheio(self, valor: float) -> _V:
        return _V("n", np.full(self.n, valor, dtype=float))

    def _atomo(self) -> _V:
        t = self._peek()
        if t is None:
            raise ErroTratamento("A expressão terminou antes do esperado.")
        tipo, val = t
        if tipo == "num":
            self.i += 1
            return self._cheio(float(val))
        if tipo == "str":
            self.i += 1
            q = val[0]
            return _V("s", np.full(self.n, val[1:-1].replace(q * 2, q), dtype=object))
        if tipo == "op" and val == "(":
            self.i += 1
            v = self._or()
            self._espera(")")
            return v
        if tipo == "name":
            if val.upper() in _OPS_PALAVRA:
                raise ErroTratamento(f"'{val}' não pode ser usado aqui.")
            self.i += 1
            if self._eh("("):
                self.i += 1
                return self._funcao(val, self._args())
            return self._variavel(val)
        raise ErroTratamento(f"Símbolo inesperado: '{val}'.")

    def _variavel(self, nome: str) -> _V:
        up = nome.upper()
        if up == "$SYSMIS":
            return self._cheio(np.nan)
        if up == "$CASENUM":
            return _V("n", np.arange(1, self.n + 1, dtype=float))
        c = _col(self.b, nome)
        if eh_texto(self.b.df[c]):
            return _V("s", valores_texto(self.b, c))
        return _V("n", valores_numericos(self.b, c))

    # ----- argumentos de função (aceita "var" e "a TO b")
    def _args(self) -> List[Tuple[str, Any]]:
        itens: List[Tuple[str, Any]] = []
        if self._aceita(")"):
            return itens
        while True:
            itens.append(self._arg())
            if self._aceita(","):
                continue
            break
        self._espera(")")
        return itens

    def _arg(self) -> Tuple[str, Any]:
        t = self._peek()
        if t and t[0] == "name" and t[1].upper() not in _OPS_PALAVRA and _col_ou_none(self.b, t[1]):
            t1, t2 = self._peek(1), self._peek(2)
            if t1 and t1[0] == "name" and t1[1].upper() == "TO" and t2 and t2[0] == "name":
                a, b = _col(self.b, t[1]), _col(self.b, t2[1])
                cols = self.b.colunas
                ia, ib = cols.index(a), cols.index(b)
                if ia > ib:
                    ia, ib = ib, ia
                self.i += 3
                return ("vars", cols[ia:ib + 1])
            if t1 is None or t1[1] in (",", ")"):
                self.i += 1
                return ("var", _col(self.b, t[1]))
        if (t and t[0] == "name" and re.fullmatch(r"[Ff]\d+(\.\d+)?", t[1]) and not _col_ou_none(self.b, t[1])
                and (self._peek(1) is None or self._peek(1)[1] in (",", ")"))):
            self.i += 1                      # formato do SPSS (F8, F8.2) em NUMBER(texto, F8): é ignorado
            return ("fmt", t[1])
        return ("val", self._or())

    def _valores(self, itens) -> List[_V]:
        saida: List[_V] = []
        for tipo, x in itens:
            if tipo == "fmt":
                continue
            if tipo == "val":
                saida.append(x)
            else:
                for nome in ([x] if tipo == "var" else x):
                    saida.append(self._variavel(nome))
        return saida

    # ----- funções
    def _funcao(self, nome: str, itens) -> _V:
        base, _, sufixo = nome.upper().partition(".")
        minimo = int(sufixo) if sufixo.isdigit() else None
        if base in ("SUM", "MEAN", "MIN", "MAX", "SD", "MEDIAN", "NVALID"):
            return self._agregada(base, self._valores(itens), minimo)
        if base in ("MISSING", "SYSMIS", "VALUE"):
            if len(itens) != 1:
                raise ErroTratamento(f"{base} recebe exatamente 1 argumento.")
            tipo, x = itens[0]
            if tipo == "val":
                v = x
                if base == "VALUE":
                    return v
                ausente = np.isnan(v.a) if v.k == "n" else np.array([e is None for e in v.a])
                return _V("n", ausente.astype(float))
            if tipo != "var":
                raise ErroTratamento(f"{base} recebe uma única variável.")
            c = x
            if eh_texto(self.b.df[c]):
                bruto = valores_texto(self.b, c, bruto=True)
                if base == "VALUE":
                    return _V("s", bruto)
                alvo = valores_texto(self.b, c, bruto=(base == "SYSMIS"))
                return _V("n", np.array([e is None for e in alvo], dtype=float))
            bruto = valores_numericos(self.b, c, bruto=True)
            if base == "VALUE":
                return _V("n", bruto)
            alvo = bruto if base == "SYSMIS" else valores_numericos(self.b, c)
            return _V("n", np.isnan(alvo).astype(float))
        vals = self._valores(itens)
        if base == "ANY":
            return self._any(vals)
        if base == "RANGE":
            return self._range(vals)
        if base in ("ABS", "SQRT", "EXP", "LN", "LG10", "RND", "TRUNC"):
            return self._matematica(base, vals)
        if base == "MOD":
            if len(vals) != 2:
                raise ErroTratamento("MOD recebe 2 argumentos.")
            return self._num2(vals[0], vals[1], np.fmod)
        if base in ("CONCAT", "UPCASE", "LOWER", "LTRIM", "RTRIM", "LENGTH", "SUBSTR", "NUMBER"):
            return self._texto(base, vals)
        raise ErroTratamento(f"Função desconhecida: {nome}")

    def _agregada(self, base: str, vals: List[_V], minimo: Optional[int]) -> _V:
        if not vals:
            raise ErroTratamento(f"{base} precisa de argumentos.")
        if any(v.k != "n" for v in vals):
            raise ErroTratamento(f"{base} só aceita valores numéricos.")
        m = np.vstack([v.a for v in vals])
        ok = ~np.isnan(m)
        cont = ok.sum(axis=0)
        if base == "NVALID":
            return _V("n", cont.astype(float))
        minimo = minimo if minimo is not None else (2 if base == "SD" else 1)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            soma = np.where(ok, m, 0.0).sum(axis=0)
            if base == "SUM":
                r = soma
            elif base == "MEAN":
                r = soma / np.where(cont == 0, np.nan, cont)
            elif base == "MIN":
                r = np.where(ok, m, np.inf).min(axis=0)
            elif base == "MAX":
                r = np.where(ok, m, -np.inf).max(axis=0)
            elif base == "MEDIAN":
                r = np.nanmedian(m, axis=0)
            else:  # SD
                media = soma / np.where(cont == 0, np.nan, cont)
                ss = np.where(ok, (m - media) ** 2, 0.0).sum(axis=0)
                r = np.sqrt(ss / np.where(cont < 2, np.nan, cont - 1))
        r = np.where((cont >= minimo) & np.isfinite(r), r, np.nan)
        return _V("n", r)

    def _any(self, vals: List[_V]) -> _V:
        if len(vals) < 2:
            raise ErroTratamento("ANY(teste, valor1, valor2, ...) precisa de ao menos 2 argumentos.")
        teste, resto = vals[0], vals[1:]
        achou = np.zeros(self.n, dtype=bool)
        for v in resto:
            r = self._comparar("eq", teste, v).a
            achou |= (r == 1)
        ausente = np.isnan(teste.a) if teste.k == "n" else np.array([e is None for e in teste.a])
        return _V("n", np.where(ausente, np.nan, achou.astype(float)))

    def _range(self, vals: List[_V]) -> _V:
        if len(vals) < 3 or len(vals) % 2 == 0:
            raise ErroTratamento("RANGE(teste, min1, max1, min2, max2, ...) precisa de pares min/max.")
        teste, resto = vals[0], vals[1:]
        achou = np.zeros(self.n, dtype=bool)
        for lo, hi in zip(resto[0::2], resto[1::2]):
            achou |= (self._comparar("ge", teste, lo).a == 1) & (self._comparar("le", teste, hi).a == 1)
        ausente = np.isnan(teste.a) if teste.k == "n" else np.array([e is None for e in teste.a])
        return _V("n", np.where(ausente, np.nan, achou.astype(float)))

    def _matematica(self, base: str, vals: List[_V]) -> _V:
        if len(vals) != 1 or vals[0].k != "n":
            raise ErroTratamento(f"{base} recebe 1 argumento numérico.")
        a = vals[0].a
        with np.errstate(all="ignore"):
            f = {"ABS": np.abs, "SQRT": np.sqrt, "EXP": np.exp, "LN": np.log, "LG10": np.log10,
                 "RND": lambda x: np.sign(x) * np.floor(np.abs(x) + 0.5), "TRUNC": np.trunc}[base]
            r = f(a)
        return _V("n", np.where(np.isfinite(r), r, np.nan))

    def _texto(self, base: str, vals: List[_V]) -> _V:
        def so_texto(v: _V) -> np.ndarray:
            if v.k != "s":
                raise ErroTratamento(f"{base} só aceita texto.")
            return v.a

        if base == "CONCAT":
            if not vals:
                raise ErroTratamento("CONCAT precisa de argumentos.")
            cols = [so_texto(v) for v in vals]
            return _V("s", np.array(["".join("" if p is None else p for p in linha) for linha in zip(*cols)],
                                    dtype=object))
        if base in ("UPCASE", "LOWER", "LTRIM", "RTRIM"):
            if len(vals) != 1:
                raise ErroTratamento(f"{base} recebe 1 argumento.")
            f = {"UPCASE": str.upper, "LOWER": str.lower, "LTRIM": str.lstrip, "RTRIM": str.rstrip}[base]
            return _V("s", np.array([None if x is None else f(x) for x in so_texto(vals[0])], dtype=object))
        if base == "LENGTH":
            return _V("n", np.array([np.nan if x is None else float(len(x)) for x in so_texto(vals[0])]))
        if base == "SUBSTR":
            if len(vals) not in (2, 3):
                raise ErroTratamento("SUBSTR(texto, posição[, tamanho]).")
            txt = so_texto(vals[0])
            pos = vals[1].a
            tam = vals[2].a if len(vals) == 3 else np.full(self.n, np.inf)
            saida = []
            for x, p, t in zip(txt, pos, tam):
                if x is None or np.isnan(p) or np.isnan(t):
                    saida.append(None)
                else:
                    ini = max(int(p) - 1, 0)
                    saida.append(x[ini:] if np.isinf(t) else x[ini:ini + int(t)])
            return _V("s", np.array(saida, dtype=object))
        # NUMBER(texto, F8.0): converte texto em número (o formato é ignorado)
        def conv(x):
            try:
                return float(str(x).strip().replace(",", "."))
            except (TypeError, ValueError):
                return np.nan
        return _V("n", np.array([conv(x) if x is not None else np.nan for x in so_texto(vals[0])]))


def avaliar(banco: Banco, expressao: str) -> _V:
    return _Avaliador(banco, expressao).avaliar()


def avaliar_condicao(banco: Banco, condicao: str) -> np.ndarray:
    """Máscara booleana: True só onde a condição é VERDADEIRA (ausente conta como não)."""
    v = avaliar(banco, condicao)
    if v.k != "n":
        raise ErroTratamento("A condição precisa resultar em verdadeiro/falso (use =, >, AND, OR...).")
    return v.a == 1


def _tri_condicao(banco: Banco, condicao: str) -> np.ndarray:
    v = avaliar(banco, condicao)
    if v.k != "n":
        raise ErroTratamento("A condição precisa resultar em verdadeiro/falso (use =, >, AND, OR...).")
    return v.a


# =============================================================================
# COMPUTE / IF
# =============================================================================
def compute(banco: Banco, alvo: str, expressao: str, condicao: Optional[str] = None,
            rotulo: Optional[str] = None) -> Resultado:
    alvo = validar_nome_variavel(alvo)
    expressao = _limpa_expr(expressao)
    condicao = _limpa_expr(condicao) if condicao else None
    b = banco.copia()
    val = avaliar(b, expressao)
    existente = _col_ou_none(b, alvo)
    nome_final = existente or alvo
    n = len(b.df)
    mascara = avaliar_condicao(b, condicao) if condicao else np.ones(n, dtype=bool)
    avisos: List[str] = []
    pre = ""

    if val.k == "n":
        if existente and eh_texto(b.df[existente]):
            raise ErroTratamento(f"'{existente}' é uma variável de texto; o resultado da expressão é numérico.")
        base = valores_numericos(b, existente, bruto=True) if existente else np.full(n, np.nan)
        b.df[nome_final] = np.where(mascara, val.a, base)
        if not existente:
            b.medida[nome_final] = "scale"
    else:
        if existente and not eh_texto(b.df[existente]):
            raise ErroTratamento(f"'{existente}' é numérica; o resultado da expressão é texto.")
        base = valores_texto(b, existente, bruto=True) if existente else np.full(n, None, dtype=object)
        novo = np.where(mascara, val.a, base)
        b.df[nome_final] = pd.Series(novo, index=b.df.index, dtype=object)
        if not existente:
            b.medida[nome_final] = "nominal"
            larg = max([len(x) for x in val.a if x is not None] + [1])
            pre = f"STRING {nome_final} (A{min(max(larg, 1), 255)}).\n"
    if rotulo:
        b.rotulos_var[nome_final] = rotulo.strip()
    if existente:
        avisos.append(f"A variável '{existente}' já existia e foi sobrescrita"
                      f"{' nos casos que atendem à condição' if condicao else ''}.")
    linha = (f"IF ({condicao}) {nome_final} = {expressao}." if condicao
             else f"COMPUTE {nome_final} = {expressao}.")
    sintaxe = pre + linha
    if rotulo:
        sintaxe += f"\nVARIABLE LABELS {nome_final} {_txt(rotulo.strip())}."
    sintaxe += "\nEXECUTE."
    desc = f"{'IF' if condicao else 'COMPUTE'}: {nome_final} = {expressao}" + (f"  [se {condicao}]" if condicao else "")
    if condicao:
        avisos.append(f"A condição foi verdadeira em {int(mascara.sum())} de {n} casos.")
    return Resultado(b, sintaxe, desc, avisos)


# =============================================================================
# RECODE
# =============================================================================
@dataclass
class _Regra:
    entrada: List[Tuple[str, Any]]      # ("val", x) | ("faixa", lo, hi) | ("missing",) | ("sysmis",) | ("else",)
    saida: Tuple[str, Any]              # ("val", x) | ("sysmis",) | ("copy",)
    texto: str


def _valor_literal(tok: str):
    tok = tok.strip()
    if len(tok) >= 2 and tok[0] in "'\"" and tok[-1] == tok[0]:
        return tok[1:-1].replace(tok[0] * 2, tok[0])
    return float(tok)


def _divide_igual(txt: str) -> Tuple[str, str]:
    """Separa 'esquerda = direita' no primeiro '=' fora de aspas."""
    aspas = None
    for i, ch in enumerate(txt):
        if aspas:
            if ch == aspas:
                aspas = None
        elif ch in "'\"":
            aspas = ch
        elif ch == "=":
            return txt[:i], txt[i + 1:]
    raise ErroTratamento(f"Regra sem '=': {txt!r}. Use o formato: 1 THRU 3 = 1")


def analisar_regras_recode(texto: str) -> List[_Regra]:
    """Lê regras no estilo SPSS: '(1 THRU 3=1) (4,5=2) (ELSE=COPY)' ou uma por linha."""
    texto = (texto or "").strip().rstrip(".")
    if not texto:
        raise ErroTratamento("Informe as regras de recodificação.")
    if "(" in texto:
        brutos = re.findall(r"\(((?:[^()']|'[^']*')*)\)", texto)
    else:
        brutos = [p for p in re.split(r"[\n;]+", texto) if p.strip()]
    regras: List[_Regra] = []
    for bruto in brutos:
        esq, dir_ = _divide_igual(bruto)
        itens = re.findall(r"'(?:[^']|'')*'|\"(?:[^\"]|\"\")*\"|[-+]?\d+(?:\.\d+)?|[A-Za-z]+", esq)
        if not itens:
            raise ErroTratamento(f"Regra sem valores de entrada: {bruto!r}")
        entrada, i = [], 0
        while i < len(itens):
            it = itens[i]
            up = it.upper()
            if up in ("ELSE", "MISSING", "SYSMIS"):
                entrada.append((up.lower(),))
                i += 1
                continue
            if (i + 1 < len(itens) and itens[i + 1].upper() == "THRU") or up in ("LO", "LOWEST"):
                lo = LOWEST if up in ("LO", "LOWEST") else float(it)
                if i + 2 >= len(itens) or itens[i + 1].upper() != "THRU":
                    raise ErroTratamento(f"Faixa incompleta em: {bruto!r}")
                fim = itens[i + 2]
                hi = HIGHEST if fim.upper() in ("HI", "HIGHEST") else float(fim)
                entrada.append(("faixa", lo, hi))
                i += 3
                continue
            if up in ("HI", "HIGHEST", "THRU"):
                raise ErroTratamento(f"Uso inválido de {it} em: {bruto!r}")
            entrada.append(("val", _valor_literal(it)))
            i += 1
        d = dir_.strip()
        if d.upper() == "SYSMIS":
            saida = ("sysmis",)
        elif d.upper() == "COPY":
            saida = ("copy",)
        else:
            try:
                saida = ("val", _valor_literal(d))
            except ValueError:
                raise ErroTratamento(f"Valor de saída inválido: {d!r} (use um número, 'texto', SYSMIS ou COPY).")
        if any(e[0] == "else" for e in entrada) and len(entrada) > 1:
            raise ErroTratamento("ELSE deve vir sozinho na regra.")
        regras.append(_Regra(entrada, saida, f"{esq.strip()}={d}"))
    if any(e[0] == "else" for r in regras[:-1] for e in r.entrada):
        raise ErroTratamento("ELSE deve ser a última regra.")
    return regras


def _casa(regra: _Regra, bruto: np.ndarray, sysmis: np.ndarray, usermiss: np.ndarray, texto: bool) -> np.ndarray:
    m = np.zeros(len(bruto), dtype=bool)
    for e in regra.entrada:
        if e[0] == "else":
            m |= True
        elif e[0] == "missing":
            m |= sysmis | usermiss
        elif e[0] == "sysmis":
            m |= sysmis
        elif e[0] == "val":
            if texto:
                alvo = str(e[1]).rstrip() if isinstance(e[1], str) else None
                m |= np.array([x is not None and alvo is not None and x.rstrip() == alvo for x in bruto])
            else:
                if isinstance(e[1], str):
                    raise ErroTratamento("Valor de texto numa variável numérica.")
                with np.errstate(invalid="ignore"):
                    m |= (bruto == e[1])
        else:  # faixa
            if texto:
                raise ErroTratamento("Faixa (THRU) não se aplica a variável de texto.")
            with np.errstate(invalid="ignore"):
                m |= (bruto >= e[1]) & (bruto <= e[2])
    return m


def recode(banco: Banco, origens: Sequence[str], regras: str, destinos: Optional[Sequence[str]] = None,
           rotulo_variavel: Optional[str] = None, rotulos_valor: Optional[Dict[Any, str]] = None) -> Resultado:
    b = banco.copia()
    srcs = _lista_vars(b, origens)
    reg = analisar_regras_recode(regras)
    destinos = [validar_nome_variavel(d) for d in (destinos or [])]
    if destinos and len(destinos) != len(srcs):
        raise ErroTratamento("Informe um nome de destino para cada variável de origem (ou nenhum, para recodificar no lugar).")
    saidas_txt = any(r.saida[0] == "val" and isinstance(r.saida[1], str) for r in reg)
    saidas_num = any(r.saida[0] in ("val", "sysmis") and not (r.saida[0] == "val" and isinstance(r.saida[1], str))
                     for r in reg)
    if saidas_txt and saidas_num:
        raise ErroTratamento("Misturou valores de saída numéricos e de texto na mesma recodificação.")
    avisos: List[str] = []
    for k, src in enumerate(srcs):
        texto_src = eh_texto(b.df[src])
        tgt = destinos[k] if destinos else src
        tgt_exist = _col_ou_none(b, tgt)
        tgt_nome = tgt_exist or tgt
        n = len(b.df)
        if texto_src:
            bruto = valores_texto(b, src, bruto=True)
            sysmis = np.array([x is None for x in bruto])
            usermiss = _mascara_missing(bruto, b.missing.get(src, []), True) if b.missing.get(src) else np.zeros(n, bool)
        else:
            bruto = valores_numericos(b, src, bruto=True)
            sysmis = np.isnan(bruto)
            usermiss = _mascara_missing(bruto, b.missing.get(src, []), False) if b.missing.get(src) else np.zeros(n, bool)
        if saidas_txt:
            saida_texto = True
        elif saidas_num:
            saida_texto = False
        else:                      # só COPY: o tipo do destino segue o da origem
            saida_texto = texto_src
        if not destinos and saida_texto != texto_src:
            raise ErroTratamento("Para trocar o tipo (texto ↔ número), recodifique para uma variável nova (INTO).")
        if tgt_exist and eh_texto(b.df[tgt_exist]) != saida_texto:
            raise ErroTratamento(f"'{tgt_exist}' tem tipo diferente do resultado da recodificação.")
        if not destinos:                       # no lugar: o que não foi citado fica como está
            atual = bruto.copy()
        elif tgt_exist:                        # destino já existe: o que não foi citado fica como estava nele
            atual = (valores_texto(b, tgt_exist, bruto=True) if saida_texto
                     else valores_numericos(b, tgt_exist, bruto=True))
        else:                                  # destino novo: o que não foi citado vira system-missing
            atual = np.full(n, None, dtype=object) if saida_texto else np.full(n, np.nan)
        saida = atual.copy()
        feito = np.zeros(n, dtype=bool)
        for r in reg:
            m = _casa(r, bruto, sysmis, usermiss, texto_src) & ~feito
            if r.saida[0] == "copy":
                if texto_src != saida_texto:
                    raise ErroTratamento("COPY só vale quando origem e destino têm o mesmo tipo.")
                saida[m] = bruto[m]
            elif r.saida[0] == "sysmis":
                if saida_texto:
                    raise ErroTratamento("SYSMIS não se aplica a variável de texto (use '').")
                saida[m] = np.nan
            else:
                saida[m] = r.saida[1]
            feito |= m
        if saida_texto:
            b.df[tgt_nome] = pd.Series(saida, index=b.df.index, dtype=object)
        else:
            b.df[tgt_nome] = saida.astype(float)
        if not tgt_exist:
            b.medida[tgt_nome] = "nominal" if saida_texto else (b.medida.get(src) or "nominal")
        if rotulo_variavel and destinos and len(srcs) == 1:
            b.rotulos_var[tgt_nome] = rotulo_variavel.strip()
        if rotulos_valor and not saida_texto:
            b.rotulos_valor[tgt_nome] = {float(c): str(t) for c, t in rotulos_valor.items()}
        if tgt_exist and destinos:
            avisos.append(f"'{tgt_exist}' já existia e foi sobrescrita (valores não citados nas regras ficam como estavam).")
        if not destinos and b.rotulos_valor.get(src):
            avisos.append(f"Os rótulos de valor de '{src}' continuam os antigos; ajuste-os em Rótulos e Missings.")
    corpo = " ".join(f"({r.texto})" for r in reg)
    destino_txt = f" INTO {_lista_sintaxe(destinos)}" if destinos else ""
    sintaxe = f"RECODE {_lista_sintaxe(srcs)} {corpo}{destino_txt}."
    if destinos and rotulo_variavel and len(destinos) == 1:
        sintaxe += f"\nVARIABLE LABELS {destinos[0]} {_txt(rotulo_variavel.strip())}."
    if destinos and rotulos_valor:
        pares = " ".join(f"{_num(float(c))} {_txt(t)}" for c, t in rotulos_valor.items())
        sintaxe += f"\nVALUE LABELS {_lista_sintaxe(destinos)} {pares}."
    sintaxe += "\nEXECUTE."
    return Resultado(b, sintaxe, f"RECODE {_lista_sintaxe(srcs)}{destino_txt}", avisos)


# =============================================================================
# Casos: SELECT IF, exclusão, SORT CASES, duplicados; variáveis: DELETE, RENAME
# =============================================================================
def selecionar_casos(banco: Banco, condicao: str, modo: str = "manter") -> Resultado:
    if modo not in ("manter", "excluir"):
        raise ErroTratamento("Modo inválido.")
    condicao = _limpa_expr(condicao)
    b = banco.copia()
    tri = _tri_condicao(b, condicao)
    verdadeiro = tri == 1
    ausente = np.isnan(tri)
    manter = verdadeiro if modo == "manter" else (tri == 0)
    avisos: List[str] = []
    if modo == "excluir" and ausente.any():
        avisos.append(f"{int(ausente.sum())} caso(s) em que a condição não pôde ser avaliada (valor ausente) "
                      "também foram removidos — é o mesmo que o SELECT IF NOT do SPSS faz.")
    if modo == "manter" and ausente.any():
        avisos.append(f"{int(ausente.sum())} caso(s) com valor ausente na condição ficaram de fora.")
    antes = len(b.df)
    b.df = b.df.loc[manter].reset_index(drop=True)
    sint = (f"SELECT IF ({condicao})." if modo == "manter" else f"SELECT IF NOT ({condicao}).") + "\nEXECUTE."
    desc = f"SELECT IF{' NOT' if modo == 'excluir' else ''} ({condicao}) — {len(b.df)} de {antes} casos ficaram"
    return Resultado(b, sint, desc, avisos)


def _ordenar_df(df: pd.DataFrame, chaves: List[Tuple[str, bool]]) -> pd.DataFrame:
    out = df
    for nome, asc in reversed(chaves):
        out = out.sort_values(nome, ascending=asc, kind="stable", na_position="first" if asc else "last")
    return out


def ordenar_casos(banco: Banco, chaves: Sequence[Tuple[str, str]]) -> Resultado:
    """chaves: [(variável, "A"|"D"), ...]"""
    if not chaves:
        raise ErroTratamento("Escolha ao menos uma variável para ordenar.")
    b = banco.copia()
    res = [(_col(b, n), str(d).upper().startswith("A")) for n, d in chaves]
    b.df = _ordenar_df(b.df, res).reset_index(drop=True)
    partes = " ".join(f"{n} ({'A' if asc else 'D'})" for n, asc in res)
    return Resultado(b, f"SORT CASES BY {partes}.", f"SORT CASES BY {partes}")


def remover_duplicados(banco: Banco, chaves: Sequence[str], manter: str = "primeiro") -> Resultado:
    if manter not in ("primeiro", "ultimo"):
        raise ErroTratamento("Escolha manter o primeiro ou o último caso de cada grupo.")
    b = banco.copia()
    ks = _lista_vars(b, chaves)
    antes = len(b.df)
    ordenado = _ordenar_df(b.df, [(k, True) for k in ks])
    b.df = ordenado.drop_duplicates(subset=ks, keep="first" if manter == "primeiro" else "last").reset_index(drop=True)
    flag = "PrimeiroCaso" if manter == "primeiro" else "UltimoCaso"
    sub = "FIRST" if manter == "primeiro" else "LAST"
    lst = _lista_sintaxe(ks)
    sint = (f"SORT CASES BY {lst}.\nMATCH FILES /FILE=* /BY {lst} /{sub}={flag}.\n"
            f"SELECT IF ({flag} = 1).\nEXECUTE.\nDELETE VARIABLES {flag}.")
    removidos = antes - len(b.df)
    return Resultado(b, sint, f"Duplicados por {lst}: {removidos} caso(s) removidos (mantido o {manter})",
                     ["Os casos ficam ordenados pelas chaves, como no SPSS."])


def excluir_variaveis(banco: Banco, nomes: Sequence[str]) -> Resultado:
    b = banco.copia()
    ns = _lista_vars(b, nomes)
    if len(ns) >= len(b.df.columns):
        raise ErroTratamento("Não dá para excluir todas as variáveis.")
    b.df = b.df.drop(columns=ns)
    for n in ns:
        b.rotulos_var.pop(n, None)
        b.rotulos_valor.pop(n, None)
        b.missing.pop(n, None)
        b.medida.pop(n, None)
    return Resultado(b, f"DELETE VARIABLES {_lista_sintaxe(ns)}.", f"DELETE VARIABLES {_lista_sintaxe(ns)}")


def renomear_variavel(banco: Banco, atual: str, novo: str) -> Resultado:
    b = banco.copia()
    a = _col(b, atual)
    novo = validar_nome_variavel(novo)
    if novo.lower() != a.lower() and _col_ou_none(b, novo):
        raise ErroTratamento(f"Já existe uma variável chamada '{novo}'.")
    b.df = b.df.rename(columns={a: novo})
    for dic in (b.rotulos_var, b.rotulos_valor, b.missing, b.medida):
        if a in dic:
            dic[novo] = dic.pop(a)
    return Resultado(b, f"RENAME VARIABLES ({a} = {novo}).", f"RENAME VARIABLES ({a} = {novo})")


# =============================================================================
# Rótulos, missings, medida
# =============================================================================
def rotular_variavel(banco: Banco, nome: str, rotulo: str) -> Resultado:
    b = banco.copia()
    c = _col(b, nome)
    rotulo = (rotulo or "").strip()
    if rotulo:
        b.rotulos_var[c] = rotulo
        return Resultado(b, f"VARIABLE LABELS {c} {_txt(rotulo)}.", f"VARIABLE LABELS {c}")
    b.rotulos_var.pop(c, None)
    return Resultado(b, f"VARIABLE LABELS {c} ''.", f"VARIABLE LABELS {c} (removido)")


def analisar_rotulos_valor(texto: str, numerico: bool) -> Dict[Any, str]:
    """Uma linha por rótulo: '1 = Sim', '1 Sim' ou '1<TAB>Sim'."""
    mapa: Dict[Any, str] = {}
    for linha in str(texto or "").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        m = re.match(r"""^('(?:[^']|'')*'|"(?:[^"]|"")*"|[^=\t]+?)\s*(?:=|\t|\s)\s*(.+)$""", linha)
        if not m:
            raise ErroTratamento(f"Linha sem rótulo: {linha!r}. Use: 1 = Sim")
        cod, rot = m.group(1).strip(), m.group(2).strip()
        if len(rot) >= 2 and rot[0] in "'\"" and rot[-1] == rot[0]:
            rot = rot[1:-1]
        if numerico:
            try:
                chave: Any = float(cod.replace(",", "."))
            except ValueError:
                raise ErroTratamento(f"'{cod}' não é um número (a variável é numérica).")
        else:
            chave = cod[1:-1] if len(cod) >= 2 and cod[0] in "'\"" and cod[-1] == cod[0] else cod
        mapa[chave] = rot
    return mapa


def rotular_valores(banco: Banco, nomes: Sequence[str], rotulos: Dict[Any, str], modo: str = "substituir") -> Resultado:
    b = banco.copia()
    ns = _lista_vars(b, nomes)
    if modo not in ("substituir", "adicionar", "remover"):
        raise ErroTratamento("Modo inválido.")
    if len({eh_texto(b.df[n]) for n in ns}) > 1:
        raise ErroTratamento("Escolha só variáveis do mesmo tipo (todas numéricas ou todas de texto).")
    texto = eh_texto(b.df[ns[0]])
    if modo != "remover":
        if not rotulos:
            raise ErroTratamento("Informe ao menos um rótulo de valor.")
        norm = {(str(k) if texto else float(k)): str(v) for k, v in rotulos.items()}
    for n in ns:
        if modo == "remover":
            b.rotulos_valor.pop(n, None)
        elif modo == "substituir":
            b.rotulos_valor[n] = dict(norm)
        else:
            b.rotulos_valor.setdefault(n, {}).update(norm)
    if modo == "remover":
        return Resultado(b, f"VALUE LABELS {_lista_sintaxe(ns)}.", f"Rótulos de valor removidos de {_lista_sintaxe(ns)}")
    pares = " ".join(f"{_txt(k) if texto else _num(k)} {_txt(v)}" for k, v in norm.items())
    cmd = "VALUE LABELS" if modo == "substituir" else "ADD VALUE LABELS"
    return Resultado(b, f"{cmd} {_lista_sintaxe(ns)} {pares}.", f"{cmd} {_lista_sintaxe(ns)}")


def analisar_missing(texto: str, texto_var: bool = False) -> List[Dict[str, Any]]:
    """'99, 98'  |  '98 THRU 99'  |  'LO THRU 0'  |  '98 THRU 99, 0'  |  vazio = nenhum."""
    t = (texto or "").strip().strip("()").strip()
    if not t or t.upper() in ("NENHUM", "NONE"):
        return []
    itens = re.findall(r"'(?:[^']|'')*'|\"(?:[^\"]|\"\")*\"|[-+]?\d+(?:\.\d+)?|[A-Za-z]+", t)
    faixas: List[Dict[str, Any]] = []
    i = 0
    while i < len(itens):
        it = itens[i]
        up = it.upper()
        eh_faixa = (i + 1 < len(itens) and itens[i + 1].upper() == "THRU") or up in ("LO", "LOWEST")
        if eh_faixa:
            if texto_var:
                raise ErroTratamento("Faixas não valem para variável de texto.")
            if i + 2 >= len(itens) or itens[i + 1].upper() != "THRU":
                raise ErroTratamento("Faixa incompleta. Use: 98 THRU 99 ou LO THRU 0.")
            lo = LOWEST if up in ("LO", "LOWEST") else float(it)
            fim = itens[i + 2]
            hi = HIGHEST if fim.upper() in ("HI", "HIGHEST") else float(fim)
            if lo > hi:
                raise ErroTratamento("Na faixa, o primeiro valor deve ser menor ou igual ao segundo.")
            faixas.append({"lo": lo, "hi": hi})
            i += 3
            continue
        try:
            v = _valor_literal(it) if not texto_var or it[0] in "'\"" else it
        except ValueError:
            raise ErroTratamento(f"Valor de missing inválido: {it!r}")
        if texto_var and not isinstance(v, str):
            v = str(it)
        if not texto_var and isinstance(v, str):
            raise ErroTratamento("Variável numérica: informe números.")
        faixas.append({"lo": v, "hi": v})
        i += 1
    disc = [f for f in faixas if f["lo"] == f["hi"]]
    fx = [f for f in faixas if f["lo"] != f["hi"]]
    if len(fx) > 1 or len(disc) > (1 if fx else 3):
        raise ErroTratamento("O SPSS aceita até 3 valores discretos, ou 1 faixa + 1 valor discreto.")
    return faixas


def _missing_texto(faixas: Sequence[Dict[str, Any]]) -> str:
    partes = []
    for f in faixas:
        if f["lo"] == f["hi"]:
            partes.append(_txt(f["lo"]) if isinstance(f["lo"], str) else _num(f["lo"]))
        else:
            partes.append(f"{_num(f['lo'])} THRU {_num(f['hi'])}")
    return ", ".join(partes)


def definir_missing(banco: Banco, nomes: Sequence[str], especificacao: str) -> Resultado:
    b = banco.copia()
    ns = _lista_vars(b, nomes)
    if len({eh_texto(b.df[n]) for n in ns}) > 1:
        raise ErroTratamento("Escolha só variáveis do mesmo tipo (todas numéricas ou todas de texto).")
    faixas = analisar_missing(especificacao, eh_texto(b.df[ns[0]]))
    for n in ns:
        if faixas:
            b.missing[n] = [dict(f) for f in faixas]
        else:
            b.missing.pop(n, None)
    corpo = _missing_texto(faixas)
    return Resultado(b, f"MISSING VALUES {_lista_sintaxe(ns)} ({corpo}).",
                     f"MISSING VALUES {_lista_sintaxe(ns)} ({corpo or 'nenhum'})")


def definir_medida(banco: Banco, nomes: Sequence[str], medida: str) -> Resultado:
    if medida not in MEDIDAS:
        raise ErroTratamento("Medida inválida (use nominal, ordinal ou scale).")
    b = banco.copia()
    ns = _lista_vars(b, nomes)
    for n in ns:
        b.medida[n] = medida
    return Resultado(b, f"VARIABLE LEVEL {_lista_sintaxe(ns)} ({medida.upper()}).",
                     f"VARIABLE LEVEL {_lista_sintaxe(ns)} ({medida.upper()})")


# =============================================================================
# ADD FILES / MATCH FILES / AGGREGATE
# =============================================================================
def _alinha_nomes(b: Banco, outro: Banco) -> Dict[str, str]:
    """nome no arquivo de fora -> nome no arquivo ativo (sem diferenciar maiúsculas)."""
    ativos = {c.lower(): c for c in b.colunas}
    return {c: ativos[c.lower()] for c in outro.colunas if c.lower() in ativos}


def _traz_metadados(b: Banco, outro: Banco, origem: str, destino: str) -> None:
    if origem in outro.rotulos_var:
        b.rotulos_var.setdefault(destino, outro.rotulos_var[origem])
    if origem in outro.rotulos_valor:
        b.rotulos_valor.setdefault(destino, dict(outro.rotulos_valor[origem]))
    if origem in outro.missing:
        b.missing.setdefault(destino, [dict(f) for f in outro.missing[origem]])
    if origem in outro.medida:
        b.medida.setdefault(destino, outro.medida[origem])


def adicionar_casos(banco: Banco, outro: Banco, indicador: Optional[str] = None) -> Resultado:
    b = banco.copia()
    mapa = _alinha_nomes(b, outro)
    if not mapa:
        raise ErroTratamento("Os dois bancos não têm nenhuma variável com o mesmo nome.")
    novo = outro.df.copy()
    avisos: List[str] = []
    for c_out, c_ativo in mapa.items():
        if eh_texto(b.df[c_ativo]) != eh_texto(novo[c_out]):
            raise ErroTratamento(f"A variável '{c_ativo}' é de texto num banco e numérica no outro.")
    novo = novo.rename(columns=mapa)
    extras = [c for c in novo.columns if c not in b.df.columns]
    if extras:
        avisos.append("Variáveis só do segundo banco (ficam ausentes nos casos do primeiro): " + ", ".join(extras))
    so_ativo = [c for c in b.df.columns if c not in novo.columns]
    if so_ativo:
        avisos.append("Variáveis só do banco ativo (ficam ausentes nos casos adicionados): " + ", ".join(so_ativo))
    n0 = len(b.df)
    b.df = pd.concat([b.df, novo], ignore_index=True, sort=False)
    for c_out in outro.colunas:
        destino = mapa.get(c_out, c_out)
        _traz_metadados(b, outro, c_out, destino)
    sint = f"ADD FILES /FILE=* /FILE={_txt(outro.nome)}"
    if indicador:
        indicador = validar_nome_variavel(indicador)
        if _col_ou_none(b, indicador):
            raise ErroTratamento(f"Já existe uma variável chamada '{indicador}'.")
        b.df[indicador] = np.r_[np.zeros(n0), np.ones(len(novo))]
        b.medida[indicador] = "nominal"
        b.rotulos_valor[indicador] = {0.0: "Banco ativo", 1.0: "Banco adicionado"}
        sint += f" /IN={indicador}"
    sint += ".\nEXECUTE."
    return Resultado(b, sint, f"ADD FILES com {outro.nome}: +{len(novo)} casos", avisos)


def adicionar_variaveis(banco: Banco, outro: Banco, chaves: Sequence[str] = (), tipo: str = "consulta",
                        trazer: Optional[Sequence[str]] = None) -> Resultado:
    """tipo: 'consulta' (mantém só os casos do banco ativo — /TABLE) | 'ambos' (todos os casos — /FILE)."""
    if tipo not in ("consulta", "ambos"):
        raise ErroTratamento("Tipo de junção inválido.")
    b = banco.copia()
    avisos: List[str] = []
    chaves = [c for c in chaves if str(c).strip()]
    novo = outro.df.copy()
    nome_arq = _txt(outro.nome)
    if chaves:
        ks_ativo = _lista_vars(b, chaves)
        ks_outro = []
        for k in ks_ativo:
            ach = [c for c in novo.columns if str(c).lower() == k.lower()]
            if not ach:
                raise ErroTratamento(f"A chave '{k}' não existe no segundo banco.")
            ks_outro.append(ach[0])
            if eh_texto(b.df[k]) != eh_texto(novo[ach[0]]):
                raise ErroTratamento(f"A chave '{k}' tem tipos diferentes nos dois bancos.")
        if novo.duplicated(subset=ks_outro).any():
            raise ErroTratamento("O segundo banco tem chaves repetidas; a junção por chave exige uma linha por chave nele.")
        if tipo == "ambos" and b.df.duplicated(subset=ks_ativo).any():
            raise ErroTratamento("O banco ativo tem chaves repetidas; para juntar todos os casos use chave única nos dois lados.")
        ren = dict(zip(ks_outro, ks_ativo))
        novo = novo.rename(columns=ren)
        candidatos = [c for c in novo.columns if c not in ks_ativo]
    else:
        candidatos = list(novo.columns)
    ativos_low = {c.lower() for c in b.colunas}
    repetidas = [c for c in candidatos if str(c).lower() in ativos_low]
    if repetidas:
        avisos.append("Variáveis com o mesmo nome nos dois bancos mantêm os valores do banco ativo: " + ", ".join(repetidas))
    trazidas = [c for c in candidatos if str(c).lower() not in ativos_low]
    if trazer is not None:
        pedidas = {str(t).lower() for t in trazer}
        trazidas = [c for c in trazidas if str(c).lower() in pedidas]
    if not trazidas:
        raise ErroTratamento("Nenhuma variável nova para trazer do segundo banco.")
    usa = ((ks_ativo if chaves else []) + trazidas)
    novo = novo[usa]
    if chaves:
        b.df = b.df.merge(novo, on=ks_ativo, how="left" if tipo == "consulta" else "outer", sort=(tipo == "ambos"))
        if tipo == "consulta":
            sem = b.df[trazidas[0]].isna().sum()
            if sem:
                avisos.append(f"{int(sem)} caso(s) do banco ativo não acharam correspondência (variáveis novas ficam ausentes).")
    else:
        n = max(len(b.df), len(novo))
        b.df = pd.concat([b.df.reindex(range(n)), novo.reset_index(drop=True).reindex(range(n))], axis=1)
        avisos.append("Junção pela ORDEM dos casos (sem chave): confira se os dois bancos estão na mesma ordem.")
    for c in trazidas:
        _traz_metadados(b, outro, c, c)
    lst = _lista_sintaxe(ks_ativo) if chaves else ""
    sint = ""
    if chaves:
        sint += f"* O SPSS exige os dois arquivos ordenados por {lst} (SORT CASES BY {lst}) antes do MATCH FILES.\n"
    sint += "MATCH FILES /FILE=*"
    sint += f" /{'TABLE' if (chaves and tipo == 'consulta') else 'FILE'}={nome_arq}"
    if chaves:
        sint += f" /BY {lst}"
    descartadas = [c for c in candidatos if c not in trazidas and str(c).lower() not in ativos_low]
    if descartadas:
        sint += f" /DROP={_lista_sintaxe(descartadas)}"
    sint += ".\nEXECUTE."
    return Resultado(b, sint, f"MATCH FILES com {outro.nome}: +{len(trazidas)} variável(is)", avisos)


FUNCOES_AGREGACAO = {
    "SUM": "Soma", "MEAN": "Média", "MEDIAN": "Mediana", "SD": "Desvio-padrão", "MIN": "Mínimo",
    "MAX": "Máximo", "FIRST": "Primeiro valor", "LAST": "Último valor", "N": "Nº de casos no grupo",
    "NVALID": "Nº de válidos", "NMISS": "Nº de ausentes",
}


def agregar(banco: Banco, quebras: Sequence[str], agregacoes: Sequence[Dict[str, str]],
            modo: str = "adicionar") -> Resultado:
    """agregacoes: [{"novo": nome, "funcao": "MEAN", "variavel": "Idade"}]  (N não usa variável)."""
    if modo not in ("adicionar", "substituir"):
        raise ErroTratamento("Modo inválido.")
    if not agregacoes:
        raise ErroTratamento("Informe ao menos uma variável agregada.")
    b = banco.copia()
    ks = [_col(b, q) for q in quebras if str(q).strip()]
    n = len(b.df)
    novos: Dict[str, np.ndarray] = {}
    grp_id = (b.df.groupby(ks, sort=False, dropna=False).ngroup().to_numpy() if ks else np.zeros(n, dtype=int))
    ngr = int(grp_id.max()) + 1 if n else 0
    espec = []
    for ag in agregacoes:
        nome = validar_nome_variavel(ag.get("novo", ""))
        f = str(ag.get("funcao", "")).upper()
        if f not in FUNCOES_AGREGACAO:
            raise ErroTratamento(f"Função de agregação desconhecida: {f}")
        if nome in novos or (modo == "adicionar" and _col_ou_none(b, nome)) or (modo == "substituir" and nome.lower() in {k.lower() for k in ks}):
            raise ErroTratamento(f"Nome de variável repetido: '{nome}'.")
        var = None
        if f != "N":
            var = _col(b, ag.get("variavel", ""))
            if eh_texto(b.df[var]) and f not in ("FIRST", "LAST", "NVALID", "NMISS"):
                raise ErroTratamento(f"{f} não se aplica à variável de texto '{var}'.")
        if f == "N":
            res = np.bincount(grp_id, minlength=ngr).astype(float)
        elif eh_texto(b.df[var]):
            x = valores_texto(b, var)
            res = np.empty(ngr, dtype=object)
            if f in ("FIRST", "LAST"):
                for g in range(ngr):
                    idx = np.where(grp_id == g)[0]
                    res[g] = x[idx[0] if f == "FIRST" else idx[-1]]
            else:
                ausente = np.array([e is None for e in x])
                res = np.bincount(grp_id, weights=(~ausente if f == "NVALID" else ausente).astype(float), minlength=ngr)
        else:
            x = valores_numericos(b, var)
            ok = ~np.isnan(x)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                cont = np.bincount(grp_id, weights=ok.astype(float), minlength=ngr)
                soma = np.bincount(grp_id, weights=np.where(ok, x, 0.0), minlength=ngr)
                if f == "NVALID":
                    res = cont
                elif f == "NMISS":
                    res = np.bincount(grp_id, minlength=ngr) - cont
                elif f == "SUM":
                    res = np.where(cont > 0, soma, np.nan)
                elif f == "MEAN":
                    res = soma / np.where(cont == 0, np.nan, cont)
                else:  # MIN, MAX, MEDIAN, SD, FIRST, LAST: calculado grupo a grupo
                    res = np.full(ngr, np.nan)
                    for g in range(ngr):
                        no_grupo = grp_id == g
                        if f in ("FIRST", "LAST"):
                            v = x[no_grupo]          # primeiro/último valor do grupo (mesmo se ausente)
                            if len(v):
                                res[g] = v[0] if f == "FIRST" else v[-1]
                            continue
                        v = x[no_grupo & ok]
                        if not len(v):
                            continue
                        if f == "SD":
                            res[g] = np.std(v, ddof=1) if len(v) > 1 else np.nan
                        else:
                            res[g] = {"MIN": np.min, "MAX": np.max, "MEDIAN": np.median}[f](v)
        novos[nome] = res
        espec.append((nome, f, var))
    sint = ["AGGREGATE", "  /OUTFILE=*" + (" MODE=ADDVARIABLES" if modo == "adicionar" else "")]
    if ks:
        sint.append(f"  /BREAK={_lista_sintaxe(ks)}")
    for nome, f, var in espec:
        sint.append(f"  /{nome} = {f}{'(' + var + ')' if var else ''}")
    sint[-1] += "."
    sint.append("EXECUTE.")
    if modo == "adicionar":
        for nome, res in novos.items():
            b.df[nome] = res[grp_id] if n else res
            b.medida[nome] = "scale"
    else:
        primeiro = pd.Series(range(n)).groupby(grp_id).first().to_numpy() if n else np.array([], dtype=int)
        saida = b.df.iloc[primeiro][ks].reset_index(drop=True) if ks else pd.DataFrame(index=range(ngr))
        for nome, res in novos.items():
            saida[nome] = res
        if ks:
            saida = _ordenar_df(saida, [(k, True) for k in ks]).reset_index(drop=True)
        manter = set(ks)
        b.df = saida
        b.rotulos_var = {k: v for k, v in b.rotulos_var.items() if k in manter}
        b.rotulos_valor = {k: v for k, v in b.rotulos_valor.items() if k in manter}
        b.missing = {k: v for k, v in b.missing.items() if k in manter}
        b.medida = {k: v for k, v in b.medida.items() if k in manter}
        for nome in novos:
            b.medida[nome] = "scale"
    desc = "AGGREGATE " + ("(adiciona variáveis) " if modo == "adicionar" else "(novo banco agregado) ") + \
        (f"por {_lista_sintaxe(ks)}" if ks else "sem quebra")
    return Resultado(b, "\n".join(sint), desc,
                     [] if modo == "adicionar" else [f"O banco agora tem {len(b.df)} linha(s), uma por grupo."])


# =============================================================================
# Registro de operações, histórico, syntax
# =============================================================================
OPERACOES: Dict[str, Callable[..., Resultado]] = {
    "compute": compute,
    "recode": recode,
    "selecionar_casos": selecionar_casos,
    "ordenar_casos": ordenar_casos,
    "remover_duplicados": remover_duplicados,
    "excluir_variaveis": excluir_variaveis,
    "renomear_variavel": renomear_variavel,
    "rotular_variavel": rotular_variavel,
    "rotular_valores": rotular_valores,
    "definir_missing": definir_missing,
    "definir_medida": definir_medida,
    "adicionar_casos": adicionar_casos,
    "adicionar_variaveis": adicionar_variaveis,
    "agregar": agregar,
}


def aplicar(banco: Banco, operacao: str, **params: Any) -> Tuple[Banco, Passo]:
    """Executa uma operação e devolve (novo banco, passo para o histórico)."""
    if operacao not in OPERACOES:
        raise ErroTratamento(f"Operação desconhecida: {operacao}")
    res = OPERACOES[operacao](banco, **params)
    return res.banco, Passo(operacao, dict(params), res.descricao, res.sintaxe, list(res.avisos))


def reaplicar(original: Banco, passos: Sequence[Passo]) -> Banco:
    """Refaz o histórico a partir do banco original (usado no 'desfazer')."""
    b = original
    for p in passos:
        b = OPERACOES[p.operacao](b, **p.params).banco
    return b


def gerar_sintaxe(nome_origem: str, passos: Sequence[Passo], nome_saida: Optional[str] = None) -> str:
    base = os.path.splitext(os.path.basename(nome_origem))[0]
    saida = nome_saida or f"{base}_tratado.sav"
    linhas = [
        "* Syntax gerada pelo módulo Tratamento de Dados (Informa-Automatizado).",
        f"GET FILE={_txt(nome_origem)}.",
        "",
    ]
    for i, p in enumerate(passos, 1):
        desc = p.descricao.splitlines()[0].rstrip(".")
        linhas.append(f"* Passo {i}: {desc}.")
        linhas.append(p.sintaxe)
        linhas.append("")
    linhas.append(f"SAVE OUTFILE={_txt(saida)}.")
    return "\n".join(linhas) + "\n"


# =============================================================================
# Leitura / gravação / exportação
# =============================================================================
def carregar_banco(conteudo: bytes, nome: str = "banco.sav") -> Banco:
    """Lê um .sav mantendo os códigos de missing de usuário (e as definições deles)."""
    from core.pyreadstat_seguro import pyreadstat

    with tempfile.TemporaryDirectory() as tmp:
        caminho = os.path.join(tmp, "entrada.sav")
        with open(caminho, "wb") as f:
            f.write(conteudo)
        df, meta = pyreadstat.read_sav(caminho, user_missing=True)
    nomes = list(df.columns)
    rotulos = {k: v for k, v in (getattr(meta, "column_names_to_labels", {}) or {}).items() if v}
    vl = {k: dict(v) for k, v in (getattr(meta, "variable_value_labels", {}) or {}).items() if v}
    miss = {k: [dict(f) for f in v] for k, v in (getattr(meta, "missing_ranges", {}) or {}).items() if v}
    med = {k: v for k, v in (getattr(meta, "variable_measure", {}) or {}).items() if v in MEDIDAS}
    for c in nomes:
        if c not in med:
            med[c] = "nominal" if eh_texto(df[c]) else "scale"
    return Banco(df=df, rotulos_var=rotulos, rotulos_valor=vl, missing=miss, medida=med, nome=nome)


def _df_para_gravar(banco: Banco) -> Tuple[pd.DataFrame, List[str]]:
    df = banco.df.copy()
    avisos: List[str] = []
    for c in df.columns:
        if eh_texto(df[c]):
            df[c] = pd.Series([None if (x is None or (isinstance(x, float) and math.isnan(x)) or x is pd.NA) else str(x)
                               for x in df[c].tolist()], index=df.index, dtype=object)
    return df, avisos


def gravar_sav(banco: Banco) -> Tuple[bytes, List[str]]:
    """Grava o banco em .sav com rótulos, missings e medida. Devolve (bytes, avisos)."""
    from core.pyreadstat_seguro import pyreadstat

    df, avisos = _df_para_gravar(banco)
    rotulos = {c: banco.rotulos_var.get(c, "") for c in df.columns}
    vl = {c: dict(t) for c, t in banco.rotulos_valor.items() if c in df.columns and t}
    miss = {c: [dict(f) for f in t] for c, t in banco.missing.items() if c in df.columns and t}
    med = {c: banco.medida.get(c, "nominal") for c in df.columns}

    if not pyreadstat.usando_leitor_puro:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                caminho = os.path.join(tmp, "saida.sav")
                kw = dict(file_label="Tratamento de Dados", column_labels=[rotulos[c] for c in df.columns],
                          variable_value_labels=vl, variable_measure=med, compress=True)
                if miss:
                    kw["missing_ranges"] = miss
                pyreadstat.write_sav(df, caminho, **kw)
                with open(caminho, "rb") as f:
                    return f.read(), avisos
        except Exception as exc:  # pyreadstat recusou algum tipo/parâmetro: cai para o escritor em Python puro
            avisos.append(f"O pyreadstat não conseguiu gravar ({exc}); usei o escritor alternativo (sem compressão).")

    from controle.exporters.sav_sem_pyreadstat import write_sav_bytes as gravar_puro

    for c in list(df.columns):
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            ref = pd.Timestamp("1582-10-14")
            df[c] = (df[c] - ref).dt.total_seconds()
            avisos.append(f"'{c}' é data e foi gravada como número de segundos (formato de data não preservado).")
    decimais = {}
    for c in df.columns:
        if not eh_texto(df[c]):
            v = pd.to_numeric(df[c], errors="coerce").dropna()
            decimais[c] = 0 if (len(v) and np.all(np.mod(v, 1) == 0)) else 2
    dados, av2 = gravar_puro(df, {}, var_labels={k: v for k, v in rotulos.items() if v}, value_labels=vl,
                             file_label="Tratamento de Dados", missing=miss, medida=med, decimais=decimais)
    return dados, avisos + list(av2)


def tabela_variaveis(banco: Banco) -> pd.DataFrame:
    """Visão 'Variable View': uma linha por variável."""
    linhas = []
    for c in banco.colunas:
        s = banco.df[c]
        linhas.append({
            "Variável": c,
            "Tipo": "Texto" if eh_texto(s) else "Numérica",
            "Rótulo": banco.rotulos_var.get(c, ""),
            "Valores": "; ".join(f"{_num(k) if not isinstance(k, str) else k}={v}"
                                 for k, v in list(banco.rotulos_valor.get(c, {}).items())[:6])
                       + (" ..." if len(banco.rotulos_valor.get(c, {})) > 6 else ""),
            "Missing": _missing_texto(banco.missing.get(c, [])),
            "Medida": banco.medida.get(c, ""),
            "Válidos": int(s.notna().sum()),
        })
    return pd.DataFrame(linhas)


def exportar_excel(banco: Banco, com_rotulos: bool = False) -> bytes:
    """Excel com os dados (códigos ou rótulos) e uma aba 'Variáveis'."""
    dados = banco.df.copy()
    if com_rotulos:
        for c, mapa in banco.rotulos_valor.items():
            if c in dados.columns and not eh_texto(dados[c]):
                dados[c] = dados[c].map(lambda v, m=mapa: m.get(v, v) if pd.notna(v) else v)
    try:
        import xlsxwriter  # noqa: F401  (bem mais rápido que o openpyxl para bancos grandes)
        motor = "xlsxwriter"
    except ImportError:
        motor = "openpyxl"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=motor) as xw:
        dados.to_excel(xw, sheet_name="Dados", index=False)
        tabela_variaveis(banco).to_excel(xw, sheet_name="Variáveis", index=False)
    return buf.getvalue()


def exportar_csv(banco: Banco, com_rotulos: bool = False) -> bytes:
    dados = banco.df.copy()
    if com_rotulos:
        for c, mapa in banco.rotulos_valor.items():
            if c in dados.columns and not eh_texto(dados[c]):
                dados[c] = dados[c].map(lambda v, m=mapa: m.get(v, v) if pd.notna(v) else v)
    return dados.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
