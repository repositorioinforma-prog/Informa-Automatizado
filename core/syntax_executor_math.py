"""
Executor de Syntax (.sps) sobre um banco (.sav) — lógica pura, sem Streamlit.

Etapa 1 do "Tratamento do Banco": roda, em Python, a parte de recodificação
da Syntax do SPSS (COMPUTE, RECODE, VARIABLE LABELS, VALUE LABELS, MRSETS)
sobre o banco original e devolve o banco tratado, um relatório por comando
e as conferências.

Princípios:
- Estrito: comando que o executor não conhece NÃO é ignorado em silêncio —
  vira um aviso (e, se mexeria nos dados, um "possível erro").
- Reaproveita o parser do Revisor de Syntax (`decode_sps`, `split_commands`,
  rótulos), para as duas ferramentas lerem a Syntax do mesmo jeito.
- Semântica do RECODE do SPSS: para cada caso, a primeira regra que casa
  vence (não há "cascata"); ELSE pega tudo que sobrou; valor não citado
  fica como está quando o RECODE é "no lugar" e é tratado conforme
  `NAO_CITADO_INTO` quando há INTO (ver a constante).
"""

from __future__ import annotations

import ast
import operator
import re
from dataclasses import dataclass, field
from typing import Any, Iterable

import numpy as np
import pandas as pd

from core.revisor_syntax_math import (
    Finding,
    _parse_value_labels,
    _parse_variable_labels,
    _strip_inline_comments,
    decode_sps,
    split_commands,
)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

#: O que acontece com valor NÃO citado (e sem ELSE) num RECODE ... INTO
#: variável NOVA. No SPSS, o valor vira "system-missing". (Conferido contra
#: o banco que o próprio SPSS gerou — ver tests/test_syntax_executor.py.)
NAO_CITADO_INTO = "sysmis"

_NOME = r"[A-Za-z_@$#][\w@$#\.]*"

COMANDOS_SEM_EFEITO = {
    "EXECUTE", "SET", "TITLE", "SUBTITLE", "DATASET", "OUTPUT", "NEW",
    "COMMENT", "PRESERVE", "RESTORE", "SHOW", "DISPLAY",
}
COMANDOS_SO_LEITURA = {
    "FREQUENCIES", "TABLES", "CTABLES", "CROSSTABS", "DESCRIPTIVES",
    "MEANS", "EXAMINE", "GRAPH", "CORRELATIONS", "REGRESSION", "SUMMARIZE",
    "REPORT",
}
COMANDOS_QUE_MEXEM_NOS_DADOS = {
    "IF", "DO", "COUNT", "STRING", "NUMERIC", "FORMATS", "MISSING",
    "SELECT", "DELETE", "FILTER", "WEIGHT", "SORT", "AGGREGATE", "MATCH",
    "ADD", "VARIABLE", "TEMPORARY", "AUTORECODE", "RANK", "COMPUTE",
    "RECODE", "SAVE", "GET", "ALTER", "RENAME", "COMPARE", "LEAVE",
}


# ---------------------------------------------------------------------------
# Estruturas
# ---------------------------------------------------------------------------

@dataclass
class Condicao:
    tipo: str                      # val | faixa | missing | sysmis | else | texto
    a: Any = None
    b: Any = None


@dataclass
class RegraRecode:
    condicoes: list[Condicao]
    destino: str                   # copy | sysmis | num | texto
    valor: Any = None
    bruto: str = ""


@dataclass
class ComandoRecode:
    linha: int
    bruto: str
    origens: list[str]
    destinos: list[str]
    regras: list[list[RegraRecode]]     # uma lista de regras por origem
    com_into: bool


@dataclass
class RegistroComando:
    """Uma linha do relatório de execução."""
    linha: int
    tipo: str                      # RECODE | COMPUTE | VAR LABEL | VALUE LABEL | MRSETS | OUTRO
    resumo: str
    status: str                    # ok | ignorado | nao_suportado | erro
    detalhe: str = ""
    variaveis: list[str] = field(default_factory=list)


@dataclass
class ResultadoExecucao:
    df: pd.DataFrame
    var_labels: dict[str, str]
    value_labels: dict[str, dict[Any, str]]
    mrsets: dict[str, list[str]]
    registros: list[RegistroComando]
    criadas: list[str]                       # variáveis novas, na ordem de criação
    modificadas: list[str]                   # variáveis existentes alteradas
    recodes: list[ComandoRecode]
    findings: list[Finding]

    def contagem_por_status(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for r in self.registros:
            out[r.status] = out.get(r.status, 0) + 1
        return out


class ErroDeSyntax(Exception):
    pass


# ---------------------------------------------------------------------------
# Tokenização do RECODE
# ---------------------------------------------------------------------------

def _fora_de_aspas(texto: str) -> list[tuple[int, str, int]]:
    """Devolve (posição, caractere, profundidade_de_parênteses) só para os
    caracteres que estão fora de aspas."""
    out = []
    quote = None
    prof = 0
    for i, ch in enumerate(texto):
        if quote:
            if ch == quote:
                quote = None
            continue
        if ch in "'\"":
            quote = ch
            continue
        if ch == "(":
            out.append((i, ch, prof))
            prof += 1
            continue
        if ch == ")":
            prof = max(prof - 1, 0)
            out.append((i, ch, prof))
            continue
        out.append((i, ch, prof))
    return out


def _tokens(texto: str) -> list[str]:
    """Divide em espaços/vírgulas fora de aspas; mantém as aspas no token."""
    toks: list[str] = []
    buf: list[str] = []
    quote = None
    for ch in texto:
        if quote:
            buf.append(ch)
            if ch == quote:
                quote = None
            continue
        if ch in "'\"":
            quote = ch
            buf.append(ch)
        elif ch.isspace() or ch == ",":
            if buf:
                toks.append("".join(buf))
                buf = []
        else:
            buf.append(ch)
    if buf:
        toks.append("".join(buf))
    return toks


def _numero(tok: str) -> float | None:
    try:
        return float(tok)
    except ValueError:
        return None


def _valor_literal(tok: str) -> tuple[str, Any]:
    if len(tok) >= 2 and tok[0] in "'\"" and tok[-1] == tok[0]:
        return "texto", tok[1:-1]
    n = _numero(tok)
    if n is None:
        raise ErroDeSyntax(f"valor não reconhecido: {tok!r}")
    return "num", n


def _limite(tok: str, baixo: bool) -> float:
    u = tok.upper()
    if u in {"LO", "LOW", "LOWEST"}:
        return -np.inf
    if u in {"HI", "HIGH", "HIGHEST"}:
        return np.inf
    n = _numero(tok)
    if n is None:
        raise ErroDeSyntax(f"limite de faixa não reconhecido: {tok!r}")
    return n


def _parse_condicoes(esquerda: str) -> list[Condicao]:
    toks = _tokens(esquerda)
    conds: list[Condicao] = []
    i = 0
    while i < len(toks):
        t = toks[i]
        u = t.upper()
        if i + 2 < len(toks) and toks[i + 1].upper() == "THRU":
            conds.append(Condicao("faixa", _limite(t, True), _limite(toks[i + 2], False)))
            i += 3
            continue
        if u == "ELSE":
            conds.append(Condicao("else"))
        elif u == "MISSING":
            conds.append(Condicao("missing"))
        elif u in {"SYSMIS", "$SYSMIS"}:
            conds.append(Condicao("sysmis"))
        else:
            tipo, v = _valor_literal(t)
            conds.append(Condicao("texto" if tipo == "texto" else "val", v))
        i += 1
    if not conds:
        raise ErroDeSyntax("regra sem valores de origem")
    return conds


def _parse_regra(grupo: str) -> RegraRecode:
    # o "=" que separa origem e destino é o último fora de aspas
    pos = [i for i, ch, _ in _fora_de_aspas(grupo) if ch == "="]
    if not pos:
        raise ErroDeSyntax(f"regra sem '=': ({grupo})")
    k = pos[-1]
    esq, dir_ = grupo[:k], grupo[k + 1:].strip()
    conds = _parse_condicoes(esq)
    u = dir_.upper()
    if u == "COPY":
        return RegraRecode(conds, "copy", bruto=grupo.strip())
    if u in {"SYSMIS", "$SYSMIS"}:
        return RegraRecode(conds, "sysmis", bruto=grupo.strip())
    tipo, v = _valor_literal(dir_)
    return RegraRecode(conds, "texto" if tipo == "texto" else "num", v, grupo.strip())


def _grupos_parenteses(texto: str) -> tuple[str, list[str]]:
    """(texto antes do 1º parêntese, conteúdo de cada grupo de nível 0)."""
    cars = _fora_de_aspas(texto)
    antes_fim = None
    grupos: list[str] = []
    ini = None
    for i, ch, prof in cars:
        if ch == "(" and prof == 0:
            if antes_fim is None:
                antes_fim = i
            ini = i + 1
        elif ch == ")" and prof == 0 and ini is not None:
            grupos.append(texto[ini:i])
            ini = None
    antes = texto[: antes_fim if antes_fim is not None else len(texto)]
    return antes, grupos


def parse_recode_completo(linha: int, bruto: str) -> ComandoRecode:
    m = re.match(r"\s*RECODE\s+(.+?)\.?\s*$", bruto, flags=re.I | re.S)
    if not m:
        raise ErroDeSyntax("RECODE mal formado")
    corpo = m.group(1)

    # INTO no nível 0 (fora de aspas e parênteses)
    prof_em = {i: prof for i, _, prof in _fora_de_aspas(corpo)}
    pos_into = None
    for mm in re.finditer(r"\bINTO\b", corpo, flags=re.I):
        if mm.start() in prof_em and prof_em[mm.start()] == 0:
            pos_into = mm.start()
            break
    if pos_into is not None:
        antes_into, depois_into = corpo[:pos_into], corpo[pos_into + 4:]
        destinos = [v.upper() for v in re.findall(_NOME, depois_into)]
    else:
        antes_into, destinos = corpo, []

    # seções separadas por "/" no nível 0
    cortes = [i for i, ch, prof in _fora_de_aspas(antes_into) if ch == "/" and prof == 0]
    secoes: list[str] = []
    ant = 0
    for c in cortes:
        secoes.append(antes_into[ant:c])
        ant = c + 1
    secoes.append(antes_into[ant:])

    origens: list[str] = []
    regras_por_origem: list[list[RegraRecode]] = []
    for sec in secoes:
        antes, grupos = _grupos_parenteses(sec)
        vars_ = [v.upper() for v in re.findall(_NOME, antes)]
        if not vars_:
            raise ErroDeSyntax("RECODE sem variável de origem")
        regras = [_parse_regra(g) for g in grupos]
        if not regras:
            raise ErroDeSyntax("RECODE sem regras")
        for v in vars_:
            origens.append(v)
            regras_por_origem.append(regras)

    if destinos and len(destinos) != len(origens):
        raise ErroDeSyntax(
            f"RECODE tem {len(origens)} variável(is) de origem e {len(destinos)} de destino (INTO)"
        )
    return ComandoRecode(linha, bruto.strip(), origens, destinos or list(origens), regras_por_origem, bool(destinos))


# ---------------------------------------------------------------------------
# Execução vetorizada do RECODE
# ---------------------------------------------------------------------------

def _eh_texto(serie: pd.Series) -> bool:
    return not pd.api.types.is_numeric_dtype(serie)


def _mascara(cond: Condicao, x: np.ndarray, texto: bool, missing_usuario: Iterable[Any]) -> np.ndarray:
    if texto:
        s = pd.Series(x, dtype=object)
        vazio = s.isna() | (s.astype(str).str.strip() == "")
        s = s.astype(str).str.rstrip()
        if cond.tipo == "else":
            return np.ones(len(x), dtype=bool)
        if cond.tipo in {"missing", "sysmis"}:
            return vazio.to_numpy()
        if cond.tipo in {"texto", "val"}:
            return (s == str(cond.a).rstrip()).to_numpy() & ~vazio.to_numpy()
        return np.zeros(len(x), dtype=bool)

    nan = np.isnan(x)
    if cond.tipo == "else":
        return np.ones(len(x), dtype=bool)
    if cond.tipo == "sysmis":
        return nan
    if cond.tipo == "missing":
        out = nan.copy()
        for mv in missing_usuario:
            if isinstance(mv, tuple):
                out |= (x >= mv[0]) & (x <= mv[1])
            else:
                out |= x == mv
        return out
    if cond.tipo == "val":
        return x == cond.a
    if cond.tipo == "faixa":
        return (x >= cond.a) & (x <= cond.b)
    return np.zeros(len(x), dtype=bool)


def aplicar_regras(
    serie_origem: pd.Series,
    regras: list[RegraRecode],
    destino_atual: pd.Series | None = None,
    missing_usuario: Iterable[Any] = (),
) -> tuple[pd.Series, dict[str, Any]]:
    """Aplica as regras a uma coluna. `destino_atual` é o conteúdo prévio da
    variável de destino (RECODE no lugar ou INTO em variável existente)."""
    n = len(serie_origem)
    texto_origem = _eh_texto(serie_origem)
    alvo_texto = any(r.destino == "texto" for r in regras)

    if texto_origem:
        x = serie_origem.to_numpy(dtype=object)
    else:
        x = serie_origem.to_numpy(dtype=float)

    if alvo_texto:
        saida = np.full(n, None, dtype=object)
        base = destino_atual.to_numpy(dtype=object) if destino_atual is not None else None
    else:
        saida = np.full(n, np.nan, dtype=float)
        base = None
        if destino_atual is not None and not _eh_texto(destino_atual):
            base = destino_atual.to_numpy(dtype=float)

    atribuido = np.zeros(n, dtype=bool)
    por_regra = []
    for regra in regras:
        casou = np.zeros(n, dtype=bool)
        for c in regra.condicoes:
            casou |= _mascara(c, x, texto_origem, missing_usuario)
        casou &= ~atribuido
        if regra.destino == "copy":
            if alvo_texto or texto_origem:
                saida[casou] = x[casou]
            else:
                saida[casou] = x[casou]
        elif regra.destino == "sysmis":
            saida[casou] = None if alvo_texto else np.nan
        else:
            saida[casou] = regra.valor
        atribuido |= casou
        por_regra.append(int(casou.sum()))

    nao_citado = ~atribuido
    if destino_atual is not None and base is not None:
        saida[nao_citado] = base[nao_citado]
    elif destino_atual is None and NAO_CITADO_INTO == "copia":
        saida[nao_citado] = x[nao_citado]
    # senão: fica system-missing

    if alvo_texto:
        res = pd.Series(saida, index=serie_origem.index, dtype=object)
    else:
        res = pd.Series(saida, index=serie_origem.index, dtype="float64")
    info = {
        "por_regra": por_regra,
        "nao_citados": int(nao_citado.sum()),
        "mascara_nao_citado": nao_citado,
        "total": n,
    }
    return res, info


# ---------------------------------------------------------------------------
# COMPUTE (expressões simples e seguras)
# ---------------------------------------------------------------------------

_OPS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.Pow: operator.pow,
}


def _avaliar(no: ast.AST, df: pd.DataFrame):
    if isinstance(no, ast.Expression):
        return _avaliar(no.body, df)
    if isinstance(no, ast.Constant) and isinstance(no.value, (int, float)):
        return float(no.value)
    if isinstance(no, ast.Name):
        for c in df.columns:
            if str(c).upper() == no.id.upper():
                return pd.to_numeric(df[c], errors="coerce")
        raise ErroDeSyntax(f"variável {no.id} não existe no banco")
    if isinstance(no, ast.UnaryOp) and isinstance(no.op, (ast.USub, ast.UAdd)):
        v = _avaliar(no.operand, df)
        return -v if isinstance(no.op, ast.USub) else v
    if isinstance(no, ast.BinOp) and type(no.op) in _OPS:
        return _OPS[type(no.op)](_avaliar(no.left, df), _avaliar(no.right, df))
    raise ErroDeSyntax("expressão do COMPUTE não suportada")


def _executar_compute(df: pd.DataFrame, bruto: str) -> tuple[str, pd.Series]:
    m = re.match(rf"\s*COMPUTE\s+({_NOME})\s*=\s*(.+?)\.?\s*$", bruto, flags=re.I | re.S)
    if not m:
        raise ErroDeSyntax("COMPUTE mal formado")
    alvo = m.group(1).upper()
    expr = m.group(2).strip().replace("**", "^").replace("^", "**")
    if re.search(r"[A-Za-z_]\w*\s*\(", expr):
        raise ErroDeSyntax("COMPUTE com função não suportado")
    try:
        arvore = ast.parse(expr, mode="eval")
    except SyntaxError as e:
        raise ErroDeSyntax(f"expressão inválida: {expr}") from e
    val = _avaliar(arvore, df)
    if not isinstance(val, pd.Series):
        val = pd.Series(val, index=df.index, dtype="float64")
    return alvo, val.astype("float64")


# ---------------------------------------------------------------------------
# MRSETS
# ---------------------------------------------------------------------------

def _parse_mrsets(bruto: str) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for sec in re.split(r"(?=/\s*M[CD]GROUP\b)", bruto, flags=re.I):
        m = re.match(r"\s*/\s*M[CD]GROUP\b(.*)$", sec, flags=re.I | re.S)
        if not m:
            continue
        corpo = m.group(1)
        nome = re.search(r"\bNAME\s*=\s*(\$[\w@$#\.]+)", corpo, flags=re.I)
        vs = re.search(r"\bVARIABLES\s*=\s*(.*?)(?=/|$|\bVALUE\s*=|\bLABEL\s*=)", corpo, flags=re.I | re.S)
        if nome and vs:
            out[nome.group(1).upper()] = [v.upper() for v in re.findall(_NOME, vs.group(1))]
    return out


_CODIGOS_NAO_SE_APLICA_MIN = 9000


def _avisar_valores_sem_regra(findings, serie_origem, mascara, linha, origem, destino, resumo, n_regras):
    if _eh_texto(serie_origem):
        vistos = serie_origem[mascara & (serie_origem.astype(str).str.strip() != "")]
    else:
        vistos = serie_origem[mascara & serie_origem.notna()]
    if vistos.empty:
        return
    contagem = vistos.value_counts().sort_index()
    lista = ", ".join(f"{_fmt_valor(v)} (n={int(c)})" for v, c in contagem.head(8).items())
    if len(contagem) > 8:
        lista += f" … (+{len(contagem) - 8} valores)"
    so_nao_se_aplica = (not _eh_texto(serie_origem)) and bool((contagem.index >= _CODIGOS_NAO_SE_APLICA_MIN).all())
    selecao = n_regras == 1
    if _eh_texto(serie_origem):
        base_valida = int((serie_origem.astype(str).str.strip() != "").sum())
    else:
        base_valida = int((serie_origem.notna() & (serie_origem < _CODIGOS_NAO_SE_APLICA_MIN)).sum())
    fatia = len(vistos) / base_valida if base_valida else 1.0
    if so_nao_se_aplica or selecao or fatia > 0.10:
        if so_nao_se_aplica:
            tipo = "código de 'não se aplica'"
        elif selecao:
            tipo = "recode de seleção (uma única regra)"
        else:
            tipo = f"{fatia:.0%} da base: parece uma separação intencional"
        gravidade = "informativo"
    else:
        gravidade, tipo = "revisar", "valor(es) que nenhuma regra cita"
    findings.append(Finding(
        gravidade, "EXEC_VALOR_SEM_REGRA",
        f"{origem} → {destino}: {int(len(vistos))} caso(s) ficaram vazios",
        f"No RECODE da linha {linha}, estes valores de {origem} não estão em nenhuma regra e ficaram "
        f"vazios em {destino} ({tipo}): {lista}.",
        linha, destino, resumo,
        None if gravidade == "informativo" else "Confirme se esses valores deviam mesmo ficar de fora (ou inclua um ELSE).",
    ))


def _fmt_valor(v) -> str:
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    except (TypeError, ValueError):
        return str(v)


# ---------------------------------------------------------------------------
# Execução da Syntax inteira
# ---------------------------------------------------------------------------

def _achar_coluna(df: pd.DataFrame, nome: str) -> str | None:
    n = nome.upper()
    for c in df.columns:
        if str(c).upper() == n:
            return c
    return None


def _primeira_palavra(cmd: str) -> str:
    m = re.match(r"\s*([A-Za-z]+)", cmd)
    return m.group(1).upper() if m else ""


def _resumo(cmd: str, limite: int = 110) -> str:
    s = " ".join(cmd.split())
    return s if len(s) <= limite else s[: limite - 1] + "…"


def executar_syntax(
    sps_bytes: bytes,
    df: pd.DataFrame,
    var_labels: dict[str, str] | None = None,
    value_labels: dict[str, dict[Any, str]] | None = None,
    missing_ranges: dict[str, list[Any]] | None = None,
) -> ResultadoExecucao:
    """Roda a Syntax sobre uma CÓPIA do banco e devolve o resultado."""
    texto, _enc, achados = decode_sps(sps_bytes)
    findings: list[Finding] = list(achados)
    comandos = split_commands(_sem_comentarios_em_linha(texto))

    d = df.copy()
    labels_var = {str(k).upper(): v for k, v in (var_labels or {}).items()}
    labels_val = {str(k).upper(): dict(v) for k, v in (value_labels or {}).items()}
    miss = {str(k).upper(): v for k, v in (missing_ranges or {}).items()}
    mrsets: dict[str, list[str]] = {}
    registros: list[RegistroComando] = []
    recodes: list[ComandoRecode] = []
    criadas: list[str] = []
    modificadas: list[str] = []

    def nova(nome: str):
        if nome not in criadas:
            criadas.append(nome)

    for linha, bruto in comandos:
        palavra = _primeira_palavra(bruto)
        resumo = _resumo(bruto)

        try:
            # ----- RECODE -------------------------------------------------
            if palavra == "RECODE":
                cmd = parse_recode_completo(linha, bruto)
                recodes.append(cmd)
                detalhes = []
                vars_tocadas = []
                for origem, destino, regras in zip(cmd.origens, cmd.destinos, cmd.regras):
                    col_o = _achar_coluna(d, origem)
                    if col_o is None:
                        msg = f"variável de origem {origem} não existe no banco"
                        registros.append(RegistroComando(linha, "RECODE", resumo, "erro", msg, [origem]))
                        findings.append(Finding(
                            "critico", "EXEC_ORIGEM_AUSENTE", f"{origem} não existe no banco",
                            f"O RECODE da linha {linha} usa {origem}, que não está no banco.",
                            linha, origem, resumo,
                            "Confira o nome da variável (ou use o banco correto).",
                        ))
                        continue
                    col_d = _achar_coluna(d, destino)
                    atual = d[col_d] if col_d is not None else None
                    mv = miss.get(origem, [])
                    res, info = aplicar_regras(d[col_o], regras, atual, mv)
                    nome_col = col_d if col_d is not None else destino
                    d[nome_col] = res
                    if col_d is None:
                        nova(destino)
                    elif destino not in modificadas and destino not in criadas:
                        modificadas.append(destino)
                    vars_tocadas.append(destino)
                    validos = int(res.notna().sum())
                    detalhes.append(f"{origem}→{destino}: {validos} com valor, {info['nao_citados']} sem regra")
                    tem_else = any(c.tipo == "else" for r in regras for c in r.condicoes)
                    if info["nao_citados"] and not tem_else and atual is None:
                        _avisar_valores_sem_regra(findings, d[col_o], info["mascara_nao_citado"],
                                                  linha, origem, destino, resumo, len(regras))
                    if atual is not None:
                        antes = atual.to_numpy(dtype=object)
                        depois = res.to_numpy(dtype=object)
                        mudou = int(sum(
                            1 for a_, b_ in zip(antes, depois)
                            if not ((pd.isna(a_) and pd.isna(b_)) or a_ == b_)
                        ))
                        if mudou:
                            findings.append(Finding(
                                "revisar", "EXEC_VARIAVEL_SOBRESCRITA",
                                f"{destino} já existia e foi alterada ({mudou} caso(s))",
                                f"O RECODE da linha {linha} gravou em {destino}, que já existia no banco, "
                                f"e mudou {mudou} caso(s).",
                                linha, destino, resumo,
                                "Se o banco já veio tratado, rode a Syntax no banco original (sem as variáveis recodificadas).",
                            ))
                registros.append(RegistroComando(
                    linha, "RECODE", resumo,
                    "ok" if vars_tocadas else "erro", " | ".join(detalhes), vars_tocadas,
                ))
                continue

            # ----- COMPUTE ------------------------------------------------
            if palavra == "COMPUTE":
                alvo, serie = _executar_compute(d, bruto)
                col = _achar_coluna(d, alvo)
                d[col if col is not None else alvo] = serie
                if col is None:
                    nova(alvo)
                elif alvo not in modificadas and alvo not in criadas:
                    modificadas.append(alvo)
                registros.append(RegistroComando(linha, "COMPUTE", resumo, "ok", "", [alvo]))
                continue

            # ----- VARIABLE LABELS ---------------------------------------
            if re.match(r"\s*(VARIABLE\s+LABELS|VAR\s+LABELS?)\b", bruto, flags=re.I):
                rot = _parse_variable_labels([(linha, bruto)])
                for v, (txt, _l) in rot.items():
                    labels_var[v] = txt
                registros.append(RegistroComando(linha, "VAR LABEL", resumo, "ok", "", list(rot)))
                continue

            # ----- VALUE LABELS ------------------------------------------
            if re.match(r"\s*(VALUE\s+LABELS?|ADD\s+VALUE\s+LABELS?)\b", bruto, flags=re.I):
                if re.match(r"\s*ADD\s+VALUE", bruto, flags=re.I):
                    bruto_vl = re.sub(r"^\s*ADD\s+", "", bruto, flags=re.I)
                else:
                    bruto_vl = bruto
                rot, _l = _parse_value_labels([(linha, bruto_vl)])
                for v, mp in rot.items():
                    labels_val.setdefault(v, {}).update(mp)
                registros.append(RegistroComando(linha, "VALUE LABEL", resumo, "ok", "", list(rot)))
                continue

            # ----- MRSETS -------------------------------------------------
            if palavra == "MRSETS":
                conj = _parse_mrsets(bruto)
                for nome, membros in conj.items():
                    if nome in mrsets and mrsets[nome] != membros:
                        findings.append(Finding(
                            "informativo", "EXEC_MRSET_REDEFINIDO", f"{nome} foi definido mais de uma vez",
                            f"O conjunto múltiplo {nome} é redefinido na linha {linha}; vale a última definição.",
                            linha, nome, resumo,
                        ))
                    mrsets[nome] = membros
                registros.append(RegistroComando(linha, "MRSETS", resumo, "ok", ", ".join(conj), list(conj)))
                continue

            # ----- sem efeito / só leitura -------------------------------
            if palavra in COMANDOS_SEM_EFEITO:
                registros.append(RegistroComando(linha, palavra, resumo, "ignorado", "sem efeito nos dados"))
                continue
            if palavra in COMANDOS_SO_LEITURA:
                registros.append(RegistroComando(linha, palavra, resumo, "ignorado", "só gera saída; não altera o banco"))
                continue

            # ----- desconhecido ------------------------------------------
            mexe = palavra in COMANDOS_QUE_MEXEM_NOS_DADOS
            registros.append(RegistroComando(linha, palavra or "?", resumo, "nao_suportado",
                                             "o executor ainda não roda este comando"))
            findings.append(Finding(
                "possivel_erro" if mexe else "revisar",
                "EXEC_COMANDO_NAO_SUPORTADO",
                f"Comando {palavra or '?'} não foi executado",
                f"A linha {linha} tem um comando {palavra or '?'} que o executor não sabe rodar"
                + ("; ele pode alterar os dados, então o banco tratado pode divergir do SPSS." if mexe else "."),
                linha, None, resumo,
                "Rode esse trecho no SPSS ou peça o suporte ao comando.",
            ))

        except ErroDeSyntax as e:
            registros.append(RegistroComando(linha, palavra or "?", resumo, "erro", str(e)))
            findings.append(Finding(
                "critico", "EXEC_ERRO_DE_SYNTAX", f"Erro na linha {linha}", str(e),
                linha, None, resumo, "Corrija o comando na Syntax.",
            ))

    return ResultadoExecucao(
        df=d, var_labels=labels_var, value_labels=labels_val, mrsets=mrsets,
        registros=registros, criadas=criadas, modificadas=modificadas,
        recodes=recodes, findings=findings,
    )


def _sem_comentarios_em_linha(texto: str) -> str:
    """Mesmo tratamento do Revisor para comentários inline."""
    return _strip_inline_comments(texto)


def variaveis_geradas(sps_bytes: bytes) -> list[str]:
    """Variáveis que a Syntax CRIA (destinos de RECODE ... INTO e alvos de
    COMPUTE), na ordem em que aparecem. Serve para rodar a partir de um banco
    "limpo", sem as versões que o SPSS já tinha gravado."""
    texto, _enc, _ach = decode_sps(sps_bytes)
    out: list[str] = []
    for linha, bruto in split_commands(_sem_comentarios_em_linha(texto)):
        palavra = _primeira_palavra(bruto)
        try:
            if palavra == "RECODE":
                cmd = parse_recode_completo(linha, bruto)
                if cmd.com_into:
                    for d in cmd.destinos:
                        if d not in out:
                            out.append(d)
            elif palavra == "COMPUTE":
                m = re.match(rf"\s*COMPUTE\s+({_NOME})\s*=", bruto, flags=re.I)
                if m and m.group(1).upper() not in out:
                    out.append(m.group(1).upper())
        except ErroDeSyntax:
            continue
    return out
