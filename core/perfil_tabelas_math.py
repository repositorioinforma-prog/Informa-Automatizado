"""
Perfil — gera as tabelas do Perfil a partir do banco (lógica pura, sem Streamlit).

Substitui rodar à mão os 5 blocos do SPSS «Perfil_<Cidade>_<UF>_<MêsAno>.sps»:

    P1              «Sexo»
    P2              «Idade»            (+ «Média» pela variável MEP2, só em projeto presencial/rua)
    P3              «Renda Familiar»
    P4              «Escolaridade»
    juncaoreligiao  «Religiões»

Cada bloco repete a conta do TABLES do SPSS com a variável de peso ligada:

    cpct(var (F5.1) :total)   -> % do total, ponderado (1 casa só na exibição; o valor guarda a precisão total);
    UNW count($t (F5.0))      -> «Base» = nº de casos válidos SEM ponderar;
    mean(MEP2 (F5.2))         -> «Média» = média ponderada de MEP2 (só presencial).

O Excel sai no mesmo formato que o leitor do slide de perfil (`core.perfil_math.ler_perfil`) já
consome: título do bloco, linha «Total», categorias com o %, «Base», («Média» só na Idade) e uma
linha em branco entre blocos.

Regras do projeto: nunca inventar valor; variável que não está no banco fica de fora (e é avisada),
nunca com zero no lugar; categoria sem nenhum caso válido não aparece.

Dependências: só pandas/numpy/openpyxl.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

from core.analise_perguntas_math import coluna_do_banco, mapa_rotulos, rotulo_categoria

#: (id do bloco, variável do banco, título impresso) — a ordem é a do .sps do Perfil
VARIAVEIS_PERFIL: Tuple[Tuple[str, str, str], ...] = (
    ("sexo", "P1", "Sexo"),
    ("idade", "P2", "Idade"),
    ("renda", "P3", "Renda Familiar"),
    ("escolaridade", "P4", "Escolaridade"),
    ("religiao", "juncaoreligiao", "Religiões"),
)
VAR_MEDIA_IDADE = "MEP2"
PESO_PADRAO = "peso"
MODO_PRESENCIAL = "presencial"
MODO_ONLINE = "online"
NOME_ABA = "Perfil"


# ---------------------------------------------------------------- resultado
@dataclass
class BlocoPerfil:
    id: str
    var: str                       # coluna usada no banco
    titulo: str
    categorias: List[str] = field(default_factory=list)
    pct: List[float] = field(default_factory=list)   # pontos percentuais, precisão total (a exibição é com 1 casa)
    base: int = 0                  # casos válidos, sem ponderar
    media: Optional[float] = None  # só Idade presencial (precisão total; exibição com 2 casas)
    var_media: Optional[str] = None


@dataclass
class ResultadoPerfil:
    blocos: List[BlocoPerfil] = field(default_factory=list)
    ausentes: List[Tuple[str, str]] = field(default_factory=list)   # (variável, título) fora do banco
    sem_dados: List[Tuple[str, str]] = field(default_factory=list)  # no banco, mas sem caso válido
    avisos: List[str] = field(default_factory=list)
    peso: Optional[str] = None
    modo: str = MODO_PRESENCIAL
    casos_banco: int = 0
    casos_sem_peso: int = 0        # casos que saíram por peso ausente ou ≤ 0


# ---------------------------------------------------------------- conta
def _pesos(df: pd.DataFrame, peso: Optional[str]) -> pd.Series:
    """Peso por caso; caso sem peso válido (ausente ou ≤ 0) vira NaN e sai da conta, como o WEIGHT do SPSS."""
    if not peso:
        return pd.Series(1.0, index=df.index)
    w = pd.to_numeric(df[peso], errors="coerce")
    return w.where(w.notna() & (w > 0))


def calcular_bloco(
    df: pd.DataFrame,
    var: str,
    bid: str,
    titulo: str,
    peso: Optional[str],
    value_labels: Optional[Dict[str, Dict[Any, str]]] = None,
    var_media: Optional[str] = None,
) -> Optional[BlocoPerfil]:
    """Tabela de uma variável do Perfil. None se não sobrar nenhum caso válido."""
    w = _pesos(df, peso)
    serie = df[var]
    ok = serie.notna() & w.notna()
    if not bool(ok.any()):
        return None
    base = pd.DataFrame({"c": serie[ok].to_numpy(), "w": w[ok].to_numpy(dtype=float)})
    soma = base.groupby("c")["w"].sum().sort_index()
    total = float(soma.sum())
    if total <= 0:
        return None
    mapa = mapa_rotulos((value_labels or {}).get(var))
    bloco = BlocoPerfil(
        id=bid,
        var=var,
        titulo=titulo,
        categorias=[rotulo_categoria(c, mapa) for c in soma.index],
        pct=[float(v / total * 100.0) for v in soma.to_numpy(dtype=float)],
        base=int(ok.sum()),
    )
    if var_media:
        m = pd.to_numeric(df[var_media], errors="coerce")
        valido = m.notna() & w.notna()
        if bool(valido.any()):
            bloco.media = float(np.average(m[valido], weights=w[valido]))
            bloco.var_media = var_media
    return bloco


def calcular_perfil(
    df: pd.DataFrame,
    value_labels: Optional[Dict[str, Dict[Any, str]]] = None,
    peso: Optional[str] = PESO_PADRAO,
    presencial: bool = True,
) -> ResultadoPerfil:
    """Calcula os blocos do Perfil. Variável que não existe no banco entra em `ausentes` (nada é inventado)."""
    colunas = [str(c) for c in df.columns]
    res = ResultadoPerfil(
        peso=peso,
        modo=MODO_PRESENCIAL if presencial else MODO_ONLINE,
        casos_banco=len(df),
    )
    if peso and peso not in df.columns:
        raise ValueError(f"A variável de peso «{peso}» não está no banco.")
    if peso:
        res.casos_sem_peso = int(_pesos(df, peso).isna().sum())
    col_media = coluna_do_banco(VAR_MEDIA_IDADE, colunas) if presencial else None
    if presencial and not col_media:
        res.avisos.append(
            f"Projeto presencial, mas a variável {VAR_MEDIA_IDADE} não está no banco: a Idade sai sem a «Média»."
        )
    for bid, var, titulo in VARIAVEIS_PERFIL:
        col = coluna_do_banco(var, colunas)
        if not col:
            res.ausentes.append((var, titulo))
            continue
        bloco = calcular_bloco(
            df, col, bid, titulo, peso, value_labels, var_media=col_media if bid == "idade" else None
        )
        if bloco is None:
            res.sem_dados.append((col, titulo))
            continue
        res.blocos.append(bloco)
    if res.ausentes:
        res.avisos.append(
            "Variáveis que não estão no banco (blocos não gerados): "
            + ", ".join(f"{v} ({t})" for v, t in res.ausentes) + "."
        )
    if res.sem_dados:
        res.avisos.append(
            "Variáveis sem nenhum caso válido (blocos não gerados): "
            + ", ".join(f"{v} ({t})" for v, t in res.sem_dados) + "."
        )
    if res.casos_sem_peso:
        res.avisos.append(
            f"{res.casos_sem_peso} caso(s) com «{peso}» ausente ou ≤ 0 ficaram de fora, como no WEIGHT do SPSS."
        )
    return res


# ---------------------------------------------------------------- nome do arquivo
_RE_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RE_UF_MES = re.compile(r"(?:^|[_\s-])([A-Z]{2})[_\s-]+([A-Za-zçÇ]{3,9}\d{2})(?=$|[_\s.-])")


def _parte_nome(texto: str) -> str:
    texto = _RE_INVALIDOS.sub("", " ".join(str(texto or "").split()))
    return re.sub(r"\s+", "_", texto).strip("_.- ")


def nome_arquivo_perfil(cidade: str, uf: str, mes_ano: str, extensao: str = ".xlsx") -> str:
    """«Perfil_<Cidade>_<UF>_<MêsAno>.xlsx» (partes em branco ficam de fora; espaços viram «_»)."""
    partes = [p for p in (_parte_nome(cidade), _parte_nome(uf).upper(), _parte_nome(mes_ano)) if p]
    return "_".join(["Perfil"] + partes) + extensao


def sugerir_identificacao(nome_banco: str) -> Tuple[str, str]:
    """(UF, MêsAno) tirados do nome do banco quando seguem o padrão do instituto
    («..._RJ_Set26.sav» -> «RJ», «Set26»); vazio quando não dá para ler. A cidade não se deduz."""
    base = str(nome_banco or "").replace("\\", "/").rsplit("/", 1)[-1]
    base = re.sub(r"\.[A-Za-z0-9]{2,4}$", "", base)
    achados = list(_RE_UF_MES.finditer(base))
    if not achados:
        return "", ""
    m = achados[-1]
    return m.group(1), m.group(2)


# ---------------------------------------------------------------- Excel
# Estrutura idêntica ao Perfil que o SPSS exporta (conferido em «Perfil_Teresópolis_RJ_Set26» e
# «Perfil_Tocantins_2_Set26»): título em negrito mesclado em A:B, cabeçalho «Total», moldura média em volta
# de cada bloco, valores com precisão total (a casa decimal é só de exibição).
# Formatação pedida pela equipe: fonte DIN 10 no arquivo todo; coluna A (largura 21) alinhada à esquerda e
# centralizada na vertical; coluna B (largura 8) centralizada nos dois sentidos.
_FONTE = "DIN"
_TAMANHO = 10
_MEDIO = Side(style="medium")
FORMATO_PCT = "0.0"
FORMATO_BASE = "###0"
FORMATO_MEDIA = "0.00"
_ALT_TITULO = 18.0
_ALT_LINHA = 16.0
_LARG_ROTULO = 21.0
_LARG_VALOR = 8.0


def _texto_numero(valor: float, formato: str) -> str:
    if formato == FORMATO_PCT:
        return f"{valor:.1f}".replace(".", ",")
    if formato == FORMATO_MEDIA:
        return f"{valor:.2f}".replace(".", ",")
    return f"{valor:.0f}"


def _escrever_blocos(ws, blocos: Iterable[BlocoPerfil], como_texto: bool = False) -> int:
    """Escreve os blocos empilhados (título, «Total», categorias, Base, Média, linha em branco).
    Devolve a última linha usada."""
    linha = 1

    def valor(v: float, formato: str):
        return _texto_numero(v, formato) if como_texto else v

    def celula(r: int, c: int, v: Any, formato: Optional[str], alinh: str, topo: bool, base: bool):
        cel = ws.cell(row=r, column=c, value=v)
        cel.font = Font(name=_FONTE, size=_TAMANHO)
        cel.alignment = Alignment(horizontal=alinh, vertical="center")
        cel.border = Border(left=_MEDIO, right=_MEDIO, top=_MEDIO if topo else Side(), bottom=_MEDIO if base else Side())
        if formato and not isinstance(v, str):
            cel.number_format = formato

    for b in blocos:
        ws.cell(row=linha, column=1, value=b.titulo)
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=2)
        t = ws.cell(row=linha, column=1)       # estilo depois da mesclagem
        t.font = Font(name=_FONTE, size=_TAMANHO, bold=True)
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[linha].height = _ALT_TITULO
        linha += 1
        celula(linha, 1, " ", None, "left", True, True)
        celula(linha, 2, "Total", None, "center", True, True)
        ws.row_dimensions[linha].height = _ALT_LINHA
        linha += 1
        itens: List[Tuple[str, Any, Optional[str]]] = [
            (rot, valor(p, FORMATO_PCT), FORMATO_PCT) for rot, p in zip(b.categorias, b.pct)
        ]
        itens.append(("Base", valor(float(b.base), FORMATO_BASE), FORMATO_BASE))
        if b.media is not None:
            itens.append(("Média", valor(b.media, FORMATO_MEDIA), FORMATO_MEDIA))
        for k, (rot, v, fmt) in enumerate(itens):
            ultimo = k == len(itens) - 1
            celula(linha, 1, rot, None, "left", k == 0, ultimo)
            celula(linha, 2, v, fmt, "center", k == 0, ultimo)
            ws.row_dimensions[linha].height = _ALT_LINHA
            linha += 1
        linha += 1      # linha em branco entre blocos
    ws.column_dimensions["A"].width = _LARG_ROTULO
    ws.column_dimensions["B"].width = _LARG_VALOR
    return linha - 1


def planilha_perfil(res: ResultadoPerfil, como_texto: bool = False):
    """Planilha (openpyxl) com os blocos, idêntica à aba «Perfil» do Excel.
    `como_texto=True` troca os números pelo texto formatado (o renderizador HTML da tela não aplica
    o formato de número das células)."""
    wb = Workbook()
    ws = wb.active
    ws.title = NOME_ABA
    _escrever_blocos(ws, res.blocos, como_texto=como_texto)
    return ws


def montar_excel_perfil(res: ResultadoPerfil) -> bytes:
    """Excel do Perfil: uma aba «Perfil», no mesmo layout do arquivo que o SPSS exporta."""
    buf = io.BytesIO()
    planilha_perfil(res).parent.save(buf)
    return buf.getvalue()
