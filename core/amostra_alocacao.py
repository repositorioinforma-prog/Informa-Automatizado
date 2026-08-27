"""Ferramentas de auditoria e aplicação de alocação manual da amostra.

O motor principal continua produzindo uma sugestão proporcional. Este módulo
permite substituir a quantidade final de cada unidade por valores informados
pelo usuário, desde que todas as unidades estejam presentes, os valores sejam
inteiros não negativos e a soma feche exatamente na amostra total.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import Any

import pandas as pd

from core.amostra_math import (
    adicionar_cotas_demograficas,
    adicionar_memoria_calculo,
    remover_cotas_calculadas,
)

CODIGO_POR_NIVEL = {
    "uf": "uf",
    "regiao_intermediaria": "cod_regiao_intermediaria",
    "regiao_imediata": "cod_regiao_imediata",
    "municipio": "cod_municipio",
    "distrito": "cod_distrito",
}

NOME_POR_NIVEL = {
    "uf": "nome_uf",
    "regiao_intermediaria": "regiao_intermediaria",
    "regiao_imediata": "regiao_imediata",
    "municipio": "municipio",
    "distrito": "distrito",
}

HIERARQUIA_NOMES = [
    "uf",
    "regiao_intermediaria",
    "regiao_imediata",
    "municipio",
    "distrito",
]


class AlocacaoManualErro(ValueError):
    """A distribuição manual é inválida ou não fecha na meta definida."""


def coluna_codigo_nivel(nivel: str) -> str:
    try:
        return CODIGO_POR_NIVEL[nivel]
    except KeyError as exc:
        raise AlocacaoManualErro(f"Nível territorial inválido: {nivel!r}.") from exc


def coluna_nome_nivel(nivel: str) -> str:
    try:
        return NOME_POR_NIVEL[nivel]
    except KeyError as exc:
        raise AlocacaoManualErro(f"Nível territorial inválido: {nivel!r}.") from exc


def _inteiro_nao_negativo(valor: Any, campo: str) -> int:
    if isinstance(valor, bool):
        raise AlocacaoManualErro(f"{campo} precisa ser um número inteiro.")
    try:
        numero = float(valor)
    except (TypeError, ValueError) as exc:
        raise AlocacaoManualErro(f"{campo} precisa ser um número inteiro.") from exc
    if pd.isna(numero) or not numero.is_integer():
        raise AlocacaoManualErro(f"{campo} precisa ser um número inteiro.")
    inteiro = int(numero)
    if inteiro < 0:
        raise AlocacaoManualErro(f"{campo} não pode ser negativo.")
    return inteiro


def normalizar_alocacoes_manuais(
    alocacoes: Mapping[str, Any] | Sequence[Mapping[str, Any]],
) -> dict[str, int]:
    """Converte dicionário ou lista JSON em ``{codigo: amostra}``."""
    if isinstance(alocacoes, Mapping):
        itens = alocacoes.items()
    elif isinstance(alocacoes, Sequence) and not isinstance(alocacoes, (str, bytes)):
        pares = []
        for indice, item in enumerate(alocacoes, start=1):
            if not isinstance(item, Mapping):
                raise AlocacaoManualErro(
                    f"A alocação manual {indice} precisa ser um objeto com código e amostra."
                )
            pares.append((item.get("codigo"), item.get("amostra")))
        itens = pares
    else:
        raise AlocacaoManualErro("As alocações manuais precisam ser uma lista ou objeto.")

    normalizadas: dict[str, int] = {}
    for codigo_bruto, valor in itens:
        codigo = str(codigo_bruto or "").strip()
        if not codigo:
            raise AlocacaoManualErro("Toda alocação manual precisa possuir um código.")
        if codigo in normalizadas:
            raise AlocacaoManualErro(f"Código repetido na alocação manual: {codigo}.")
        normalizadas[codigo] = _inteiro_nao_negativo(
            valor,
            f"Amostra da unidade {codigo}",
        )
    return normalizadas


def validar_alocacoes_manuais(
    alocacoes: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    codigos_esperados: Sequence[Any],
    amostra_total: int,
) -> dict[str, int]:
    """Valida cobertura integral das unidades e fechamento da meta."""
    meta = _inteiro_nao_negativo(amostra_total, "amostra_total")
    if meta <= 0:
        raise AlocacaoManualErro("amostra_total precisa ser maior que zero.")

    normalizadas = normalizar_alocacoes_manuais(alocacoes)
    esperados = [str(codigo) for codigo in codigos_esperados]
    conjunto_esperado = set(esperados)
    conjunto_recebido = set(normalizadas)

    faltantes = sorted(conjunto_esperado - conjunto_recebido)
    extras = sorted(conjunto_recebido - conjunto_esperado)
    if faltantes or extras:
        partes = []
        if faltantes:
            partes.append(f"{len(faltantes)} unidade(s) sem valor")
        if extras:
            partes.append(f"{len(extras)} código(s) fora do recorte")
        raise AlocacaoManualErro(
            "A distribuição manual não corresponde ao recorte atual: " + "; ".join(partes) + "."
        )

    soma = sum(normalizadas.values())
    if soma != meta:
        diferenca = meta - soma
        if diferenca > 0:
            detalhe = f"faltam {diferenca} entrevista(s)"
        else:
            detalhe = f"há {-diferenca} entrevista(s) em excesso"
        raise AlocacaoManualErro(
            f"A distribuição manual soma {soma}, mas a meta é {meta}: {detalhe}."
        )
    return normalizadas


def preparar_tabela_alocacao_manual(
    resultado_proporcional: pd.DataFrame,
    nivel: str,
    alocacoes_iniciais: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
) -> pd.DataFrame:
    """Cria a tabela amigável usada pelo editor da interface."""
    codigo = coluna_codigo_nivel(nivel)
    if codigo not in resultado_proporcional.columns:
        raise AlocacaoManualErro(
            f"O resultado não contém o identificador necessário para {nivel}: {codigo}."
        )

    df = adicionar_memoria_calculo(
        resultado_proporcional,
        int(pd.to_numeric(resultado_proporcional["amostra"], errors="coerce").fillna(0).sum()),
    )
    colunas_hierarquia = [col for col in HIERARQUIA_NOMES if col in df.columns]
    colunas = colunas_hierarquia + [
        codigo,
        "populacao_total",
        "populacao_base",
        "percentual",
        "amostra_ideal",
        "amostra_proporcional",
    ]
    tabela = df[colunas].copy()
    tabela[codigo] = tabela[codigo].astype(str)

    iniciais: dict[str, int] = {}
    if alocacoes_iniciais:
        try:
            iniciais = normalizar_alocacoes_manuais(alocacoes_iniciais)
        except AlocacaoManualErro:
            iniciais = {}

    # Só reaproveita valores salvos quando eles pertencem exatamente ao recorte
    # atual e fecham na mesma meta. Isso evita que uma edição de outro projeto
    # seja aplicada parcialmente depois de o usuário trocar território ou N.
    codigos_atuais = set(tabela[codigo].astype(str))
    meta_atual = int(
        pd.to_numeric(resultado_proporcional["amostra"], errors="coerce")
        .fillna(0)
        .sum()
    )
    if set(iniciais) != codigos_atuais or sum(iniciais.values()) != meta_atual:
        iniciais = {}

    tabela["amostra_manual"] = [
        iniciais.get(str(cod), int(sugerida))
        for cod, sugerida in zip(tabela[codigo], tabela["amostra_proporcional"])
    ]
    return tabela


def aplicar_alocacao_manual(
    resultado_proporcional: pd.DataFrame,
    *,
    nivel: str,
    alocacoes_manuais: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    amostra_total: int,
    faixa_flex: float = 0.20,
) -> pd.DataFrame:
    """Substitui a alocação proporcional e recalcula bandas e cotas internas."""
    if faixa_flex < 0:
        raise AlocacaoManualErro("faixa_flex não pode ser negativa.")

    codigo = coluna_codigo_nivel(nivel)
    if codigo not in resultado_proporcional.columns:
        raise AlocacaoManualErro(
            f"O resultado não contém a coluna de código da unidade final: {codigo}."
        )

    resultado = adicionar_memoria_calculo(resultado_proporcional, int(amostra_total))
    codigos = resultado[codigo].astype(str).tolist()
    alocacoes = validar_alocacoes_manuais(alocacoes_manuais, codigos, int(amostra_total))

    resultado = resultado.copy()
    resultado[codigo] = resultado[codigo].astype(str)
    resultado["amostra"] = resultado[codigo].map(alocacoes).astype(int)
    resultado["minimo"] = resultado["amostra"].apply(
        lambda n: max(0, round(int(n) * (1 - faixa_flex)))
    )
    resultado["maximo"] = resultado["amostra"].apply(
        lambda n: round(int(n) * (1 + faixa_flex))
    )

    resultado = remover_cotas_calculadas(resultado)
    resultado = adicionar_cotas_demograficas(resultado)
    resultado = adicionar_memoria_calculo(resultado, int(amostra_total))
    resultado["metodo_distribuicao"] = "manual"
    return resultado


def alocacoes_para_json(
    tabela_editada: pd.DataFrame,
    *,
    nivel: str,
    coluna_amostra: str = "amostra_manual",
) -> list[dict[str, int | str]]:
    """Extrai do editor uma lista portátil de códigos e quantidades."""
    codigo = coluna_codigo_nivel(nivel)
    if codigo not in tabela_editada.columns or coluna_amostra not in tabela_editada.columns:
        raise AlocacaoManualErro("A tabela editada não possui código e amostra manual.")
    saida = []
    for _, row in tabela_editada.iterrows():
        saida.append({
            "codigo": str(row[codigo]),
            "amostra": _inteiro_nao_negativo(row[coluna_amostra], "Meta da unidade"),
        })
    return saida
