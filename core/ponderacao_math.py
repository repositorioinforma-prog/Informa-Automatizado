"""Cálculos puros do módulo de ponderação.

O módulo mantém dois modos:
- razão simples: reproduz o processo atual (universo / amostra por categoria
  e multiplicação dos fatores das variáveis);
- raking: recalibra iterativamente as margens para tentar aproximá-las dos
  percentuais de universo dentro da tolerância informada.
"""
from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from typing import Any, Mapping

import numpy as np
import pandas as pd


@dataclass
class ResultadoPonderacao:
    pesos: pd.Series
    fatores: dict[str, OrderedDict]
    diagnostico: pd.DataFrame
    convergiu: bool
    iteracoes: int
    maior_diferenca_pp: float
    metodo: str


def categorias_validas(serie: pd.Series) -> list[Any]:
    """Retorna categorias não nulas na ordem em que aparecem na base."""
    return list(pd.unique(serie.dropna()))


def normalizar_alvos(alvos: Mapping[str, Mapping[Any, float]]) -> dict[str, OrderedDict]:
    """Valida e normaliza cada margem para totalizar 100%."""
    saida: dict[str, OrderedDict] = {}

    if not alvos:
        raise ValueError("Selecione ao menos uma variável para ponderação.")

    for variavel, mapa in alvos.items():
        if not mapa:
            raise ValueError(f"A variável '{variavel}' não possui metas de universo.")

        valores = OrderedDict()
        total = 0.0
        for categoria, valor in mapa.items():
            try:
                numero = float(valor)
            except (TypeError, ValueError):
                raise ValueError(
                    f"Percentual de universo inválido em '{variavel}', categoria '{categoria}'."
                ) from None

            if not np.isfinite(numero) or numero < 0:
                raise ValueError(
                    f"O percentual de universo deve ser >= 0 em '{variavel}', categoria '{categoria}'."
                )
            valores[categoria] = numero
            total += numero

        if total <= 0:
            raise ValueError(f"Os percentuais de universo de '{variavel}' somam zero.")

        saida[variavel] = OrderedDict(
            (categoria, valor * 100.0 / total) for categoria, valor in valores.items()
        )

    return saida


def frequencia_percentual(
    serie: pd.Series,
    pesos: pd.Series | None = None,
    categorias: list[Any] | None = None,
) -> pd.DataFrame:
    """Calcula frequência e percentual válido, com ou sem peso."""
    if categorias is None:
        categorias = categorias_validas(serie)

    linhas = []
    mascara_valida = serie.notna()

    if pesos is None:
        denominador = float(mascara_valida.sum())
        for categoria in categorias:
            freq = float((serie[mascara_valida] == categoria).sum())
            pct = (freq / denominador * 100.0) if denominador else np.nan
            linhas.append((categoria, freq, pct))
    else:
        pesos_num = pd.to_numeric(pesos, errors="coerce")
        mascara = mascara_valida & pesos_num.notna()
        denominador = float(pesos_num[mascara].sum())
        for categoria in categorias:
            cat_mask = mascara & (serie == categoria)
            freq = float(pesos_num[cat_mask].sum())
            pct = (freq / denominador * 100.0) if denominador else np.nan
            linhas.append((categoria, freq, pct))

    return pd.DataFrame(linhas, columns=["codigo", "frequencia", "percentual"])


def _validar_categorias_amostra(dados: pd.DataFrame, alvos: Mapping[str, Mapping[Any, float]]) -> None:
    for variavel, mapa in alvos.items():
        if variavel not in dados.columns:
            raise ValueError(f"A variável '{variavel}' não existe na base.")

        serie = dados[variavel]
        for categoria, alvo in mapa.items():
            existe = bool((serie == categoria).any())
            if alvo > 0 and not existe:
                raise ValueError(
                    f"Não é possível atingir {alvo:.2f}% em '{variavel}' = '{categoria}': "
                    "a categoria não possui casos na amostra."
                )


def _fatores_razao(dados: pd.DataFrame, alvos: Mapping[str, Mapping[Any, float]]) -> dict[str, OrderedDict]:
    fatores: dict[str, OrderedDict] = {}

    for variavel, mapa in alvos.items():
        serie = dados[variavel]
        freq = frequencia_percentual(serie, categorias=list(mapa.keys()))
        pct_amostra = dict(zip(freq["codigo"], freq["percentual"]))
        fatores_var = OrderedDict()

        for categoria, alvo in mapa.items():
            atual = float(pct_amostra.get(categoria, 0.0) or 0.0)
            if atual <= 0:
                if alvo <= 0:
                    fator = 0.0
                else:
                    raise ValueError(
                        f"Não há casos em '{variavel}' = '{categoria}' para calcular o peso."
                    )
            else:
                fator = float(alvo) / atual
            fatores_var[categoria] = fator

        fatores[variavel] = fatores_var

    return fatores


def aplicar_fatores(dados: pd.DataFrame, fatores: Mapping[str, Mapping[Any, float]]) -> pd.Series:
    """Aplica o produto dos fatores por variável; casos sem fator ficam sem peso."""
    pesos = pd.Series(1.0, index=dados.index, dtype="float64")
    valido = pd.Series(True, index=dados.index)

    for variavel, mapa in fatores.items():
        componente = pd.to_numeric(dados[variavel].map(mapa), errors="coerce")
        valido &= componente.notna()
        pesos = pesos * componente.fillna(1.0)

    pesos.loc[~valido] = np.nan
    return pesos


def diagnosticar(
    dados: pd.DataFrame,
    pesos: pd.Series,
    alvos: Mapping[str, Mapping[Any, float]],
    fatores: Mapping[str, Mapping[Any, float]],
    labels: Mapping[str, Mapping[Any, str]] | None = None,
    tolerancia_pp: float = 2.0,
) -> pd.DataFrame:
    """Monta a tabela de conferência antes/depois para todas as margens."""
    labels = labels or {}
    linhas: list[dict[str, Any]] = []

    for variavel, mapa in alvos.items():
        cats = list(mapa.keys())
        sem = frequencia_percentual(dados[variavel], categorias=cats).set_index("codigo")
        com = frequencia_percentual(dados[variavel], pesos=pesos, categorias=cats).set_index("codigo")

        for categoria, universo in mapa.items():
            sem_freq = float(sem.at[categoria, "frequencia"]) if categoria in sem.index else 0.0
            sem_pct = float(sem.at[categoria, "percentual"]) if categoria in sem.index else np.nan
            com_freq = float(com.at[categoria, "frequencia"]) if categoria in com.index else 0.0
            com_pct = float(com.at[categoria, "percentual"]) if categoria in com.index else np.nan
            diferenca = com_pct - float(universo) if np.isfinite(com_pct) else np.nan
            abs_diff = abs(diferenca) if np.isfinite(diferenca) else np.nan

            linhas.append(
                {
                    "variavel": variavel,
                    "codigo": categoria,
                    "categoria": labels.get(variavel, {}).get(categoria, str(categoria)),
                    "frequencia_sem_peso": sem_freq,
                    "percentual_sem_peso": sem_pct,
                    "percentual_universo": float(universo),
                    "fator": float(fatores.get(variavel, {}).get(categoria, np.nan)),
                    "frequencia_com_peso": com_freq,
                    "percentual_com_peso": com_pct,
                    "diferenca_pp": diferenca,
                    "diferenca_abs_pp": abs_diff,
                    "status": "OK" if np.isfinite(abs_diff) and abs_diff <= tolerancia_pp else "REVISAR",
                }
            )

    return pd.DataFrame(linhas)


def calcular_razao_simples(
    dados: pd.DataFrame,
    alvos: Mapping[str, Mapping[Any, float]],
    labels: Mapping[str, Mapping[Any, str]] | None = None,
    tolerancia_pp: float = 2.0,
) -> ResultadoPonderacao:
    """Reproduz o método atual: razão por margem e produto dos fatores."""
    alvos_norm = normalizar_alvos(alvos)
    _validar_categorias_amostra(dados, alvos_norm)
    fatores = _fatores_razao(dados, alvos_norm)
    pesos = aplicar_fatores(dados, fatores)
    diag = diagnosticar(dados, pesos, alvos_norm, fatores, labels, tolerancia_pp)
    maior = float(diag["diferenca_abs_pp"].max()) if not diag.empty else np.nan

    return ResultadoPonderacao(
        pesos=pesos,
        fatores=fatores,
        diagnostico=diag,
        convergiu=bool(np.isfinite(maior) and maior <= tolerancia_pp),
        iteracoes=1,
        maior_diferenca_pp=maior,
        metodo="Razão simples",
    )


def calcular_raking(
    dados: pd.DataFrame,
    alvos: Mapping[str, Mapping[Any, float]],
    labels: Mapping[str, Mapping[Any, str]] | None = None,
    tolerancia_pp: float = 2.0,
    max_iteracoes: int = 100,
) -> ResultadoPonderacao:
    """Ajusta iterativamente as margens (raking) até a tolerância ou o limite."""
    if max_iteracoes < 1:
        raise ValueError("O número máximo de iterações deve ser pelo menos 1.")

    alvos_norm = normalizar_alvos(alvos)
    _validar_categorias_amostra(dados, alvos_norm)

    variaveis = list(alvos_norm.keys())
    completa = dados[variaveis].notna().all(axis=1)

    # Também exige que cada valor observado tenha um alvo configurado.
    for variavel, mapa in alvos_norm.items():
        completa &= dados[variavel].isin(list(mapa.keys()))

    if not bool(completa.any()):
        raise ValueError("Não há casos completos para as variáveis selecionadas.")

    pesos = pd.Series(np.nan, index=dados.index, dtype="float64")
    pesos.loc[completa] = 1.0

    fatores: dict[str, OrderedDict] = {
        var: OrderedDict((cat, 1.0) for cat in mapa.keys())
        for var, mapa in alvos_norm.items()
    }

    convergiu = False
    maior = np.inf
    iteracao = 0

    for iteracao in range(1, max_iteracoes + 1):
        for variavel, mapa in alvos_norm.items():
            serie = dados[variavel]
            total = float(pesos.loc[completa].sum())
            if total <= 0:
                raise ValueError("A soma dos pesos ficou igual ou menor que zero.")

            for categoria, alvo_pct in mapa.items():
                mascara_cat = completa & (serie == categoria)
                atual = float(pesos.loc[mascara_cat].sum())
                desejado = total * float(alvo_pct) / 100.0

                if atual <= 0:
                    if desejado <= 0:
                        fator = 1.0
                    else:
                        raise ValueError(
                            f"A categoria '{variavel}' = '{categoria}' ficou sem peso durante o ajuste."
                        )
                else:
                    fator = desejado / atual

                pesos.loc[mascara_cat] *= fator
                fatores[variavel][categoria] *= fator

        diag_iter = diagnosticar(
            dados, pesos, alvos_norm, fatores, labels=labels, tolerancia_pp=tolerancia_pp
        )
        maior = float(diag_iter["diferenca_abs_pp"].max())
        if np.isfinite(maior) and maior <= tolerancia_pp:
            convergiu = True
            break

    diag = diagnosticar(
        dados, pesos, alvos_norm, fatores, labels=labels, tolerancia_pp=tolerancia_pp
    )
    maior = float(diag["diferenca_abs_pp"].max()) if not diag.empty else np.nan

    return ResultadoPonderacao(
        pesos=pesos,
        fatores=fatores,
        diagnostico=diag,
        convergiu=convergiu,
        iteracoes=iteracao,
        maior_diferenca_pp=maior,
        metodo="Raking iterativo",
    )


def resumo_pesos(pesos: pd.Series) -> dict[str, float]:
    validos = pd.to_numeric(pesos, errors="coerce").dropna()
    if validos.empty:
        return {
            "casos_com_peso": 0,
            "soma_pesos": np.nan,
            "media": np.nan,
            "minimo": np.nan,
            "maximo": np.nan,
            "cv": np.nan,
            "efeito_desenho_aprox": np.nan,
            "n_efetivo_aprox": np.nan,
        }

    media = float(validos.mean())
    desvio = float(validos.std(ddof=0))
    cv = desvio / media if media else np.nan
    deff = 1.0 + cv**2 if np.isfinite(cv) else np.nan
    n_eff = (float(validos.sum()) ** 2 / float((validos**2).sum())) if float((validos**2).sum()) else np.nan

    return {
        "casos_com_peso": int(validos.shape[0]),
        "soma_pesos": float(validos.sum()),
        "media": media,
        "minimo": float(validos.min()),
        "maximo": float(validos.max()),
        "cv": float(cv),
        "efeito_desenho_aprox": float(deff),
        "n_efetivo_aprox": float(n_eff),
    }


def _literal_spss(valor: Any) -> str:
    if isinstance(valor, (str, np.str_)):
        texto = str(valor).replace("'", "''")
        return f"'{texto}'"
    if isinstance(valor, (int, np.integer)):
        return str(int(valor))
    if isinstance(valor, (float, np.floating)) and float(valor).is_integer():
        return str(int(valor))
    if isinstance(valor, (float, np.floating)):
        return format(float(valor), ".15g")
    return str(valor)


def gerar_syntax_spss(
    fatores: Mapping[str, Mapping[Any, float]],
    nome_peso: str = "peso",
) -> str:
    """Gera syntax SPSS equivalente aos fatores calculados."""
    linhas = ["* Syntax gerada automaticamente pelo módulo de Ponderação.", ""]
    componentes = []

    for indice, (variavel, mapa) in enumerate(fatores.items(), start=1):
        componente = f"peso{indice}"
        componentes.append(componente)
        linhas.append(f"*************** {variavel}.")
        linhas.append(f"COMPUTE {componente}=$SYSMIS.")
        for categoria, fator in mapa.items():
            linhas.append(
                f"IF ({variavel}={_literal_spss(categoria)}) {componente}={format(float(fator), '.15g')}."
            )
        linhas.append("")

    produto = "*".join(componentes)
    linhas.extend(
        [
            "******* Ponderação Final **********.",
            f"COMPUTE {nome_peso}={produto}.",
            "EXECUTE.",
            f"WEIGHT BY {nome_peso}.",
            "EXECUTE.",
            "",
        ]
    )
    return "\n".join(linhas)
