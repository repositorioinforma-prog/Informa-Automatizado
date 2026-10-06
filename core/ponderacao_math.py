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


def _grupos_de_componentes(
    fatores: Mapping[str, Mapping[Any, float]],
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> tuple[OrderedDict[str, int], OrderedDict[int, list[str]]]:
    """Normaliza e agrupa as variáveis por componente de peso."""
    atribuicoes = normalizar_componentes_peso(fatores, componentes_por_variavel)
    grupos: OrderedDict[int, list[str]] = OrderedDict()
    for variavel, numero in atribuicoes.items():
        grupos.setdefault(numero, []).append(variavel)
    return atribuicoes, grupos


def _serie_fator_variavel(
    dados: pd.DataFrame,
    variavel: str,
    mapa: Mapping[Any, float],
) -> tuple[pd.Series, pd.Series]:
    """Mapeia uma variável para fatores e identifica valores não configurados.

    Valores ausentes na variável são tratados como ``não aplicáveis``. Já um
    valor preenchido que não exista no mapa de fatores é considerado inválido.
    """
    if variavel not in dados.columns:
        raise ValueError(f"A variável '{variavel}' não existe na base.")

    serie = dados[variavel]
    fator = pd.to_numeric(serie.map(mapa), errors="coerce")
    valor_nao_mapeado = serie.notna() & fator.isna()
    return fator, valor_nao_mapeado


def diagnosticar_componentes(
    dados: pd.DataFrame,
    fatores: Mapping[str, Mapping[Any, float]],
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> pd.DataFrame:
    """Resume a cobertura de cada componente na base.

    Quando várias variáveis compartilham um mesmo ``pesoN``, cada caso pode
    contribuir com uma ou mais delas. Isso permite, por exemplo, reunir no
    mesmo componente variáveis regionais mutuamente exclusivas: ``VN_sud``,
    ``VN_nord`` etc. O componente é válido quando ao menos uma variável do
    grupo fornece um fator e nenhuma variável preenchida possui categoria sem
    fator configurado.
    """
    _, grupos = _grupos_de_componentes(fatores, componentes_por_variavel)
    linhas: list[dict[str, Any]] = []

    for numero, variaveis in sorted(grupos.items()):
        fatores_var: list[pd.Series] = []
        nao_mapeados = pd.Series(False, index=dados.index)
        for variavel in variaveis:
            fator, invalido = _serie_fator_variavel(dados, variavel, fatores[variavel])
            fatores_var.append(fator.rename(variavel))
            nao_mapeados |= invalido

        matriz = pd.concat(fatores_var, axis=1)
        contribuicoes = matriz.notna().sum(axis=1)
        valido = (contribuicoes >= 1) & ~nao_mapeados

        linhas.append(
            {
                "componente": f"peso{numero}",
                "numero": int(numero),
                "variaveis": ", ".join(variaveis),
                "quantidade_variaveis": len(variaveis),
                "casos_total": int(len(dados)),
                "casos_validos": int(valido.sum()),
                "casos_sem_valor_aplicavel": int((contribuicoes == 0).sum()),
                "casos_com_valor_nao_mapeado": int(nao_mapeados.sum()),
                "casos_com_multiplas_contribuicoes": int((contribuicoes > 1).sum()),
            }
        )

    return pd.DataFrame(linhas)


def aplicar_fatores(
    dados: pd.DataFrame,
    fatores: Mapping[str, Mapping[Any, float]],
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> pd.Series:
    """Aplica os fatores respeitando os componentes de peso.

    Regras por componente:
    - se apenas uma variável do grupo tiver valor no caso, usa o fator dela;
    - se duas ou mais tiverem valor, multiplica os fatores disponíveis;
    - valores ausentes são tratados como ``não aplicáveis`` dentro do grupo;
    - se nenhuma variável do componente se aplicar, o peso final fica ausente;
    - uma categoria preenchida sem fator configurado invalida o componente.

    Entre componentes, o peso final continua sendo o produto de ``peso1``,
    ``peso2`` e assim por diante.
    """
    if not fatores:
        raise ValueError("Não há fatores para aplicar.")

    _, grupos = _grupos_de_componentes(fatores, componentes_por_variavel)
    pesos = pd.Series(1.0, index=dados.index, dtype="float64")
    valido_final = pd.Series(True, index=dados.index)

    for _, variaveis in sorted(grupos.items()):
        series_fatores: list[pd.Series] = []
        nao_mapeados = pd.Series(False, index=dados.index)

        for variavel in variaveis:
            fator, invalido = _serie_fator_variavel(dados, variavel, fatores[variavel])
            series_fatores.append(fator.rename(variavel))
            nao_mapeados |= invalido

        matriz = pd.concat(series_fatores, axis=1)
        componente = matriz.prod(axis=1, skipna=True, min_count=1)
        componente.loc[nao_mapeados] = np.nan

        valido_componente = componente.notna() & np.isfinite(componente)
        valido_final &= valido_componente
        pesos = pesos * componente.fillna(1.0)

    pesos.loc[~valido_final] = np.nan
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
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> ResultadoPonderacao:
    """Reproduz o método atual: razão por margem e produto dos fatores."""
    alvos_norm = normalizar_alvos(alvos)
    _validar_categorias_amostra(dados, alvos_norm)
    fatores = _fatores_razao(dados, alvos_norm)
    pesos = aplicar_fatores(
        dados,
        fatores,
        componentes_por_variavel=componentes_por_variavel,
    )
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
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> ResultadoPonderacao:
    """Ajusta iterativamente as margens (raking).

    Variáveis agrupadas no mesmo componente podem ser mutuamente exclusivas.
    Um caso é elegível quando possui ao menos uma categoria válida em cada
    componente. Para cada margem, o denominador do ajuste considera apenas os
    casos em que aquela variável está preenchida, preservando metas condicionais
    como voto por região.
    """
    if max_iteracoes < 1:
        raise ValueError("O número máximo de iterações deve ser pelo menos 1.")

    alvos_norm = normalizar_alvos(alvos)
    _validar_categorias_amostra(dados, alvos_norm)

    _, grupos = _grupos_de_componentes(alvos_norm, componentes_por_variavel)
    completa = pd.Series(True, index=dados.index)

    # Cada componente precisa ter ao menos uma variável aplicável por caso.
    # Categorias preenchidas que não estejam no universo informado invalidam o caso.
    for _, variaveis_grupo in sorted(grupos.items()):
        tem_contribuicao = pd.Series(False, index=dados.index)
        tem_valor_nao_mapeado = pd.Series(False, index=dados.index)

        for variavel in variaveis_grupo:
            serie = dados[variavel]
            categorias = list(alvos_norm[variavel].keys())
            mapeado = serie.notna() & serie.isin(categorias)
            tem_contribuicao |= mapeado
            tem_valor_nao_mapeado |= serie.notna() & ~serie.isin(categorias)

        completa &= tem_contribuicao & ~tem_valor_nao_mapeado

    if not bool(completa.any()):
        raise ValueError(
            "Não há casos com ao menos uma categoria válida em cada componente de peso. "
            "Quando variáveis são alternativas por região, atribua todas ao mesmo peso."
        )

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
            aplicavel = completa & serie.notna() & serie.isin(list(mapa.keys()))
            total = float(pesos.loc[aplicavel].sum())
            if total <= 0:
                raise ValueError(
                    f"A variável '{variavel}' não possui casos válidos após combinar os componentes."
                )

            for categoria, alvo_pct in mapa.items():
                mascara_cat = aplicavel & (serie == categoria)
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


def normalizar_componentes_peso(
    fatores: Mapping[str, Mapping[Any, float]],
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> OrderedDict[str, int]:
    """Normaliza o componente (peso1, peso2, ...) atribuído a cada variável.

    Por padrão, cada variável recebe um componente próprio, preservando o
    comportamento histórico. Quando ``componentes_por_variavel`` é informado,
    várias variáveis podem compartilhar o mesmo componente.
    """
    saida: OrderedDict[str, int] = OrderedDict()
    for indice, variavel in enumerate(fatores.keys(), start=1):
        valor = indice if componentes_por_variavel is None else componentes_por_variavel.get(variavel, indice)
        try:
            numero = int(valor)
        except (TypeError, ValueError):
            raise ValueError(f"Componente de peso inválido para '{variavel}': {valor!r}.") from None
        if numero < 1:
            raise ValueError(f"O componente de peso de '{variavel}' deve ser um inteiro maior ou igual a 1.")
        saida[str(variavel)] = numero
    return saida


def gerar_syntax_spss(
    fatores: Mapping[str, Mapping[Any, float]],
    nome_peso: str = "peso",
    componentes_por_variavel: Mapping[str, int] | None = None,
) -> str:
    """Gera syntax SPSS equivalente aos fatores calculados.

    ``componentes_por_variavel`` permite agrupar várias margens no mesmo
    componente. O ``pesoN`` multiplica os fatores disponíveis no caso e ignora
    variáveis ausentes dentro do grupo, desde que ao menos uma seja aplicável.
    Isso atende tanto variáveis simultâneas quanto blocos mutuamente exclusivos
    por região.
    """
    if not fatores:
        raise ValueError("Não há fatores para gerar a syntax SPSS.")

    atribuicoes = normalizar_componentes_peso(fatores, componentes_por_variavel)
    grupos: OrderedDict[int, list[str]] = OrderedDict()
    for variavel, numero in atribuicoes.items():
        grupos.setdefault(numero, []).append(variavel)

    linhas = ["* Syntax gerada automaticamente pelo módulo de Ponderação.", ""]
    componentes = []

    for numero in sorted(grupos):
        componente = f"peso{numero}"
        componentes.append(componente)
        variaveis_grupo = grupos[numero]
        linhas.append(
            f"*************** {componente}: " + ", ".join(variaveis_grupo) + "."
        )

        if len(variaveis_grupo) == 1:
            variavel = variaveis_grupo[0]
            mapa = fatores[variavel]
            linhas.append(f"COMPUTE {componente}=$SYSMIS.")
            for categoria, fator in mapa.items():
                linhas.append(
                    f"IF ({variavel}={_literal_spss(categoria)}) {componente}={format(float(fator), '.15g')}."
                )
            linhas.append("")
            continue

        # Variáveis no mesmo componente podem ser alternativas (por exemplo,
        # uma variável de voto para cada região) ou simultâneas. O componente
        # multiplica apenas os fatores disponíveis no caso. Variáveis ausentes
        # não anulam o grupo; ao menos uma precisa contribuir.
        ok = f"#pndok{numero}"
        erro = f"#pnderr{numero}"
        fator_tmp = f"#pndf{numero}"
        linhas.append(f"COMPUTE {componente}=$SYSMIS.")
        linhas.append(f"COMPUTE {ok}=0.")
        linhas.append(f"COMPUTE {erro}=0.")

        for variavel in variaveis_grupo:
            mapa = fatores[variavel]
            linhas.append(f"* {variavel} -> {componente}.")
            linhas.append(f"COMPUTE {fator_tmp}=$SYSMIS.")
            for categoria, fator in mapa.items():
                linhas.append(
                    f"IF ({variavel}={_literal_spss(categoria)}) "
                    f"{fator_tmp}={format(float(fator), '.15g')}."
                )
            linhas.append(f"DO IF (NOT MISSING({fator_tmp})).")
            linhas.append(f"  DO IF ({ok}=0).")
            linhas.append(f"    COMPUTE {componente}={fator_tmp}.")
            linhas.append("  ELSE.")
            linhas.append(f"    COMPUTE {componente}={componente}*{fator_tmp}.")
            linhas.append("  END IF.")
            linhas.append(f"  COMPUTE {ok}=1.")
            linhas.append(f"ELSE IF (NOT MISSING({variavel})).")
            linhas.append(f"  COMPUTE {erro}=1.")
            linhas.append("END IF.")

        linhas.append(f"IF ({ok}=0 OR {erro}=1) {componente}=$SYSMIS.")
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
