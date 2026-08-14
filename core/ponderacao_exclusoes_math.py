"""Heurística de apoio a exclusões para melhorar margens de ponderação.

O objetivo não é provar uma solução ótima. A rotina procura, de forma gulosa e
rastreável, perfis de entrevistas cuja remoção aproxima as distribuições brutas
das margens do Universo, respeitando um N mínimo informado pelo usuário.

A ponderação deve ser recalculada após aplicar as exclusões sugeridas. O módulo
retorna diagnóstico antes/depois e nunca remove casos por conta própria.
"""
from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Mapping

import numpy as np
import pandas as pd


@dataclass
class ResultadoSugestaoExclusoes:
    ids_excluir: list[Any]
    indices_excluir: list[Any]
    resumo_margens: pd.DataFrame
    sobras_remanescentes: pd.DataFrame
    n_inicial: int
    n_final: int
    n_minimo: int
    limite_exclusoes: int
    atingiu_tolerancia_bruta: bool
    atingiu_limite_exclusoes: bool
    maior_diferenca_bruta_antes_pp: float
    maior_diferenca_bruta_depois_pp: float
    motivo_parada: str
    base_minima_categoria: int = 0
    atingiu_n_alvo: bool = False
    faltam_exclusoes_para_alvo: int = 0
    restricoes_ativas: pd.DataFrame = field(default_factory=pd.DataFrame)


def _normalizar_alvos_sem_renomear(
    alvos: Mapping[str, Mapping[Any, float]],
) -> dict[str, OrderedDict]:
    saida: dict[str, OrderedDict] = {}
    for variavel, mapa in alvos.items():
        if not mapa:
            raise ValueError(f"A variável '{variavel}' não possui metas de Universo.")
        total = 0.0
        valores = OrderedDict()
        for categoria, valor in mapa.items():
            numero = float(valor)
            if not np.isfinite(numero) or numero < 0:
                raise ValueError(
                    f"Percentual inválido em '{variavel}', categoria '{categoria}'."
                )
            valores[categoria] = numero
            total += numero
        if total <= 0:
            raise ValueError(f"Os percentuais de '{variavel}' somam zero.")
        saida[variavel] = OrderedDict(
            (categoria, valor * 100.0 / total) for categoria, valor in valores.items()
        )
    return saida


def _estado_margens(
    dados: pd.DataFrame,
    ativos: pd.Series,
    alvos: Mapping[str, Mapping[Any, float]],
) -> tuple[dict[str, dict[Any, int]], dict[str, int], pd.DataFrame]:
    contagens: dict[str, dict[Any, int]] = {}
    denominadores: dict[str, int] = {}
    linhas: list[dict[str, Any]] = []

    for variavel, mapa in alvos.items():
        serie = dados.loc[ativos, variavel]
        validos = serie.notna()
        denominador = int(validos.sum())
        denominadores[variavel] = denominador
        counts = serie[validos].value_counts(dropna=True).to_dict()
        contagens[variavel] = {cat: int(counts.get(cat, 0)) for cat in mapa}

        for categoria, universo in mapa.items():
            n = int(contagens[variavel].get(categoria, 0))
            pct = (n / denominador * 100.0) if denominador else np.nan
            diff = pct - float(universo) if np.isfinite(pct) else np.nan
            linhas.append(
                {
                    "variavel": variavel,
                    "codigo": categoria,
                    "n": n,
                    "percentual_amostra": pct,
                    "percentual_universo": float(universo),
                    "diferenca_pp": diff,
                    "diferenca_abs_pp": abs(diff) if np.isfinite(diff) else np.nan,
                }
            )

    return contagens, denominadores, pd.DataFrame(linhas)


def _objetivo_diagnostico(diag: pd.DataFrame, tolerancia_pp: float) -> tuple[float, float, float]:
    if diag.empty:
        return (np.inf, np.inf, np.inf)
    diffs = pd.to_numeric(diag["diferenca_abs_pp"], errors="coerce").dropna()
    if diffs.empty:
        return (np.inf, np.inf, np.inf)
    excesso = np.maximum(diffs.to_numpy(dtype=float) - float(tolerancia_pp), 0.0)
    return (
        float(excesso.max()) if excesso.size else 0.0,
        float(excesso.sum()),
        float((diffs.to_numpy(dtype=float) ** 2).sum()),
    )


def _diagnostico_apos_remover_perfil_contagens(
    contagens: Mapping[str, Mapping[Any, int]],
    denominadores: Mapping[str, int],
    perfil: tuple[Any, ...],
    variaveis: list[str],
    alvos: Mapping[str, Mapping[Any, float]],
) -> pd.DataFrame:
    """Simula uma remoção usando somente contagens agregadas.

    Evita reler a base inteira para cada perfil candidato e deixa a busca
    viável mesmo quando o N-alvo exige centenas de exclusões.
    """
    valor_perfil = dict(zip(variaveis, perfil))
    linhas: list[dict[str, Any]] = []
    for variavel, mapa in alvos.items():
        removido = valor_perfil.get(variavel)
        remove_valido = not pd.isna(removido)
        denom = int(denominadores.get(variavel, 0)) - (1 if remove_valido else 0)
        for categoria, universo in mapa.items():
            n = int(contagens.get(variavel, {}).get(categoria, 0))
            if remove_valido and removido == categoria:
                n -= 1
            pct = (n / denom * 100.0) if denom > 0 else np.nan
            diff = pct - float(universo) if np.isfinite(pct) else np.nan
            linhas.append(
                {
                    "variavel": variavel,
                    "codigo": categoria,
                    "n": n,
                    "percentual_amostra": pct,
                    "percentual_universo": float(universo),
                    "diferenca_pp": diff,
                    "diferenca_abs_pp": abs(diff) if np.isfinite(diff) else np.nan,
                }
            )
    return pd.DataFrame(linhas)


def _perfis_candidatos(
    dados: pd.DataFrame,
    ativos: pd.Series,
    variaveis: list[str],
    diag: pd.DataFrame,
    contagens: Mapping[str, Mapping[Any, int]],
    alvos: Mapping[str, Mapping[Any, float]],
    base_minima_categoria: int = 0,
    codigos_protegidos: Mapping[str, set[Any]] | None = None,
    permitir_sem_excesso: bool = False,
    limite: int = 40,
) -> list[tuple[Any, ...]]:
    """Prioriza perfis em excesso sem violar proteções de cota.

    Um perfil fica inelegível quando pertence a uma categoria marcada como
    "nunca excluir" ou quando sua retirada faria alguma categoria usada na
    ponderação cair abaixo da base mínima configurada.
    """
    excesso_por_chave = {
        (row["variavel"], row["codigo"]): max(0.0, float(row["diferenca_pp"]))
        for _, row in diag.iterrows()
        if np.isfinite(row["diferenca_pp"])
    }
    deficit_por_chave = {
        (row["variavel"], row["codigo"]): max(0.0, -float(row["diferenca_pp"]))
        for _, row in diag.iterrows()
        if np.isfinite(row["diferenca_pp"])
    }
    protegidos = codigos_protegidos or {}
    minimo = max(0, int(base_minima_categoria))

    perfis = dados.loc[ativos, variaveis].copy()
    grupos = perfis.groupby(variaveis, dropna=False, sort=False).size().reset_index(name="_n")
    scored: list[tuple[float, tuple[Any, ...]]] = []
    for _, row in grupos.iterrows():
        perfil = tuple(row[v] for v in variaveis)

        bloqueado = False
        for variavel, categoria in zip(variaveis, perfil):
            if pd.isna(categoria):
                continue
            if categoria in protegidos.get(variavel, set()):
                bloqueado = True
                break
            if categoria in alvos.get(variavel, {}) and minimo > 0:
                n_atual = int(contagens.get(variavel, {}).get(categoria, 0))
                if n_atual - 1 < minimo:
                    bloqueado = True
                    break
        if bloqueado:
            continue

        bonus = 0.0
        penalidade = 0.0
        for variavel, categoria in zip(variaveis, perfil):
            if pd.isna(categoria):
                continue
            bonus += excesso_por_chave.get((variavel, categoria), 0.0)
            penalidade += deficit_por_chave.get((variavel, categoria), 0.0)
        score = bonus - 1.5 * penalidade
        if bonus > 0 or permitir_sem_excesso:
            scored.append((score, perfil))

    scored.sort(key=lambda x: x[0], reverse=True)
    return [perfil for _, perfil in scored[: max(1, int(limite))]]


def _escolher_indice_do_perfil(
    dados: pd.DataFrame,
    ativos: pd.Series,
    variaveis: list[str],
    perfil: tuple[Any, ...],
    id_var: str,
):
    mascara = ativos.copy()
    for variavel, valor in zip(variaveis, perfil):
        if pd.isna(valor):
            mascara &= dados[variavel].isna()
        else:
            mascara &= dados[variavel].eq(valor)
    candidatos = dados.loc[mascara, [id_var]].copy()
    if candidatos.empty:
        return None
    # Ordem estável/reprodutível: tenta ordenar pelo ID; se tipos mistos impedirem, usa índice.
    try:
        candidatos = candidatos.sort_values(id_var, kind="stable")
    except Exception:
        pass
    return candidatos.index[0]


def sugerir_exclusoes_ponderacao(
    dados: pd.DataFrame,
    id_var: str,
    alvos: Mapping[str, Mapping[Any, float]],
    n_minimo: int,
    tolerancia_pp: float = 2.0,
    labels: Mapping[str, Mapping[Any, str]] | None = None,
    max_perfis_avaliados: int = 40,
    parar_ao_atingir_tolerancia_bruta: bool = True,
    base_minima_categoria: int = 0,
    codigos_protegidos: Mapping[str, set[Any]] | None = None,
    forcar_atingir_n_minimo: bool = False,
) -> ResultadoSugestaoExclusoes:
    """Sugere exclusões para aproximar as margens brutas ao Universo.

    A busca é gulosa: a cada passo avalia perfis de respondentes em categorias
    com excesso e remove apenas um caso do perfil que mais melhora o conjunto
    de margens. O processo para quando entra na tolerância, quando nenhuma
    remoção melhora o objetivo ou quando o N mínimo é alcançado.
    """
    if id_var not in dados.columns:
        raise ValueError(f"A variável de ID '{id_var}' não existe na base.")
    if dados[id_var].isna().any():
        raise ValueError("A variável de ID possui valores vazios. Corrija antes de gerar exclusões.")
    if dados[id_var].duplicated().any():
        raise ValueError(
            "A variável de ID possui valores duplicados. Use um identificador único para evitar excluir casos indevidos."
        )

    alvos_norm = _normalizar_alvos_sem_renomear(alvos)
    variaveis = list(alvos_norm.keys())
    base_minima_categoria = max(0, int(base_minima_categoria))
    codigos_protegidos = {
        variavel: set(valores or set())
        for variavel, valores in (codigos_protegidos or {}).items()
    }
    for variavel, mapa in alvos_norm.items():
        if variavel not in dados.columns:
            raise ValueError(f"A variável '{variavel}' não existe na base.")
        for categoria, pct in mapa.items():
            if pct > 0 and not bool(dados[variavel].eq(categoria).any()):
                raise ValueError(
                    f"'{variavel}' = '{categoria}' tem Universo {pct:.2f}% mas não possui casos na amostra."
                )

    n_inicial = int(len(dados))
    n_minimo = int(n_minimo)
    if n_minimo < 1:
        raise ValueError("O N mínimo deve ser pelo menos 1.")
    if n_minimo > n_inicial:
        raise ValueError(
            f"O N mínimo ({n_minimo}) não pode ser maior que a base atual ({n_inicial})."
        )

    limite_exclusoes = n_inicial - n_minimo
    ativos = pd.Series(True, index=dados.index)
    _, _, diag_inicial = _estado_margens(dados, ativos, alvos_norm)
    obj_atual = _objetivo_diagnostico(diag_inicial, tolerancia_pp)
    maior_antes = float(diag_inicial["diferenca_abs_pp"].max()) if not diag_inicial.empty else np.nan

    indices_excluir: list[Any] = []
    motivo = "A amostra bruta já está dentro da tolerância informada."

    while len(indices_excluir) < limite_exclusoes and (
        forcar_atingir_n_minimo
        or obj_atual[0] > 0
        or not parar_ao_atingir_tolerancia_bruta
    ):
        contagens_atual, denominadores_atual, diag_atual = _estado_margens(dados, ativos, alvos_norm)
        perfis = _perfis_candidatos(
            dados,
            ativos,
            variaveis,
            diag_atual,
            contagens=contagens_atual,
            alvos=alvos_norm,
            base_minima_categoria=base_minima_categoria,
            codigos_protegidos=codigos_protegidos,
            permitir_sem_excesso=forcar_atingir_n_minimo,
            limite=max_perfis_avaliados,
        )
        if not perfis:
            if forcar_atingir_n_minimo:
                motivo = (
                    "Não há mais casos elegíveis para exclusão sem violar as proteções configuradas."
                )
            else:
                motivo = "Não há mais perfis elegíveis em categorias com excesso que possam ser priorizados."
            break

        melhor_perfil = None
        melhor_obj = None if forcar_atingir_n_minimo else obj_atual
        for perfil in perfis:
            diag_sim = _diagnostico_apos_remover_perfil_contagens(
                contagens_atual, denominadores_atual, perfil, variaveis, alvos_norm
            )
            if diag_sim.empty:
                continue
            obj_sim = _objetivo_diagnostico(diag_sim, tolerancia_pp)
            if melhor_obj is None or obj_sim < melhor_obj:
                melhor_obj = obj_sim
                melhor_perfil = perfil

        if melhor_perfil is None:
            motivo = (
                "Nenhuma exclusão individual adicional respeitou as proteções e melhorou o conjunto de margens."
            )
            break

        idx = _escolher_indice_do_perfil(
            dados, ativos, variaveis, melhor_perfil, id_var
        )
        if idx is None:
            motivo = "Não foi possível localizar um caso correspondente ao melhor perfil sugerido."
            break

        ativos.loc[idx] = False
        indices_excluir.append(idx)
        obj_atual = melhor_obj
        motivo = "Busca concluída."

    contagens_finais, _, diag_final = _estado_margens(dados, ativos, alvos_norm)
    maior_depois = float(diag_final["diferenca_abs_pp"].max()) if not diag_final.empty else np.nan
    atingiu_tol = bool(np.isfinite(maior_depois) and maior_depois <= float(tolerancia_pp))
    n_final = n_inicial - len(indices_excluir)
    atingiu_n_alvo = n_final <= n_minimo
    faltam_exclusoes = max(0, n_final - n_minimo)
    atingiu_limite = len(indices_excluir) >= limite_exclusoes and not atingiu_tol

    if forcar_atingir_n_minimo:
        if atingiu_n_alvo and atingiu_tol:
            motivo = "O N-alvo foi atingido e as margens brutas ficaram dentro da tolerância."
        elif atingiu_n_alvo:
            motivo = (
                "O N-alvo foi atingido, mas ainda restam margens brutas fora da tolerância. "
                "Não foram sugeridas exclusões abaixo do N informado."
            )
        elif faltam_exclusoes > 0:
            motivo = (
                f"Não foi possível atingir o N-alvo sem violar as proteções. "
                f"Ainda seriam necessárias {faltam_exclusoes} exclusão(ões)."
            )
    elif atingiu_tol:
        motivo = "As margens brutas ficaram dentro da tolerância antes de atingir o N mínimo."
    elif atingiu_limite:
        motivo = (
            "O N mínimo permitido foi atingido e ainda restam margens fora da tolerância. "
            "O assistente não pode sugerir exclusões adicionais."
        )

    labels = labels or {}
    antes_lookup = {
        (row["variavel"], row["codigo"]): row for _, row in diag_inicial.iterrows()
    }
    linhas = []
    sobras = []
    for _, row in diag_final.iterrows():
        chave = (row["variavel"], row["codigo"])
        antes = antes_lookup[chave]
        categoria_label = labels.get(row["variavel"], {}).get(row["codigo"], str(row["codigo"]))
        excesso_n = 0.0
        # N aproximado acima do Universo na base final, usando o denominador válido da variável.
        denom_final = int(dados.loc[ativos, row["variavel"]].notna().sum())
        alvo_n = float(row["percentual_universo"]) / 100.0 * denom_final
        excesso_n = max(0.0, float(row["n"]) - alvo_n)
        item = {
            "variavel": row["variavel"],
            "codigo": row["codigo"],
            "categoria": categoria_label,
            "n_antes": int(antes["n"]),
            "pct_antes": float(antes["percentual_amostra"]),
            "pct_universo": float(row["percentual_universo"]),
            "diferenca_antes_pp": float(antes["diferenca_pp"]),
            "n_depois": int(row["n"]),
            "pct_depois": float(row["percentual_amostra"]),
            "diferenca_depois_pp": float(row["diferenca_pp"]),
            "excesso_n_aprox": float(excesso_n),
            "status": "OK" if float(row["diferenca_abs_pp"]) <= float(tolerancia_pp) else "REVISAR",
        }
        linhas.append(item)
        if float(row["diferenca_pp"]) > float(tolerancia_pp):
            sobras.append(item.copy())

    resumo = pd.DataFrame(linhas)
    sobras_df = pd.DataFrame(sobras)
    if not sobras_df.empty:
        sobras_df = sobras_df.sort_values(
            ["diferenca_depois_pp", "excesso_n_aprox"], ascending=False
        ).reset_index(drop=True)

    restricoes = []
    labels = labels or {}
    for variavel, mapa in alvos_norm.items():
        for categoria in mapa:
            n_cat = int(contagens_finais.get(variavel, {}).get(categoria, 0))
            motivos = []
            if categoria in codigos_protegidos.get(variavel, set()):
                motivos.append("Proteção manual: nunca excluir")
            if base_minima_categoria > 0 and n_cat <= base_minima_categoria:
                motivos.append(f"Base mínima da categoria atingida ({base_minima_categoria})")
            if motivos:
                restricoes.append(
                    {
                        "variavel": variavel,
                        "codigo": categoria,
                        "categoria": labels.get(variavel, {}).get(categoria, str(categoria)),
                        "n_final": n_cat,
                        "base_minima": base_minima_categoria,
                        "motivo": "; ".join(motivos),
                    }
                )
    restricoes_df = pd.DataFrame(restricoes)

    ids_excluir = dados.loc[indices_excluir, id_var].tolist()
    return ResultadoSugestaoExclusoes(
        ids_excluir=ids_excluir,
        indices_excluir=indices_excluir,
        resumo_margens=resumo,
        sobras_remanescentes=sobras_df,
        n_inicial=n_inicial,
        n_final=n_final,
        n_minimo=n_minimo,
        limite_exclusoes=limite_exclusoes,
        atingiu_tolerancia_bruta=atingiu_tol,
        atingiu_limite_exclusoes=atingiu_limite,
        maior_diferenca_bruta_antes_pp=maior_antes,
        maior_diferenca_bruta_depois_pp=maior_depois,
        motivo_parada=motivo,
        base_minima_categoria=base_minima_categoria,
        atingiu_n_alvo=atingiu_n_alvo,
        faltam_exclusoes_para_alvo=faltam_exclusoes,
        restricoes_ativas=restricoes_df,
    )


def gerar_syntax_exclusao_ids(
    id_var: str,
    ids_excluir: list[Any],
    tratar_como_texto: bool | None = None,
    modo: str = "marcar",
    nome_flag: str = "excluir_pond",
) -> str:
    """Gera syntax SPSS para marcar ou excluir os IDs sugeridos."""
    if not ids_excluir:
        return "* Nenhuma exclusao sugerida pelo assistente de ponderacao. *."

    if tratar_como_texto is None:
        ids_sao_texto = any(isinstance(i, (str, np.str_)) for i in ids_excluir)
    else:
        ids_sao_texto = bool(tratar_como_texto)

    def fmt(valor: Any) -> str:
        if ids_sao_texto:
            texto = str(valor).replace('"', '""')
            return f'"{texto}"'
        try:
            numero = float(valor)
            if numero.is_integer():
                return str(int(numero))
        except (TypeError, ValueError):
            pass
        return str(valor)

    valores = [fmt(v) for v in ids_excluir]
    por_linha = 10
    grupos = [valores[i : i + por_linha] for i in range(0, len(valores), por_linha)]
    corpo = ",\n    ".join(", ".join(g) for g in grupos)

    cabecalho = [
        "* Sugestao de exclusoes para teste de ponderacao - gerado pelo Analitico. *.",
        f"* Total de IDs sugeridos: {len(ids_excluir)}. *.",
        "* Revise os IDs e salve uma copia do banco antes de aplicar exclusoes. *.",
        "",
    ]

    if modo == "excluir":
        return "\n".join(
            cabecalho
            + [
                "FILTER OFF.",
                "USE ALL.",
                f"SELECT IF (NOT ANY({id_var},",
                f"    {corpo})).",
                "EXECUTE.",
            ]
        )

    return "\n".join(
        cabecalho
        + [
            f"COMPUTE {nome_flag}=ANY({id_var},",
            f"    {corpo}).",
            f"VARIABLE LABELS {nome_flag} 'Sugestao de exclusao para melhorar ponderacao'.",
            f"VALUE LABELS {nome_flag} 0 'Manter' 1 'Excluir sugerido'.",
            f"FREQUENCIES VARIABLES={nome_flag}.",
            "EXECUTE.",
            "",
            "* Depois de conferir a marcacao acima, use as linhas abaixo para excluir. *.",
            f"* SELECT IF ({nome_flag}=0).",
            "* EXECUTE.",
        ]
    )
