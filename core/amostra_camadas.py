"""
Motor de montagem de amostra por camadas territoriais.

A proposta deste módulo é separar o desenho hierárquico da amostra da tela.
Cada camada recebe a quantidade fechada pelo nível pai e redistribui esse
valor entre seus filhos. Assim, os arredondamentos são resolvidos localmente
e cada nível da árvore fecha exatamente com o nível anterior.

Versão 1:
    - camadas territoriais IBGE;
    - distribuição proporcional à população;
    - fechamento recursivo pai -> filhos;
    - resumos por camada e validações estruturais.

As próximas estratégias (mínimo + proporcional, fixo, manual, hard/soft e
travamentos) podem ser adicionadas sem alterar o contrato básico do motor.
"""
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd

from core.amostra_math import (
    COLS_AUX_AGREGACAO,
    COLS_IDADE_TODAS,
    COLS_RENDA,
    COLS_SEXO,
    _carregar_distritos,
    _carregar_municipios,
    adicionar_cotas_demograficas,
    adicionar_memoria_calculo,
    maior_resto,
    preparar_colunas_auxiliares,
)
from core.amostra_alocacao import aplicar_alocacao_manual


ORDEM_NIVEIS = [
    "uf",
    "regiao_intermediaria",
    "regiao_imediata",
    "municipio",
    "distrito",
]

ROTULOS_NIVEIS = {
    "uf": "Estado (UF)",
    "regiao_intermediaria": "Região Intermediária",
    "regiao_imediata": "Região Imediata",
    "municipio": "Município",
    "distrito": "Distrito",
}

# Chaves cumulativas. Mesmo quando um nível intermediário não é controlado,
# mantemos a hierarquia completa para evitar colisão de nomes/códigos e para
# facilitar a leitura dos resumos.
CHAVES_POR_NIVEL = {
    "uf": ["uf", "nome_uf"],
    "regiao_intermediaria": [
        "uf", "nome_uf",
        "cod_regiao_intermediaria", "regiao_intermediaria",
    ],
    "regiao_imediata": [
        "uf", "nome_uf",
        "cod_regiao_intermediaria", "regiao_intermediaria",
        "cod_regiao_imediata", "regiao_imediata",
    ],
    "municipio": [
        "uf", "nome_uf",
        "cod_regiao_intermediaria", "regiao_intermediaria",
        "cod_regiao_imediata", "regiao_imediata",
        "cod_municipio", "municipio",
    ],
    "distrito": [
        "uf", "nome_uf",
        "cod_regiao_intermediaria", "regiao_intermediaria",
        "cod_regiao_imediata", "regiao_imediata",
        "cod_municipio", "municipio",
        "cod_distrito", "distrito",
    ],
}

COLUNAS_SOMA = [
    "_peso",
    "populacao_total",
    *COLS_SEXO,
    *COLS_IDADE_TODAS,
    *COLS_RENDA,
    *COLS_AUX_AGREGACAO,
]


@dataclass
class ResultadoAmostraCamadas:
    """Resultado completo do motor hierárquico."""

    resultado: pd.DataFrame
    resumos: Dict[str, pd.DataFrame]
    validacoes: List[dict]
    camadas: List[str]


# ---------------------------------------------------------------------------
# Validação e preparação
# ---------------------------------------------------------------------------
def normalizar_camadas(camadas):
    """Valida e ordena uma sequência de níveis territoriais.

    A ordem informada precisa respeitar a hierarquia IBGE. É permitido pular
    níveis (ex.: Região Intermediária -> Município), mas não voltar na árvore.
    """
    if not camadas:
        raise ValueError("Selecione pelo menos uma camada territorial.")

    camadas = list(camadas)
    desconhecidas = [c for c in camadas if c not in ORDEM_NIVEIS]
    if desconhecidas:
        raise ValueError(f"Camada(s) inválida(s): {', '.join(desconhecidas)}")

    if len(camadas) != len(set(camadas)):
        raise ValueError("Uma camada territorial não pode aparecer mais de uma vez.")

    indices = [ORDEM_NIVEIS.index(c) for c in camadas]
    if indices != sorted(indices):
        raise ValueError(
            "As camadas devem seguir a hierarquia territorial: "
            "UF -> Região Intermediária -> Região Imediata -> Município -> Distrito."
        )
    return camadas


def _preparar_base(
    df_base,
    base_populacional="16_mais",
    uf=None,
    codigos_municipios=None,
    codigos_distritos=None,
    codigos_regioes_intermediarias=None,
    codigos_regioes_imediatas=None,
):
    df = df_base.copy()

    if uf:
        df = df[df["uf"] == uf]
    if codigos_regioes_intermediarias:
        df = df[df["cod_regiao_intermediaria"].isin(codigos_regioes_intermediarias)]
    if codigos_regioes_imediatas:
        df = df[df["cod_regiao_imediata"].isin(codigos_regioes_imediatas)]
    if codigos_municipios:
        df = df[df["cod_municipio"].isin([str(c) for c in codigos_municipios])]
    if codigos_distritos:
        if "cod_distrito" not in df.columns:
            raise ValueError(
                "A seleção de distritos exige uma base distrital. "
                "Ative o recorte por distritos ou inclua a camada Distrito."
            )
        df = df[df["cod_distrito"].isin([str(c) for c in codigos_distritos])]

    if df.empty:
        raise ValueError("Nenhuma unidade encontrada com esses filtros.")

    if base_populacional == "total":
        df["_peso"] = pd.to_numeric(df["populacao_total"], errors="coerce").fillna(0)
    elif base_populacional == "16_mais":
        if "populacao_16_mais_referencia" in df.columns:
            df["_peso"] = pd.to_numeric(
                df["populacao_16_mais_referencia"], errors="coerce"
            ).fillna(0)
        else:
            df["_peso"] = df[COLS_IDADE_TODAS].apply(
                pd.to_numeric, errors="coerce"
            ).fillna(0).sum(axis=1)
    else:
        raise ValueError(f"base_populacional inválido: {base_populacional}")

    if float(df["_peso"].sum()) <= 0:
        raise ValueError("A população-base do recorte selecionado é zero.")

    return preparar_colunas_auxiliares(df, base_populacional)


def _agregar_nivel(df, nivel):
    chaves = CHAVES_POR_NIVEL[nivel]
    faltantes = [c for c in chaves + COLUNAS_SOMA if c not in df.columns]
    if faltantes:
        raise ValueError(
            "A base não possui todas as colunas necessárias para a camada "
            f"{nivel}: {', '.join(faltantes)}"
        )
    return df.groupby(chaves, as_index=False, dropna=False)[COLUNAS_SOMA].sum(min_count=1)


def _filtrar_pelo_pai(df, parent_row, chaves_pai):
    mascara = pd.Series(True, index=df.index)
    for chave in chaves_pai:
        valor = parent_row[chave]
        if pd.isna(valor):
            mascara &= df[chave].isna()
        else:
            mascara &= df[chave] == valor
    return df[mascara]


def _adicionar_percentuais(df, peso_total, chaves_pai=None):
    df = df.copy()
    df["populacao_base"] = df["_peso"]
    df["percentual"] = df["_peso"] / peso_total * 100

    if not chaves_pai:
        df["percentual_no_pai"] = df["percentual"]
    else:
        soma_pai = df.groupby(chaves_pai, dropna=False)["_peso"].transform("sum")
        df["percentual_no_pai"] = df["_peso"] / soma_pai * 100
    return df


# ---------------------------------------------------------------------------
# Distribuição pura: aceita DataFrame injetado e é fácil de testar
# ---------------------------------------------------------------------------
def distribuir_amostra_em_camadas(
    df_base,
    amostra_total,
    camadas,
    base_populacional="16_mais",
    faixa_flex=0.20,
    uf=None,
    codigos_municipios=None,
    codigos_regioes_intermediarias=None,
    codigos_regioes_imediatas=None,
    codigos_distritos=None,
):
    """Distribui a amostra recursivamente entre as camadas selecionadas."""
    camadas = normalizar_camadas(camadas)
    amostra_total = int(amostra_total)
    if amostra_total <= 0:
        raise ValueError("amostra_total deve ser maior que zero.")
    if faixa_flex < 0:
        raise ValueError("faixa_flex não pode ser negativa.")

    df = _preparar_base(
        df_base=df_base,
        base_populacional=base_populacional,
        uf=uf,
        codigos_municipios=codigos_municipios,
        codigos_distritos=codigos_distritos,
        codigos_regioes_intermediarias=codigos_regioes_intermediarias,
        codigos_regioes_imediatas=codigos_regioes_imediatas,
    )
    peso_total = float(df["_peso"].sum())

    resumos = {}
    resumo_anterior = None
    nivel_anterior = None

    for nivel in camadas:
        if resumo_anterior is None:
            resumo = _agregar_nivel(df, nivel)
            pesos = resumo["_peso"].tolist()
            if sum(pesos) <= 0:
                raise ValueError(f"A camada {ROTULOS_NIVEIS[nivel]} não possui peso populacional.")
            resumo["amostra_ideal"] = [
                peso * amostra_total / sum(pesos) for peso in pesos
            ]
            resumo["amostra"] = maior_resto(pesos, amostra_total)
            resumo["amostra_proporcional"] = resumo["amostra"]
            resumo = _adicionar_percentuais(resumo, peso_total)
            resumo = adicionar_memoria_calculo(resumo, amostra_total)
        else:
            chaves_pai = CHAVES_POR_NIVEL[nivel_anterior]
            blocos = []

            for _, pai in resumo_anterior.iterrows():
                base_pai = _filtrar_pelo_pai(df, pai, chaves_pai)
                filhos = _agregar_nivel(base_pai, nivel)
                pesos = filhos["_peso"].tolist()
                n_pai = int(pai["amostra"])

                if sum(pesos) <= 0 and n_pai > 0:
                    nome = " / ".join(str(pai[c]) for c in chaves_pai if not str(c).startswith("cod_"))
                    raise ValueError(
                        f"Não é possível distribuir {n_pai} entrevistas abaixo de {nome}: "
                        "o peso populacional dos filhos é zero."
                    )

                if n_pai > 0 and sum(pesos) > 0:
                    filhos["amostra_ideal"] = [
                        peso * n_pai / sum(pesos) for peso in pesos
                    ]
                    filhos["amostra"] = maior_resto(pesos, n_pai)
                else:
                    filhos["amostra_ideal"] = 0.0
                    filhos["amostra"] = 0
                filhos["amostra_proporcional"] = filhos["amostra"]
                filhos["amostra_pai"] = n_pai
                blocos.append(filhos)

            resumo = pd.concat(blocos, ignore_index=True) if blocos else pd.DataFrame()
            resumo = _adicionar_percentuais(resumo, peso_total, chaves_pai=chaves_pai)
            resumo = adicionar_memoria_calculo(resumo, amostra_total)

        chaves_ordem = CHAVES_POR_NIVEL[nivel]
        resumo = resumo.sort_values(chaves_ordem).reset_index(drop=True)
        resumos[nivel] = resumo.copy()
        resumo_anterior = resumo
        nivel_anterior = nivel

    resultado = resumo_anterior.copy()
    resultado["minimo"] = resultado["amostra"].apply(
        lambda n: max(0, round(int(n) * (1 - faixa_flex)))
    )
    resultado["maximo"] = resultado["amostra"].apply(
        lambda n: round(int(n) * (1 + faixa_flex))
    )
    resultado = adicionar_cotas_demograficas(resultado)
    resultado = adicionar_memoria_calculo(resultado, amostra_total)
    resultado = resultado.drop(columns=["_peso", *COLS_AUX_AGREGACAO], errors="ignore")

    # Também removemos _peso dos resumos entregues à interface. Mantemos
    # populacao_base, que é a informação legível equivalente.
    for nivel, resumo in list(resumos.items()):
        resumos[nivel] = resumo.drop(
            columns=["_peso", *COLS_AUX_AGREGACAO],
            errors="ignore",
        )

    validacoes = validar_resultado_camadas(
        resumos=resumos,
        camadas=camadas,
        amostra_total=amostra_total,
    )

    return ResultadoAmostraCamadas(
        resultado=resultado,
        resumos=resumos,
        validacoes=validacoes,
        camadas=camadas,
    )


# ---------------------------------------------------------------------------
# Wrapper usando as bases mestre do projeto
# ---------------------------------------------------------------------------
def calcular_amostra_camadas(
    camadas,
    amostra_total=1000,
    uf=None,
    codigos_municipios=None,
    codigos_regioes_intermediarias=None,
    codigos_regioes_imediatas=None,
    base_populacional="16_mais",
    faixa_flex=0.20,
    df_base=None,
    codigos_distritos=None,
):
    """Calcula uma amostra em camadas usando as bases mestre IBGE do projeto."""
    camadas = normalizar_camadas(camadas)
    if df_base is None:
        usar_distritos = "distrito" in camadas
        df_base = _carregar_distritos() if usar_distritos else _carregar_municipios()

    return distribuir_amostra_em_camadas(
        df_base=df_base,
        amostra_total=amostra_total,
        camadas=camadas,
        base_populacional=base_populacional,
        faixa_flex=faixa_flex,
        uf=uf,
        codigos_municipios=codigos_municipios,
        codigos_distritos=codigos_distritos,
        codigos_regioes_intermediarias=codigos_regioes_intermediarias,
        codigos_regioes_imediatas=codigos_regioes_imediatas,
    )


# ---------------------------------------------------------------------------
# Aplicação manual no último nível e reconstrução dos fechamentos
# ---------------------------------------------------------------------------
def aplicar_alocacao_manual_camadas(
    calculo,
    *,
    alocacoes_manuais,
    amostra_total,
    faixa_flex=0.20,
):
    """Aplica valores manuais nas unidades finais e agrega os níveis pais.

    A edição acontece sempre no último nível escolhido pelo usuário. Os
    controles superiores são reconstruídos pela soma dos filhos, mantendo o
    fechamento estrutural da árvore.
    """
    if not isinstance(calculo, ResultadoAmostraCamadas):
        raise ValueError("calculo precisa ser um ResultadoAmostraCamadas.")

    nivel_final = calculo.camadas[-1]
    resultado = aplicar_alocacao_manual(
        calculo.resultado,
        nivel=nivel_final,
        alocacoes_manuais=alocacoes_manuais,
        amostra_total=int(amostra_total),
        faixa_flex=faixa_flex,
    )

    resumos = {}
    for nivel in calculo.camadas:
        chaves = CHAVES_POR_NIVEL[nivel]
        original = calculo.resumos[nivel].copy()
        if "amostra_proporcional" not in original.columns:
            original["amostra_proporcional"] = original["amostra"]

        amostra_nova = (
            resultado.groupby(chaves, dropna=False, as_index=False)["amostra"]
            .sum()
            .rename(columns={"amostra": "_amostra_manual"})
        )
        original = original.drop(columns=["amostra"], errors="ignore").merge(
            amostra_nova,
            on=chaves,
            how="left",
            validate="one_to_one",
        )
        original["amostra"] = original.pop("_amostra_manual").fillna(0).astype(int)
        original["percentual_amostra"] = original["amostra"] / int(amostra_total) * 100
        original["diferenca_vs_proporcional"] = (
            original["amostra"]
            - pd.to_numeric(original["amostra_proporcional"], errors="coerce").fillna(0)
        ).astype(int)
        original["metodo_distribuicao"] = "manual"
        resumos[nivel] = original

    validacoes = validar_resultado_camadas(
        resumos=resumos,
        camadas=calculo.camadas,
        amostra_total=int(amostra_total),
    )
    return ResultadoAmostraCamadas(
        resultado=resultado,
        resumos=resumos,
        validacoes=validacoes,
        camadas=list(calculo.camadas),
    )


# ---------------------------------------------------------------------------
# Validações estruturais
# ---------------------------------------------------------------------------
def validar_resultado_camadas(resumos, camadas, amostra_total):
    validacoes = []

    for nivel in camadas:
        total_nivel = int(resumos[nivel]["amostra"].sum())
        validacoes.append({
            "ok": total_nivel == int(amostra_total),
            "tipo": "fechamento_nivel",
            "nivel": nivel,
            "mensagem": (
                f"{ROTULOS_NIVEIS[nivel]} fecha em {total_nivel} entrevistas "
                f"(esperado: {int(amostra_total)})."
            ),
        })

    for idx in range(1, len(camadas)):
        nivel_pai = camadas[idx - 1]
        nivel_filho = camadas[idx]
        pais = resumos[nivel_pai]
        filhos = resumos[nivel_filho]
        chaves_pai = CHAVES_POR_NIVEL[nivel_pai]

        soma_filhos = (
            filhos.groupby(chaves_pai, dropna=False, as_index=False)["amostra"]
            .sum()
            .rename(columns={"amostra": "amostra_filhos"})
        )
        comparacao = pais[chaves_pai + ["amostra"]].merge(
            soma_filhos,
            on=chaves_pai,
            how="left",
        )
        comparacao["amostra_filhos"] = comparacao["amostra_filhos"].fillna(0).astype(int)
        divergentes = comparacao[
            comparacao["amostra"].astype(int) != comparacao["amostra_filhos"]
        ]

        validacoes.append({
            "ok": divergentes.empty,
            "tipo": "fechamento_pai_filhos",
            "nivel": f"{nivel_pai}->{nivel_filho}",
            "mensagem": (
                f"Fechamento {ROTULOS_NIVEIS[nivel_pai]} -> "
                f"{ROTULOS_NIVEIS[nivel_filho]}: "
                f"{0 if divergentes.empty else len(divergentes)} divergência(s)."
            ),
        })

    return validacoes
