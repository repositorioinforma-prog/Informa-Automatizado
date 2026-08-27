"""
Motor de cálculo do Gerador de Amostra.

Recebe os parâmetros da pesquisa e devolve a distribuição amostral pronta,
com número de entrevistas por unidade territorial + banda mín/máx + cotas
demográficas (sexo/idade/renda) dentro de cada unidade.

Lógica pura, sem dependência de Streamlit — testável isoladamente com
scripts Python diretos, seguindo a convenção de arquitetura do projeto
(core/ nunca importa streamlit; quem chama a UI é analises/gerador_amostra.py).

Fonte dos dados: bases mestre pré-processadas em dados/gerador_amostra/
(master_municipios.csv e master_distritos.csv), consolidadas a partir de
DTB 2025, Censo 2022 (sexo/idade) e Censo 2010 (renda) via
scripts/amostrador_build_master.py.

Uso como biblioteca:
    from core.amostra_math import calcular_amostra
    resultado = calcular_amostra(
        uf="SP",
        amostra_total=1000,
        nivel_quebra="regiao_intermediaria",
        base_populacional="16_mais",
    )
"""
from pathlib import Path
import pandas as pd

RAIZ_PROJETO = Path(__file__).parent.parent
BASE = RAIZ_PROJETO / "dados" / "gerador_amostra"


# ---------------------------------------------------------------------------
# Colunas de população por faixa (usadas em várias funções)
# ---------------------------------------------------------------------------
COLS_SEXO = ["masc_total", "fem_total"]
FAIXAS_ETARIAS = ["16_19", "20_29", "30_39", "40_49", "50mais"]
COLS_IDADE_MASC = [f"masc_{f}" for f in FAIXAS_ETARIAS]
COLS_IDADE_FEM = [f"fem_{f}" for f in FAIXAS_ETARIAS]
COLS_IDADE_TODAS = COLS_IDADE_MASC + COLS_IDADE_FEM
COLS_RENDA = ["renda_ate_2sm", "renda_2_a_5sm", "renda_5_a_10sm", "renda_mais_10sm"]
COLS_AUX_AGREGACAO = [
    "_sexo_masc_peso",
    "_sexo_fem_peso",
    "_renda_unidades_total",
    "_renda_unidades_disponiveis",
]

COLS_COTAS_RESULTADO = (
    ["cota_masc", "cota_fem"]
    + [f"cota_{f}" for f in FAIXAS_ETARIAS]
    + [f"cota_{c}" for c in COLS_RENDA]
    + ["renda_disponivel", "renda_cobertura_percentual"]
)


def remover_cotas_calculadas(df):
    """Remove cotas já calculadas antes de refazê-las para outro N."""
    return df.drop(columns=COLS_COTAS_RESULTADO, errors="ignore")


def adicionar_memoria_calculo(df, amostra_total):
    """Acrescenta colunas de auditoria da distribuição amostral.

    ``percentual`` continua representando a participação populacional real do
    recorte. ``percentual_amostra`` representa a participação efetivamente
    alocada. Em modo manual, a diferença entre ambas fica explícita.
    """
    resultado = df.copy()
    total = int(amostra_total)
    if total <= 0:
        raise ValueError("amostra_total deve ser maior que zero.")

    populacao = pd.to_numeric(resultado.get("populacao_base"), errors="coerce").fillna(0)
    soma_populacao = float(populacao.sum())
    if "amostra_ideal" not in resultado.columns:
        resultado["amostra_ideal"] = (
            populacao / soma_populacao * total if soma_populacao > 0 else 0.0
        )
    if "amostra_proporcional" not in resultado.columns:
        resultado["amostra_proporcional"] = pd.to_numeric(
            resultado["amostra"], errors="coerce"
        ).fillna(0).astype(int)

    amostra = pd.to_numeric(resultado["amostra"], errors="coerce").fillna(0)
    resultado["percentual_amostra"] = amostra / total * 100
    resultado["diferenca_vs_proporcional"] = (
        amostra - pd.to_numeric(resultado["amostra_proporcional"], errors="coerce").fillna(0)
    ).astype(int)
    if "metodo_distribuicao" not in resultado.columns:
        resultado["metodo_distribuicao"] = "proporcional"
    return resultado


# ---------------------------------------------------------------------------
# Carregamento das bases mestre (cached em módulo)
# ---------------------------------------------------------------------------
_cache = {}


def _carregar_municipios():
    if "mun" not in _cache:
        df = pd.read_csv(BASE / "master_municipios.csv", dtype={
            "cod_municipio": str,
            "cod_regiao_intermediaria": str,
            "cod_regiao_imediata": str,
        })
        # Força colunas de renda a numérico (algumas vêm como "-" no censo)
        for c in COLS_RENDA:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        _cache["mun"] = df
    return _cache["mun"]


def _carregar_distritos():
    if "dist" not in _cache:
        df = pd.read_csv(BASE / "master_distritos.csv", dtype={
            "cod_distrito": str,
            "cod_municipio": str,
            "cod_regiao_intermediaria": str,
            "cod_regiao_imediata": str,
        })
        for c in COLS_RENDA:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        _cache["dist"] = df
    return _cache["dist"]


# ---------------------------------------------------------------------------
# Método do Maior Resto (Largest Remainder / Hamilton)
# ---------------------------------------------------------------------------
def maior_resto(pesos, total):
    """
    Distribui um total inteiro entre N unidades proporcionalmente aos pesos,
    garantindo que a soma dos resultados == total (sem perder por arredondamento).

    Args:
        pesos: lista de números (populações, por exemplo)
        total: inteiro a ser distribuído

    Returns:
        lista de inteiros, mesma ordem dos pesos, soma == total
    """
    soma_pesos = sum(pesos)
    if soma_pesos == 0:
        return [0] * len(pesos)

    # Valor "ideal" com decimal
    ideais = [p * total / soma_pesos for p in pesos]
    # Parte inteira
    inteiros = [int(x) for x in ideais]
    # Restos (parte decimal)
    restos = [(ideais[i] - inteiros[i], i) for i in range(len(pesos))]
    # Faltam N unidades pra chegar no total
    faltam = total - sum(inteiros)
    # Distribui as N unidades restantes para as unidades com maior resto
    restos.sort(reverse=True)
    for k in range(faltam):
        _, idx = restos[k]
        inteiros[idx] += 1
    return inteiros


def preparar_colunas_auxiliares(df, base_populacional):
    """Prepara pesos de sexo e indicadores de cobertura de renda.

    Bases novas podem fornecer sexo especificamente para 16+. As bases legadas
    não possuem esse recorte e continuam usando sexo total, preservando o
    comportamento anterior como fallback.
    """
    df = df.copy()
    if (
        base_populacional == "16_mais"
        and {"masc_16mais", "fem_16mais"}.issubset(df.columns)
    ):
        df["_sexo_masc_peso"] = pd.to_numeric(df["masc_16mais"], errors="coerce").fillna(0)
        df["_sexo_fem_peso"] = pd.to_numeric(df["fem_16mais"], errors="coerce").fillna(0)
    else:
        df["_sexo_masc_peso"] = pd.to_numeric(df["masc_total"], errors="coerce").fillna(0)
        df["_sexo_fem_peso"] = pd.to_numeric(df["fem_total"], errors="coerce").fillna(0)

    for coluna in COLS_RENDA:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    if "renda_disponivel" in df.columns:
        disponibilidade = pd.to_numeric(df["renda_disponivel"], errors="coerce").fillna(0) > 0
    else:
        disponibilidade = df[COLS_RENDA].fillna(0).sum(axis=1) > 0
    df["_renda_unidades_total"] = 1
    df["_renda_unidades_disponiveis"] = disponibilidade.astype(int)
    return df


# ---------------------------------------------------------------------------
# Cotas demográficas (reutilizadas pelo motor clássico e pelo motor em camadas)
# ---------------------------------------------------------------------------
def calcular_cotas_demograficas_da_unidade(row):
    """Calcula sexo, idade e renda para a amostra já atribuída a uma unidade."""
    n = int(row["amostra"])
    nomes_cotas = (
        ["cota_masc", "cota_fem"]
        + [f"cota_{f}" for f in FAIXAS_ETARIAS]
        + [f"cota_{c}" for c in COLS_RENDA]
    )
    total_unidades_renda = int(row.get("_renda_unidades_total", 1) or 0)
    unidades_renda_disponiveis = int(row.get("_renda_unidades_disponiveis", 1) or 0)
    renda_completa = (
        total_unidades_renda > 0
        and unidades_renda_disponiveis == total_unidades_renda
    )
    cobertura_renda = (
        unidades_renda_disponiveis / total_unidades_renda * 100
        if total_unidades_renda > 0
        else 0.0
    )

    if n == 0:
        saida_zero = {c: 0 for c in nomes_cotas}
        if not renda_completa:
            for c in COLS_RENDA:
                saida_zero[f"cota_{c}"] = pd.NA
        saida_zero["renda_disponivel"] = int(renda_completa)
        saida_zero["renda_cobertura_percentual"] = cobertura_renda
        return pd.Series(saida_zero)

    cotas_sexo = maior_resto(
        [row.get("_sexo_masc_peso", row["masc_total"]), row.get("_sexo_fem_peso", row["fem_total"])],
        n,
    )
    pop_idade = [row[f"masc_{f}"] + row[f"fem_{f}"] for f in FAIXAS_ETARIAS]
    cotas_idade = maior_resto(pop_idade, n)
    pop_renda = [row[c] for c in COLS_RENDA]
    soma_renda = sum(float(valor or 0) for valor in pop_renda if not pd.isna(valor))
    cotas_renda = maior_resto(pop_renda, n) if renda_completa and soma_renda > 0 else None

    out = {"cota_masc": cotas_sexo[0], "cota_fem": cotas_sexo[1]}
    for i, f in enumerate(FAIXAS_ETARIAS):
        out[f"cota_{f}"] = cotas_idade[i]
    for i, c in enumerate(COLS_RENDA):
        out[f"cota_{c}"] = cotas_renda[i] if cotas_renda is not None else pd.NA
    out["renda_disponivel"] = int(cotas_renda is not None)
    out["renda_cobertura_percentual"] = cobertura_renda
    return pd.Series(out)


def adicionar_cotas_demograficas(df):
    """Adiciona ao DataFrame as cotas demográficas por unidade."""
    cotas = df.apply(calcular_cotas_demograficas_da_unidade, axis=1)
    return pd.concat([df.reset_index(drop=True), cotas.reset_index(drop=True)], axis=1)


# ---------------------------------------------------------------------------
# Função principal: distribuição amostral por unidade territorial
# ---------------------------------------------------------------------------
def calcular_amostra(
    uf=None,
    codigos_municipios=None,
    codigos_regioes_intermediarias=None,
    codigos_regioes_imediatas=None,
    amostra_total=1000,
    nivel_quebra="regiao_intermediaria",
    base_populacional="16_mais",
    incluir_distritos=False,
    faixa_flex=0.20,
    df_base=None,
    codigos_distritos=None,
):
    """
    Calcula a distribuição amostral proporcional à população.

    Args:
        uf: sigla do estado (ex: "SP"). Se None, usa Brasil inteiro.
        codigos_municipios: lista de códigos IBGE de municípios (filtro extra opcional).
            Ex: ["3550308", "3509502"] para recortar Capital de SP + outra cidade.
        codigos_distritos: lista de códigos IBGE de distritos (filtro extra opcional).
            Requer uma base distrital, mesmo quando a saída final é agregada em município
            ou região.
        codigos_regioes_intermediarias: lista de códigos (filtro extra opcional).
        codigos_regioes_imediatas: lista de códigos (filtro extra opcional).
        amostra_total: total de entrevistas a distribuir (int).
        nivel_quebra: nível territorial onde a distribuição é calculada.
            Valores: "regiao_intermediaria", "regiao_imediata", "municipio", "distrito"
        base_populacional: qual população usar como peso.
            Valores: "total" (todas as idades) ou "16_mais" (soma das faixas da pesquisa)
        incluir_distritos: se True, usa master_distritos; senão, master_municipios.
        faixa_flex: percentual da banda mín/máx (0.20 = ±20%).
        df_base: base já padronizada, opcional. Quando ausente, usa os CSVs
            legados. Permite que o Streamlit consulte o SQLite sem duplicar o
            motor matemático.

    Returns:
        pandas.DataFrame com colunas:
            - hierarquia territorial (varia com nivel_quebra)
            - populacao_base, percentual, amostra, minimo, maximo
            - cotas demográficas dentro de cada unidade
    """
    # 1. Escolhe a base (município ou distrito)
    if df_base is not None:
        df = df_base.copy()
    else:
        df = _carregar_distritos().copy() if incluir_distritos else _carregar_municipios().copy()

    # 2. Aplica filtros de recorte territorial
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
                "Ative o recorte por distritos ou use nível de quebra Distrito."
            )
        df = df[df["cod_distrito"].isin([str(c) for c in codigos_distritos])]

    if df.empty:
        raise ValueError("Nenhuma unidade encontrada com esses filtros.")

    # 3. Define coluna de "peso" populacional
    if base_populacional == "total":
        df["_peso"] = df["populacao_total"]
    elif base_populacional == "16_mais":
        if "populacao_16_mais_referencia" in df.columns:
            df["_peso"] = pd.to_numeric(
                df["populacao_16_mais_referencia"], errors="coerce"
            ).fillna(0)
        else:
            df["_peso"] = df[COLS_IDADE_TODAS].sum(axis=1)
    else:
        raise ValueError(f"base_populacional inválido: {base_populacional}")

    df = preparar_colunas_auxiliares(df, base_populacional)

    # 4. Define agrupamento pelo nível de quebra
    mapa_quebra = {
        "regiao_intermediaria": ["cod_regiao_intermediaria", "regiao_intermediaria"],
        "regiao_imediata": [
            "cod_regiao_intermediaria", "regiao_intermediaria",
            "cod_regiao_imediata", "regiao_imediata",
        ],
        "municipio": [
            "cod_regiao_intermediaria", "regiao_intermediaria",
            "cod_regiao_imediata", "regiao_imediata",
            "cod_municipio", "municipio",
        ],
        "distrito": [
            "cod_regiao_intermediaria", "regiao_intermediaria",
            "cod_regiao_imediata", "regiao_imediata",
            "cod_municipio", "municipio",
            "cod_distrito", "distrito",
        ],
    }
    if nivel_quebra not in mapa_quebra:
        raise ValueError(f"nivel_quebra inválido: {nivel_quebra}")
    if nivel_quebra == "distrito" and not incluir_distritos:
        raise ValueError("Para quebra por distrito, use incluir_distritos=True")

    chaves_grupo = mapa_quebra[nivel_quebra]

    # 5. Agrega população e cotas demográficas por unidade
    colunas_soma = (
        ["_peso", "populacao_total"]
        + COLS_SEXO
        + COLS_IDADE_TODAS
        + COLS_RENDA
        + COLS_AUX_AGREGACAO
    )
    agregado = df.groupby(chaves_grupo, as_index=False)[colunas_soma].sum(min_count=1)

    # 6. Distribui a amostra total proporcionalmente ao peso (com Maior Resto)
    pesos = agregado["_peso"].tolist()
    soma_pesos = sum(pesos)
    agregado["amostra_ideal"] = agregado["_peso"] / soma_pesos * int(amostra_total)
    amostras = maior_resto(pesos, amostra_total)
    agregado["amostra"] = amostras
    agregado["amostra_proporcional"] = amostras

    # 7. Calcula percentual e banda mín/máx por unidade
    agregado["percentual"] = agregado["_peso"] / soma_pesos * 100
    agregado["minimo"] = agregado["amostra"].apply(lambda n: max(0, round(n * (1 - faixa_flex))))
    agregado["maximo"] = agregado["amostra"].apply(lambda n: round(n * (1 + faixa_flex)))

    # Mantém explicitamente a população usada como peso. Isso evita que a UI
    # chame populacao_total de "população base" quando a pesquisa usa 16+.
    agregado["populacao_base"] = agregado["_peso"]

    # 8. Calcula cotas demográficas DENTRO de cada unidade (com Maior Resto)
    resultado = adicionar_cotas_demograficas(agregado)
    resultado = adicionar_memoria_calculo(resultado, int(amostra_total))

    # 9. Ordena alfabeticamente pela hierarquia e limpa colunas auxiliares
    colunas_ordem = [c for c in [
        "regiao_intermediaria", "regiao_imediata", "municipio", "distrito"
    ] if c in resultado.columns]
    resultado = resultado.sort_values(colunas_ordem).reset_index(drop=True)
    resultado = resultado.drop(columns=["_peso", *COLS_AUX_AGREGACAO], errors="ignore")

    return resultado


# ---------------------------------------------------------------------------
# Teste rápido: reproduz o exemplo SP-1000
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("Teste 1: SP, 1000 entrevistas, quebra por Região Intermediária, base 16+\n")
    r = calcular_amostra(
        uf="SP",
        amostra_total=1000,
        nivel_quebra="regiao_intermediaria",
        base_populacional="16_mais",
    )
    print(r[[
        "regiao_intermediaria", "populacao_total",
        "percentual", "amostra", "minimo", "maximo",
        "cota_masc", "cota_fem",
    ]].to_string(index=False))
    print(f"\nTotal amostra: {r['amostra'].sum()} (esperado: 1000)")
    print(f"Total população: {r['populacao_total'].sum():,}")
