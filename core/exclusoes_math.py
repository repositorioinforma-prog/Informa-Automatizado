"""
Motor de cálculo do equilíbrio de cotas por exclusão.

Diferente de ponderação (rim weighting/raking), aqui a ideia é remover
entrevistas do banco até que a contagem de cada categoria de cada variável
de cota fique dentro de uma tolerância em torno da meta esperada — sem
nunca deixar nenhuma categoria abaixo de uma base mínima, e sem excluir
ninguém de categorias que já estão abaixo da própria meta (essas só podem
ser corrigidas usando as outras categorias, nunca removendo mais gente
delas).

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import pandas as pd


def calcular_exclusoes(dados, id_var, metas, tolerancia_n, base_minima=30):
    """
    Parâmetros
    ----------
    dados : DataFrame com uma linha por entrevistado.
    id_var : nome da coluna que identifica unicamente cada entrevistado.
    metas : dict {variavel: {categoria: meta_n (int)}}.
    tolerancia_n : dict {variavel: {categoria: tolerancia absoluta (int)}}.
    base_minima : nº mínimo de casos que qualquer categoria pode ter ao final
        (piso rígido — nunca é furado, mesmo que isso deixe a categoria
        fora da tolerância).

    Retorna
    -------
    dict com:
      - 'ids_excluir': lista de valores de id_var a excluir, na ordem em
        que foram selecionados.
      - 'resumo': DataFrame com Variável, Categoria, Meta, Tolerância,
        Limite inferior/superior, Antes, Depois, Status.
    """
    variaveis = list(metas.keys())
    restante = dados.dropna(subset=[id_var]).copy()
    contagem_original = {v: restante[v].value_counts(dropna=True).to_dict() for v in variaveis}

    excluidos = []

    while True:
        atuais = {v: restante[v].value_counts(dropna=True).to_dict() for v in variaveis}

        limite_inferior = {}
        limite_superior = {}
        necessidade_remocao = {}
        protegidas = {}

        for v in variaveis:
            limite_inferior[v] = {}
            limite_superior[v] = {}
            necessidade_remocao[v] = {}
            protegidas[v] = {}
            for c, meta_n in metas[v].items():
                tol = tolerancia_n[v][c]
                li = max(base_minima, meta_n - tol)
                ls = meta_n + tol
                atual = atuais[v].get(c, 0)
                limite_inferior[v][c] = li
                limite_superior[v][c] = ls
                necessidade_remocao[v][c] = max(0, atual - ls)
                protegidas[v][c] = atual <= li

        if all(necessidade_remocao[v][c] == 0 for v in variaveis for c in metas[v]):
            break  # tudo dentro da tolerância — nada mais a fazer

        # quanto cada categoria ainda pode ceder nesta rodada sem furar o piso
        capacidade = {
            v: {c: max(0, atuais[v].get(c, 0) - limite_inferior[v][c]) for c in metas[v]}
            for v in variaveis
        }

        # protege quem pertence a QUALQUER categoria já no piso/abaixo da meta
        mascaras_protegido = []
        for v in variaveis:
            mascaras_protegido.append(restante[v].map(protegidas[v]).fillna(False))
        protegido = pd.concat(mascaras_protegido, axis=1).any(axis=1)
        candidatos = restante[~protegido].copy()

        if candidatos.empty:
            break  # ninguém pode ser removido sem violar uma categoria protegida

        # pontuação: em quantas variáveis com excesso o candidato está
        colunas_pontuacao = []
        for v in variaveis:
            necessidade_map = {c: (1 if necessidade_remocao[v][c] > 0 else 0) for c in metas[v]}
            colunas_pontuacao.append(candidatos[v].map(necessidade_map).fillna(0))
        candidatos["_pontuacao"] = pd.concat(colunas_pontuacao, axis=1).sum(axis=1)
        candidatos = candidatos[candidatos["_pontuacao"] > 0]

        if candidatos.empty:
            break  # excesso remanescente só existe em categorias protegidas — não dá pra corrigir

        candidatos = candidatos.sort_values("_pontuacao", ascending=False)

        selecionados = []
        for _, row in candidatos.iterrows():
            pode_remover = all(capacidade[v].get(row[v], 0) > 0 for v in variaveis)
            if pode_remover:
                for v in variaveis:
                    capacidade[v][row[v]] -= 1
                selecionados.append(row[id_var])

        if not selecionados:
            break  # nenhum candidato pôde ser removido com segurança nesta rodada

        excluidos.extend(selecionados)
        restante = restante[~restante[id_var].isin(selecionados)]

    finais = {v: restante[v].value_counts(dropna=True).to_dict() for v in variaveis}
    total_antes = {v: sum(contagem_original[v].values()) for v in variaveis}
    total_depois = {v: sum(finais[v].values()) for v in variaveis}

    linhas_resumo = []
    for v in variaveis:
        for c, meta_n in metas[v].items():
            tol = tolerancia_n[v][c]
            li = max(base_minima, meta_n - tol)
            ls = meta_n + tol
            antes = contagem_original[v].get(c, 0)
            depois = finais[v].get(c, 0)

            if li <= depois <= ls:
                status = "Dentro da tolerância"
            elif depois < li:
                status = "Abaixo da meta (exclusão não resolve — precisa de mais casos)"
            else:
                status = "Ainda acima da tolerância (limitado pela base mínima ou por outra variável)"

            pct_meta = round(100 * meta_n / total_antes[v], 2) if total_antes[v] else 0.0
            pct_antes = round(100 * antes / total_antes[v], 2) if total_antes[v] else 0.0
            pct_depois = round(100 * depois / total_depois[v], 2) if total_depois[v] else 0.0

            linhas_resumo.append({
                "Variável": v,
                "Categoria": c,
                "Meta (%)": pct_meta,
                "Meta (N)": meta_n,
                "Tolerância (±)": tol,
                "Limite inferior": li,
                "Limite superior": ls,
                "Antes (%)": pct_antes,
                "Antes (N)": antes,
                "Depois (%)": pct_depois,
                "Depois (N)": depois,
                "Status": status,
            })

    resumo = pd.DataFrame(linhas_resumo)
    return {"ids_excluir": excluidos, "resumo": resumo}
