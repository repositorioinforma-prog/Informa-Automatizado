"""Interface Streamlit do Gerador de Amostra.

O SQLite integrado é a fonte recomendada para município e distrito. As bases
CSV anteriores continuam disponíveis como fallback. Projetos de amostra podem
ser salvos e reabertos em JSON; o JSON guarda apenas parâmetros e regras, nunca
replica o banco demográfico.

O recorte territorial pode ser feito em qualquer ponto da hierarquia: região
intermediária, região imediata, município ou distrito. Seleções de níveis
inferiores funcionam diretamente, sem obrigar o usuário a selecionar os pais.
"""

from __future__ import annotations

import io
import os
import tempfile
import pandas as pd
import streamlit as st

from core.amostra_camadas import (
    ORDEM_NIVEIS,
    ROTULOS_NIVEIS,
    ResultadoAmostraCamadas,
    aplicar_alocacao_manual_camadas,
    calcular_amostra_camadas,
)
from core.amostra_exportador import exportar_amostra
from core.amostra_alocacao import (
    AlocacaoManualErro,
    alocacoes_para_json,
    aplicar_alocacao_manual,
    coluna_codigo_nivel,
    preparar_tabela_alocacao_manual,
)
from core.amostra_math import calcular_amostra, _carregar_distritos, _carregar_municipios
from core.banco_amostras import BancoAmostrasErro, carregar_base_motor_amostragem
from core.projetos_amostra import (
    ProjetoAmostraErro,
    criar_configuracao_projeto,
    desserializar_configuracao_projeto,
    nome_arquivo_projeto,
    serializar_configuracao_projeto,
)

FONTES_DADOS_UI = {
    "Banco integrado (SQLite — recomendado)": "banco_sqlite",
    "Bases anteriores em CSV (fallback)": "bases_legadas",
}
ROTULOS_FONTES_DADOS = {valor: chave for chave, valor in FONTES_DADOS_UI.items()}

FAIXA_FLEX_PADRAO_PCT = 20

METODOS_DISTRIBUICAO_UI = {
    "Proporcional à população": "proporcional",
    "Metas manuais por unidade": "manual",
}
ROTULOS_METODOS_DISTRIBUICAO = {valor: chave for chave, valor in METODOS_DISTRIBUICAO_UI.items()}

FORMATOS_COTAS_UI = {
    "Somente totais (n)": "totais",
    "Somente percentuais (%)": "percentuais",
    "Totais e percentuais": "ambos",
}
ROTULOS_FORMATOS_COTAS = {valor: chave for chave, valor in FORMATOS_COTAS_UI.items()}

NIVEIS_QUEBRA_MAP = {
    "Região Intermediária": "regiao_intermediaria",
    "Região Imediata": "regiao_imediata",
    "Município": "municipio",
    "Distrito": "distrito",
}
ROTULOS_NIVEIS_CLASSICO = {valor: chave for chave, valor in NIVEIS_QUEBRA_MAP.items()}

CHAVES_CAMADAS_UI = {
    "uf": "ga_camada_uf",
    "regiao_intermediaria": "ga_camada_ri",
    "regiao_imediata": "ga_camada_rim",
    "municipio": "ga_camada_mun",
    "distrito": "ga_camada_dist",
}

CHAVES_FILTROS_UI = {
    "regioes_intermediarias": "ga_filtro_ri_codigos",
    "regioes_imediatas": "ga_filtro_rim_codigos",
    "municipios": "ga_filtro_municipios_codigos",
    "distritos": "ga_filtro_distritos_codigos",
}

CONFIG_NIVEIS_FILTRO = {
    "regiao_intermediaria": {
        "codigo": "cod_regiao_intermediaria",
        "nome": "regiao_intermediaria",
    },
    "regiao_imediata": {
        "codigo": "cod_regiao_imediata",
        "nome": "regiao_imediata",
    },
    "municipio": {
        "codigo": "cod_municipio",
        "nome": "municipio",
    },
    "distrito": {
        "codigo": "cod_distrito",
        "nome": "distrito",
    },
}


def _inicializar_estado() -> None:
    padroes = {
        "ga_resultado": None,
        "ga_params": None,
        "ga_campanhas": None,
        "ga_camadas_resumos": None,
        "ga_validacoes": None,
        "ga_modo_resultado": "Clássico",
        "ga_mensagem_projeto": None,
        "ga_filtro_ri_codigos": [],
        "ga_filtro_rim_codigos": [],
        "ga_filtro_municipios_codigos": [],
        "ga_filtro_distritos_codigos": [],
        "ga_habilitar_filtro_distritos": False,
        "ga_resultado_base_manual": None,
        "ga_calculo_base_manual": None,
        "ga_alocacoes_manuais_salvas": [],
        "ga_manual_signature": None,
        "ga_faixa_flex_pct": FAIXA_FLEX_PADRAO_PCT,
    }
    for chave, valor in padroes.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def _executar_rerun() -> None:
    if hasattr(st, "rerun"):
        st.rerun()
    else:  # compatibilidade com versões antigas do Streamlit
        st.experimental_rerun()


def _limpar_recorte_especifico() -> None:
    for chave in CHAVES_FILTROS_UI.values():
        st.session_state[chave] = []


def _aplicar_projeto_no_estado(configuracao: dict) -> None:
    """Preenche os widgets antes de eles serem instanciados no novo rerun."""
    modo_ui = "Clássico" if configuracao["modo"] == "classico" else "Por camadas (Beta)"
    territorio = configuracao["territorio"]

    st.session_state["ga_nome_projeto"] = configuracao["nome"]
    st.session_state["ga_fonte_dados_label"] = ROTULOS_FONTES_DADOS[configuracao["fonte_dados"]]
    st.session_state["ga_modo_montagem"] = modo_ui
    st.session_state["ga_uf_opcao"] = territorio["uf"] or "(Brasil inteiro)"
    for campo, chave_widget in CHAVES_FILTROS_UI.items():
        st.session_state[chave_widget] = [
            str(item["codigo"]) for item in territorio.get(campo, [])
        ]

    st.session_state["ga_amostra_total"] = configuracao["amostra_total"]
    st.session_state["ga_base_populacional_label"] = (
        "16 anos ou mais (idade da pesquisa)"
        if configuracao["base_populacional"] == "16_mais"
        else "Total absoluto (todas as idades)"
    )
    # A banda mínimo/máximo continua existindo internamente, mas deixou de ser
    # um parâmetro visível porque não altera a distribuição principal.
    st.session_state["ga_faixa_flex_pct"] = FAIXA_FLEX_PADRAO_PCT
    distribuicao = configuracao.get("distribuicao", {})
    metodo = distribuicao.get("metodo", "proporcional")
    st.session_state["ga_metodo_distribuicao_label"] = ROTULOS_METODOS_DISTRIBUICAO[metodo]
    st.session_state["ga_alocacoes_manuais_salvas"] = list(
        distribuicao.get("alocacoes_manuais", [])
    )
    exibicao = configuracao.get("exportacao", {}).get("exibicao_cotas", "ambos")
    st.session_state["ga_exibicao_cotas_label"] = ROTULOS_FORMATOS_COTAS[exibicao]

    classico = configuracao["classico"]
    if classico["nivel_quebra"]:
        st.session_state["ga_nivel_quebra_label"] = ROTULOS_NIVEIS_CLASSICO[
            classico["nivel_quebra"]
        ]

    camadas = set(configuracao["camadas"])
    for nivel, chave_widget in CHAVES_CAMADAS_UI.items():
        st.session_state[chave_widget] = nivel in camadas

    st.session_state["ga_habilitar_filtro_distritos"] = bool(
        classico["incluir_distritos"]
        or territorio.get("distritos")
        or "distrito" in camadas
    )

    # O projeto restaura parâmetros, mas não reutiliza um resultado calculado
    # com outra fonte ou versão de banco.
    st.session_state["ga_resultado"] = None
    st.session_state["ga_params"] = None
    st.session_state["ga_campanhas"] = None
    st.session_state["ga_camadas_resumos"] = None
    st.session_state["ga_validacoes"] = None
    st.session_state["ga_resultado_base_manual"] = None
    st.session_state["ga_calculo_base_manual"] = None
    st.session_state["ga_manual_signature"] = None
    st.session_state.pop("ga_editor_amostra_manual", None)


def _interface_abrir_projeto() -> None:
    with st.expander("Abrir projeto de amostra (.json)", expanded=False):
        st.caption(
            "O arquivo JSON contém somente as escolhas do desenho amostral. "
            "Os dados demográficos continuam no banco SQLite do aplicativo."
        )
        arquivo = st.file_uploader(
            "Selecione um projeto",
            type=["json"],
            key="ga_arquivo_projeto_json",
        )
        carregar = st.button(
            "Carregar parâmetros do projeto",
            use_container_width=True,
            key="ga_carregar_projeto_json",
        )
        if carregar:
            if arquivo is None:
                st.warning("Selecione um arquivo JSON antes de carregar.")
                return
            try:
                configuracao = desserializar_configuracao_projeto(arquivo.getvalue())
                _aplicar_projeto_no_estado(configuracao)
                st.session_state["ga_mensagem_projeto"] = (
                    f"Projeto “{configuracao['nome']}” carregado. "
                    "Revise os parâmetros e gere novamente a amostra."
                )
                _executar_rerun()
            except ProjetoAmostraErro as exc:
                st.error(f"Projeto inválido: {exc}")


def _carregar_base_fonte(fonte_dados: str, nivel: str) -> pd.DataFrame:
    if fonte_dados == "banco_sqlite":
        return carregar_base_motor_amostragem(nivel)
    if fonte_dados == "bases_legadas":
        return _carregar_distritos() if nivel == "distrito" else _carregar_municipios()
    raise ValueError(f"Fonte de dados desconhecida: {fonte_dados}")


def _filtrar_dataframe(
    df: pd.DataFrame,
    *,
    uf: str | None = None,
    codigos_ri: list[str] | None = None,
    codigos_rim: list[str] | None = None,
    codigos_municipios: list[str] | None = None,
    codigos_distritos: list[str] | None = None,
) -> pd.DataFrame:
    """Aplica o mesmo recorte hierárquico usado pelos motores."""
    recorte = df
    if uf:
        recorte = recorte[recorte["uf"].astype(str) == str(uf)]

    filtros = (
        ("cod_regiao_intermediaria", codigos_ri),
        ("cod_regiao_imediata", codigos_rim),
        ("cod_municipio", codigos_municipios),
        ("cod_distrito", codigos_distritos),
    )
    for coluna, codigos in filtros:
        if not codigos:
            continue
        if coluna not in recorte.columns:
            raise ValueError(f"A base selecionada não contém a coluna {coluna}.")
        codigos_texto = [str(codigo) for codigo in codigos]
        recorte = recorte[recorte[coluna].astype("string").isin(codigos_texto)]
    return recorte


def _texto(valor: object) -> str:
    return "" if pd.isna(valor) else str(valor)


def _montar_opcoes_territoriais(
    df: pd.DataFrame,
    nivel: str,
) -> tuple[list[str], dict[str, str]]:
    """Monta códigos como valores e rótulos contextuais para os multiselects."""
    config = CONFIG_NIVEIS_FILTRO[nivel]
    codigo = config["codigo"]
    nome = config["nome"]
    if codigo not in df.columns or nome not in df.columns:
        return [], {}

    colunas = []
    for coluna in [
        "uf",
        "regiao_intermediaria",
        "regiao_imediata",
        "municipio",
        codigo,
        nome,
    ]:
        if coluna in df.columns and coluna not in colunas:
            colunas.append(coluna)
    tabela = df[colunas].dropna(subset=[codigo, nome]).copy()
    tabela[codigo] = tabela[codigo].astype(str)
    tabela = tabela.drop_duplicates(subset=[codigo], keep="first")

    def rotulo(linha: pd.Series) -> str:
        uf_linha = _texto(linha.get("uf"))
        nome_linha = _texto(linha.get(nome))
        codigo_linha = _texto(linha.get(codigo))
        if nivel == "regiao_intermediaria":
            return f"{uf_linha} · {nome_linha} ({codigo_linha})"
        if nivel == "regiao_imediata":
            return (
                f"{uf_linha} · {_texto(linha.get('regiao_intermediaria'))} "
                f"→ {nome_linha} ({codigo_linha})"
            )
        if nivel == "municipio":
            return f"{uf_linha} · {nome_linha} ({codigo_linha})"
        return (
            f"{uf_linha} · {_texto(linha.get('municipio'))} "
            f"→ {nome_linha} ({codigo_linha})"
        )

    tabela["_rotulo"] = tabela.apply(rotulo, axis=1)
    tabela = tabela.sort_values(["_rotulo", codigo], kind="stable")
    opcoes = tabela[codigo].tolist()
    rotulos = dict(zip(tabela[codigo], tabela["_rotulo"]))
    return opcoes, rotulos


def _podar_selecao_invalida(chave: str, opcoes: list[str]) -> None:
    """Remove do estado códigos que deixaram de pertencer ao recorte pai."""
    atuais = st.session_state.get(chave, [])
    if not isinstance(atuais, (list, tuple, set)):
        atuais = []
    permitidos = set(opcoes)
    validos = [str(valor) for valor in atuais if str(valor) in permitidos]
    if list(atuais) != validos:
        st.session_state[chave] = validos


def _multiselect_territorial(
    *,
    label: str,
    df: pd.DataFrame,
    nivel: str,
    key: str,
    help_text: str,
) -> list[str]:
    opcoes, rotulos = _montar_opcoes_territoriais(df, nivel)
    _podar_selecao_invalida(key, opcoes)
    return st.multiselect(
        label,
        options=opcoes,
        format_func=lambda codigo: rotulos.get(str(codigo), str(codigo)),
        help=help_text,
        key=key,
    )


def _selecoes_para_json(
    df: pd.DataFrame,
    nivel: str,
    codigos_selecionados: list[str],
) -> list[dict[str, str]]:
    if not codigos_selecionados:
        return []
    config = CONFIG_NIVEIS_FILTRO[nivel]
    codigo = config["codigo"]
    nome = config["nome"]
    codigos = {str(valor) for valor in codigos_selecionados}
    selecionadas = df[df[codigo].astype("string").isin(codigos)][[codigo, nome]].copy()
    selecionadas = selecionadas.dropna(subset=[codigo, nome]).drop_duplicates(subset=[codigo])
    selecionadas[codigo] = selecionadas[codigo].astype(str)
    selecionadas[nome] = selecionadas[nome].astype(str)
    selecionadas = selecionadas.sort_values([nome, codigo], kind="stable")
    return [
        {"codigo": linha[codigo], "nome": linha[nome]}
        for _, linha in selecionadas.iterrows()
    ]


def _descricao_recorte(
    *,
    uf: str | None,
    codigos_ri: list[str],
    codigos_rim: list[str],
    codigos_municipios: list[str],
    codigos_distritos: list[str],
) -> str:
    partes = [uf or "Brasil"]
    if codigos_ri:
        partes.append(f"{len(codigos_ri)} região(ões) intermediária(s)")
    if codigos_rim:
        partes.append(f"{len(codigos_rim)} região(ões) imediata(s)")
    if codigos_municipios:
        partes.append(f"{len(codigos_municipios)} município(s)")
    if codigos_distritos:
        partes.append(f"{len(codigos_distritos)} distrito(s)")
    if len(partes) == 1:
        partes.append("todas as unidades")
    return " · ".join(partes)


def _nivel_mais_profundo(niveis: list[str]) -> str:
    """Retorna o nível territorial mais detalhado da lista."""
    niveis_validos = [nivel for nivel in niveis if nivel in ORDEM_NIVEIS]
    if not niveis_validos:
        raise ValueError("Não foi possível determinar o nível territorial.")
    return max(niveis_validos, key=ORDEM_NIVEIS.index)


def _determinar_nivel_alocacao_manual(
    nivel_saida: str | None,
    *,
    codigos_ri: list[str] | None = None,
    codigos_rim: list[str] | None = None,
    codigos_municipios: list[str] | None = None,
    codigos_distritos: list[str] | None = None,
) -> str:
    """Define em qual unidade o usuário editará a meta bruta.

    A regra é simples: usa o nível mais detalhado entre a saída configurada e
    as seleções territoriais explícitas. Assim, ao selecionar cinco municípios,
    o editor sempre exibe cinco linhas, mesmo que a saída anterior estivesse
    configurada em Região Intermediária.
    """
    candidatos: list[str] = []
    if nivel_saida:
        candidatos.append(nivel_saida)
    if codigos_ri:
        candidatos.append("regiao_intermediaria")
    if codigos_rim:
        candidatos.append("regiao_imediata")
    if codigos_municipios:
        candidatos.append("municipio")
    if codigos_distritos:
        candidatos.append("distrito")
    return _nivel_mais_profundo(candidatos)


def _garantir_nivel_manual_nas_camadas(
    camadas: list[str],
    nivel_manual: str,
) -> list[str]:
    """Garante que o nível editável seja o último nível do cálculo."""
    resultado = list(camadas)
    if not resultado:
        return [nivel_manual]
    ultimo = resultado[-1]
    if ORDEM_NIVEIS.index(nivel_manual) > ORDEM_NIVEIS.index(ultimo):
        resultado.append(nivel_manual)
    return resultado


def _renderizar_distribuicao_manual(params: dict, faixa_flex_pct: int) -> bool:
    """Mostra o editor manual e aplica os valores quando o fechamento é válido."""
    base = st.session_state.get("ga_resultado_base_manual")
    if base is None:
        return False

    nivel = params.get("nivel_alocacao_manual") or params["nivel_quebra"]
    codigo = coluna_codigo_nivel(nivel)
    iniciais = st.session_state.get("ga_alocacoes_manuais_salvas") or []
    tabela = preparar_tabela_alocacao_manual(
        base,
        nivel=nivel,
        alocacoes_iniciais=iniciais,
    )

    assinatura = (
        nivel,
        int(params["amostra_total"]),
        tuple(tabela[codigo].astype(str).tolist()),
    )
    if st.session_state.get("ga_manual_signature") != assinatura:
        st.session_state["ga_manual_signature"] = assinatura
        st.session_state.pop("ga_editor_amostra_manual", None)

    st.subheader("Metas manuais por unidade")
    st.info(
        f"A edição está sendo feita por **{ROTULOS_NIVEIS[nivel]}**. "
        "Informe somente a quantidade bruta de entrevistas de cada unidade. "
        "Depois da aplicação, gênero, idade e renda são redistribuídos "
        "automaticamente dentro de cada meta."
    )
    st.caption(
        "A coluna **Sugestão proporcional** mostra o cálculo baseado na população real. "
        "Altere somente **Meta da unidade (n)**. A soma de todas as linhas precisa fechar "
        "exatamente no total de entrevistas do projeto."
    )
    if len(tabela) > 500:
        st.warning(
            f"O recorte possui {len(tabela)} unidades. A edição continua disponível, "
            "mas pode ficar mais confortável após limitar o território ou selecionar "
            "unidades específicas."
        )

    colunas_ordem = [
        coluna
        for coluna in [
            "uf",
            "regiao_intermediaria",
            "regiao_imediata",
            "municipio",
            "distrito",
            codigo,
            "populacao_total",
            "populacao_base",
            "percentual",
            "amostra_ideal",
            "amostra_proporcional",
            "amostra_manual",
        ]
        if coluna in tabela.columns
    ]
    configuracao_colunas = {
        "uf": st.column_config.TextColumn("UF", disabled=True),
        "regiao_intermediaria": st.column_config.TextColumn(
            "Região Intermediária", disabled=True
        ),
        "regiao_imediata": st.column_config.TextColumn(
            "Região Imediata", disabled=True
        ),
        "municipio": st.column_config.TextColumn("Município", disabled=True),
        "distrito": st.column_config.TextColumn("Distrito", disabled=True),
        codigo: st.column_config.TextColumn("Código", disabled=True),
        "populacao_total": st.column_config.NumberColumn(
            "População total", disabled=True, format="%d"
        ),
        "populacao_base": st.column_config.NumberColumn(
            "População-base", disabled=True, format="%d"
        ),
        "percentual": st.column_config.NumberColumn(
            "% real do recorte", disabled=True, format="%.2f%%"
        ),
        "amostra_ideal": st.column_config.NumberColumn(
            "Amostra ideal", disabled=True, format="%.3f"
        ),
        "amostra_proporcional": st.column_config.NumberColumn(
            "Sugestão proporcional", disabled=True, format="%d"
        ),
        "amostra_manual": st.column_config.NumberColumn(
            "Meta da unidade (n)", min_value=0, step=1, format="%d", required=True
        ),
    }
    editada = st.data_editor(
        tabela,
        column_order=colunas_ordem,
        column_config=configuracao_colunas,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="ga_editor_amostra_manual",
    )

    valores = pd.to_numeric(editada["amostra_manual"], errors="coerce")
    invalidos = bool(
        valores.isna().any()
        or (valores < 0).any()
        or ((valores % 1).abs() > 1e-9).any()
    )
    soma = int(valores.fillna(0).sum()) if not invalidos else 0
    meta = int(params["amostra_total"] )
    diferenca = meta - soma

    c1, c2, c3 = st.columns(3)
    c1.metric("Total geral", meta)
    c2.metric("Soma preenchida", soma)
    if invalidos:
        c3.metric("Situação", "Valor inválido")
    elif diferenca > 0:
        c3.metric("Faltam", diferenca)
    elif diferenca < 0:
        c3.metric("Excedente", -diferenca)
    else:
        c3.metric("Fechamento", "OK")

    if invalidos:
        st.error("Use somente números inteiros não negativos na coluna Meta da unidade (n).")
    elif diferenca != 0:
        acao = "faltam" if diferenca > 0 else "sobram"
        st.warning(
            f"A soma ainda não fecha: {acao} {abs(diferenca)} entrevista(s) "
            f"para chegar a {meta}."
        )
    else:
        st.success("As metas das unidades fecham exatamente no total geral da amostra.")

    aplicar = st.button(
        "✅ Aplicar metas e recalcular segmentos",
        type="primary",
        use_container_width=True,
        disabled=invalidos or diferenca != 0,
        key="ga_aplicar_amostra_manual",
    )
    if aplicar:
        try:
            alocacoes = alocacoes_para_json(editada, nivel=nivel)
            calculo_base = st.session_state.get("ga_calculo_base_manual")
            if isinstance(calculo_base, ResultadoAmostraCamadas):
                calculo_manual = aplicar_alocacao_manual_camadas(
                    calculo_base,
                    alocacoes_manuais=alocacoes,
                    amostra_total=meta,
                    faixa_flex=faixa_flex_pct / 100,
                )
                st.session_state.ga_resultado = calculo_manual.resultado
                st.session_state.ga_camadas_resumos = calculo_manual.resumos
                st.session_state.ga_validacoes = calculo_manual.validacoes
            else:
                st.session_state.ga_resultado = aplicar_alocacao_manual(
                    base,
                    nivel=nivel,
                    alocacoes_manuais=alocacoes,
                    amostra_total=meta,
                    faixa_flex=faixa_flex_pct / 100,
                )
                st.session_state.ga_camadas_resumos = None
                st.session_state.ga_validacoes = None
            st.session_state.ga_alocacoes_manuais_salvas = alocacoes
            st.session_state.ga_params["metodo_distribuicao"] = "manual"
            st.session_state.ga_params["metodo_distribuicao_label"] = "Metas manuais por unidade"
            st.session_state.ga_campanhas = None
            st.success("Metas aplicadas e segmentos internos recalculados.")
        except (AlocacaoManualErro, ValueError) as exc:
            st.error(f"Não foi possível aplicar a distribuição manual: {exc}")

    return st.session_state.ga_resultado is not None


def modulo_gerador_amostra():
    _inicializar_estado()

    st.header("Gerador de Amostra")
    st.caption(
        "Monte amostras territoriais com fechamento proporcional por unidade. "
        "Você pode recortar diretamente regiões, municípios ou distritos específicos."
    )

    if st.session_state.get("ga_mensagem_projeto"):
        st.success(st.session_state.pop("ga_mensagem_projeto"))

    _interface_abrir_projeto()

    # -----------------------------------------------------------------
    # Sidebar — parâmetros
    # -----------------------------------------------------------------
    with st.sidebar:
        st.header("Parâmetros da pesquisa")

        nome_projeto = st.text_input(
            "Nome do projeto",
            value="Projeto de amostra",
            key="ga_nome_projeto",
        )

        fonte_dados_label = st.selectbox(
            "Fonte dos dados territoriais",
            options=list(FONTES_DADOS_UI.keys()),
            index=0,
            help=(
                "O SQLite integrado preserva fontes, cobertura e valores ausentes. "
                "Os CSVs anteriores ficam disponíveis somente como contingência."
            ),
            key="ga_fonte_dados_label",
        )
        fonte_dados = FONTES_DADOS_UI[fonte_dados_label]

        modo_montagem = st.radio(
            "Modo de montagem",
            options=["Clássico", "Por camadas (Beta)"],
            horizontal=True,
            help=(
                "O modo clássico faz uma quebra. Por camadas distribui "
                "recursivamente e garante o fechamento entre pais e filhos."
            ),
            key="ga_modo_montagem",
        )

        st.markdown("**Estrutura da saída**")
        camadas_selecionadas: list[str] = []
        nivel_quebra_label = None
        nivel_quebra_atual = None

        if modo_montagem == "Clássico":
            nivel_quebra_label = st.selectbox(
                "Nível de quebra da amostra",
                options=list(NIVEIS_QUEBRA_MAP.keys()),
                index=0,
                key="ga_nivel_quebra_label",
            )
            nivel_quebra_atual = NIVEIS_QUEBRA_MAP[nivel_quebra_label]
        else:
            st.caption(
                "Nesta versão, as camadas usam distribuição proporcional e "
                "fecham recursivamente."
            )
            if st.checkbox("1. Estado (UF)", value=True, key="ga_camada_uf"):
                camadas_selecionadas.append("uf")
            if st.checkbox(
                "2. Região Intermediária",
                value=True,
                key="ga_camada_ri",
            ):
                camadas_selecionadas.append("regiao_intermediaria")
            if st.checkbox(
                "3. Região Imediata",
                value=False,
                key="ga_camada_rim",
            ):
                camadas_selecionadas.append("regiao_imediata")
            if st.checkbox("4. Município", value=True, key="ga_camada_mun"):
                camadas_selecionadas.append("municipio")
            if st.checkbox("5. Distrito", value=False, key="ga_camada_dist"):
                camadas_selecionadas.append("distrito")

        distrito_obrigatorio = (
            nivel_quebra_atual == "distrito" or "distrito" in camadas_selecionadas
        )
        if distrito_obrigatorio:
            st.session_state["ga_habilitar_filtro_distritos"] = True
            usar_base_distrital = True
            st.caption("Base distrital ativada automaticamente pela estrutura da saída.")
        else:
            usar_base_distrital = st.toggle(
                "Habilitar seleção de distritos específicos",
                value=False,
                help=(
                    "Ative para recortar um ou mais distritos mesmo que a saída final "
                    "seja agregada em município ou região."
                ),
                key="ga_habilitar_filtro_distritos",
            )
        incluir_distritos = bool(usar_base_distrital)
        nivel_base_filtros = "distrito" if usar_base_distrital else "municipio"

        try:
            df_base_filtros = _carregar_base_fonte(fonte_dados, nivel_base_filtros)
        except (BancoAmostrasErro, OSError, ValueError) as exc:
            st.error(f"Não foi possível carregar a fonte selecionada: {exc}")
            st.stop()

        st.markdown("**Recorte territorial**")
        st.caption(
            "Os filtros são opcionais e combinados. Você pode escolher municípios ou "
            "distritos diretamente, sem selecionar antes as regiões correspondentes."
        )

        ufs_disponiveis = sorted(
            df_base_filtros["uf"].dropna().astype(str).unique().tolist()
        )
        opcoes_uf = ["(Brasil inteiro)"] + ufs_disponiveis
        uf_estado = st.session_state.get("ga_uf_opcao")
        if uf_estado not in opcoes_uf:
            st.session_state["ga_uf_opcao"] = (
                "SP" if "SP" in ufs_disponiveis else opcoes_uf[0]
            )

        indice_uf = opcoes_uf.index("SP") if "SP" in opcoes_uf else 0
        uf_opcao = st.selectbox(
            "Território inicial",
            options=opcoes_uf,
            index=indice_uf,
            key="ga_uf_opcao",
        )
        uf = None if uf_opcao == "(Brasil inteiro)" else uf_opcao

        df_apos_uf = _filtrar_dataframe(df_base_filtros, uf=uf)
        codigos_ri = _multiselect_territorial(
            label="Regiões Intermediárias específicas (opcional)",
            df=df_apos_uf,
            nivel="regiao_intermediaria",
            key="ga_filtro_ri_codigos",
            help_text="Vazio significa todas as regiões dentro do território inicial.",
        )

        df_apos_ri = _filtrar_dataframe(df_apos_uf, codigos_ri=codigos_ri)
        codigos_rim = _multiselect_territorial(
            label="Regiões Imediatas específicas (opcional)",
            df=df_apos_ri,
            nivel="regiao_imediata",
            key="ga_filtro_rim_codigos",
            help_text=(
                "Pode ser usada diretamente. Quando regiões intermediárias também forem "
                "selecionadas, vale somente dentro delas."
            ),
        )

        df_apos_rim = _filtrar_dataframe(df_apos_ri, codigos_rim=codigos_rim)
        codigos_municipios = _multiselect_territorial(
            label="Municípios específicos (opcional)",
            df=df_apos_rim,
            nivel="municipio",
            key="ga_filtro_municipios_codigos",
            help_text=(
                "Digite parte do nome ou do código. Vazio inclui todos os municípios "
                "do recorte acima."
            ),
        )

        df_apos_municipios = _filtrar_dataframe(
            df_apos_rim,
            codigos_municipios=codigos_municipios,
        )
        codigos_distritos: list[str] = []
        if usar_base_distrital:
            codigos_distritos = _multiselect_territorial(
                label="Distritos específicos (opcional)",
                df=df_apos_municipios,
                nivel="distrito",
                key="ga_filtro_distritos_codigos",
                help_text=(
                    "Digite o distrito ou o município. Vazio inclui todos os distritos "
                    "do recorte acima."
                ),
            )
        else:
            st.session_state["ga_filtro_distritos_codigos"] = []

        df_recorte = _filtrar_dataframe(
            df_apos_municipios,
            codigos_distritos=codigos_distritos,
        )
        if df_recorte.empty:
            st.error("O recorte atual não contém unidades territoriais.")
        else:
            quant_municipios = int(df_recorte["cod_municipio"].nunique())
            resumo = f"{quant_municipios} município(s) disponível(is)"
            if usar_base_distrital and "cod_distrito" in df_recorte.columns:
                resumo += f" · {int(df_recorte['cod_distrito'].nunique())} distrito(s)"
            st.info(resumo)

        st.button(
            "Limpar seleções específicas",
            on_click=_limpar_recorte_especifico,
            use_container_width=True,
            key="ga_limpar_recorte",
        )

        if modo_montagem != "Clássico":
            raiz = "Brasil" if uf is None else uf
            fluxo = [raiz] + [ROTULOS_NIVEIS[c] for c in camadas_selecionadas]
            st.info("Fluxo: " + " → ".join(fluxo))

        amostra_total = st.number_input(
            "Total geral da amostra",
            min_value=1,
            value=1000,
            step=100,
            help=(
                "É a soma das metas de todas as unidades. Ex.: ao selecionar cinco "
                "municípios e informar 500, as cinco metas municipais devem somar 500."
            ),
            key="ga_amostra_total",
        )

        metodo_distribuicao_label = st.radio(
            "Distribuição do total",
            options=list(METODOS_DISTRIBUICAO_UI.keys()),
            index=0,
            help=(
                "Proporcional usa a população-base. Metas manuais exibem uma linha para "
                "cada unidade selecionada e permitem informar somente a quantidade bruta. "
                "Gênero, idade e renda são recalculados automaticamente."
            ),
            key="ga_metodo_distribuicao_label",
        )
        metodo_distribuicao = METODOS_DISTRIBUICAO_UI[metodo_distribuicao_label]

        nivel_saida_configurado = (
            nivel_quebra_atual
            if modo_montagem == "Clássico"
            else (camadas_selecionadas[-1] if camadas_selecionadas else None)
        )
        nivel_alocacao_manual = None
        if metodo_distribuicao == "manual" and nivel_saida_configurado:
            nivel_alocacao_manual = _determinar_nivel_alocacao_manual(
                nivel_saida_configurado,
                codigos_ri=codigos_ri,
                codigos_rim=codigos_rim,
                codigos_municipios=codigos_municipios,
                codigos_distritos=codigos_distritos,
            )
            st.info(
                f"As metas serão editadas por **{ROTULOS_NIVEIS[nivel_alocacao_manual]}**. "
                "Cada linha receberá uma quantidade bruta e a soma deverá fechar no "
                f"total de {int(amostra_total)} entrevistas."
            )

        base_populacional_label = st.radio(
            "Base populacional",
            options=[
                "16 anos ou mais (idade da pesquisa)",
                "Total absoluto (todas as idades)",
            ],
            index=0,
            key="ga_base_populacional_label",
        )

        # A banda mínimo/máximo permanece no cálculo e no relatório técnico,
        # mas não aparece como controle na interface porque não altera a
        # quantidade principal distribuída.
        faixa_flex_pct = FAIXA_FLEX_PADRAO_PCT
        st.session_state["ga_faixa_flex_pct"] = faixa_flex_pct

        st.markdown("**Preferências do Excel**")
        exibicao_cotas_label = st.radio(
            "Campos de gênero, idade e renda",
            options=list(FORMATOS_COTAS_UI.keys()),
            index=2,
            help=(
                "Amostra total e % da amostra aparecem sempre. Esta preferência altera "
                "somente as categorias que compõem gênero, idade e renda."
            ),
            key="ga_exibicao_cotas_label",
        )
        exibicao_cotas = FORMATOS_COTAS_UI[exibicao_cotas_label]

        texto_botao = (
            "🎯 Gerar amostra proporcional"
            if metodo_distribuicao == "proporcional"
            else "🧮 Definir metas por unidade"
        )
        submitted = st.button(
            texto_botao,
            type="primary",
            use_container_width=True,
            key="ga_gerar_amostra",
        )

    # -----------------------------------------------------------------
    # Salvar projeto em JSON — disponível mesmo antes do cálculo
    # -----------------------------------------------------------------
    base_populacional = "16_mais" if base_populacional_label.startswith("16") else "total"
    modo_interno = "classico" if modo_montagem == "Clássico" else "camadas"

    regioes_json = _selecoes_para_json(
        df_base_filtros,
        "regiao_intermediaria",
        codigos_ri,
    )
    regioes_imediatas_json = _selecoes_para_json(
        df_base_filtros,
        "regiao_imediata",
        codigos_rim,
    )
    municipios_json = _selecoes_para_json(
        df_base_filtros,
        "municipio",
        codigos_municipios,
    )
    distritos_json = _selecoes_para_json(
        df_base_filtros,
        "distrito",
        codigos_distritos,
    ) if usar_base_distrital else []

    with st.expander("Salvar configuração do projeto", expanded=False):
        try:
            configuracao_atual = criar_configuracao_projeto(
                nome=nome_projeto,
                fonte_dados=fonte_dados,
                modo=modo_interno,
                amostra_total=int(amostra_total),
                base_populacional=base_populacional,
                faixa_flex_pct=int(faixa_flex_pct),
                uf=uf,
                regioes_intermediarias=regioes_json,
                regioes_imediatas=regioes_imediatas_json,
                municipios=municipios_json,
                distritos=distritos_json,
                nivel_quebra=nivel_quebra_atual,
                incluir_distritos=incluir_distritos,
                camadas=camadas_selecionadas,
                metodo_distribuicao=metodo_distribuicao,
                alocacoes_manuais=(
                    st.session_state.get("ga_alocacoes_manuais_salvas", [])
                    if metodo_distribuicao == "manual"
                    else []
                ),
                exibicao_cotas=exibicao_cotas,
            )
            st.download_button(
                "⬇️ Baixar projeto em JSON",
                data=serializar_configuracao_projeto(configuracao_atual),
                file_name=nome_arquivo_projeto(nome_projeto),
                mime="application/json",
                use_container_width=True,
                key="ga_download_projeto_json",
            )
            st.caption(
                "O JSON guarda o recorte completo por códigos IBGE, mas não contém "
                "a população nem as tabelas do banco."
            )
        except ProjetoAmostraErro as exc:
            st.info(f"Complete a configuração antes de salvar: {exc}")

    # -----------------------------------------------------------------
    # Cálculo
    # -----------------------------------------------------------------
    if submitted:
        try:
            if df_recorte.empty:
                raise ValueError("O recorte atual não contém unidades territoriais.")

            calculo_camadas = None
            nivel_saida_solicitado = nivel_saida_configurado
            nivel_manual_efetivo = (
                nivel_alocacao_manual
                if metodo_distribuicao == "manual"
                else nivel_saida_configurado
            )
            if not nivel_manual_efetivo:
                raise ValueError("Não foi possível determinar o nível territorial da amostra.")

            if modo_montagem == "Clássico":
                # Quando existem unidades específicas, a edição manual ocorre no
                # nível mais detalhado selecionado. Ex.: cinco municípios geram
                # cinco linhas, mesmo que a quebra anterior estivesse em região.
                nivel_quebra = nivel_manual_efetivo
                nivel_quebra_label = ROTULOS_NIVEIS[nivel_quebra]
                resultado_base = calcular_amostra(
                    uf=uf,
                    codigos_regioes_intermediarias=codigos_ri or None,
                    codigos_regioes_imediatas=codigos_rim or None,
                    codigos_municipios=codigos_municipios or None,
                    codigos_distritos=codigos_distritos or None,
                    amostra_total=int(amostra_total),
                    nivel_quebra=nivel_quebra,
                    base_populacional=base_populacional,
                    incluir_distritos=incluir_distritos,
                    faixa_flex=faixa_flex_pct / 100,
                    df_base=df_base_filtros,
                )
                camadas_calculo: list[str] = []
            else:
                if not camadas_selecionadas:
                    raise ValueError("Selecione pelo menos uma camada territorial.")
                camadas_calculo = list(camadas_selecionadas)
                if metodo_distribuicao == "manual":
                    camadas_calculo = _garantir_nivel_manual_nas_camadas(
                        camadas_calculo,
                        nivel_manual_efetivo,
                    )
                calculo_camadas = calcular_amostra_camadas(
                    camadas=camadas_calculo,
                    uf=uf,
                    codigos_regioes_intermediarias=codigos_ri or None,
                    codigos_regioes_imediatas=codigos_rim or None,
                    codigos_municipios=codigos_municipios or None,
                    codigos_distritos=codigos_distritos or None,
                    amostra_total=int(amostra_total),
                    base_populacional=base_populacional,
                    faixa_flex=faixa_flex_pct / 100,
                    df_base=df_base_filtros,
                )
                resultado_base = calculo_camadas.resultado
                nivel_quebra = camadas_calculo[-1]
                nivel_quebra_label = ROTULOS_NIVEIS[nivel_quebra]

            recorte_descricao = _descricao_recorte(
                uf=uf,
                codigos_ri=codigos_ri,
                codigos_rim=codigos_rim,
                codigos_municipios=codigos_municipios,
                codigos_distritos=codigos_distritos,
            )
            st.session_state.ga_params = {
                "nome_projeto": nome_projeto,
                "uf": uf,
                "amostra_total": int(amostra_total),
                "nivel_quebra": nivel_quebra,
                "nivel_quebra_label": nivel_quebra_label,
                "nivel_saida_solicitado": nivel_saida_solicitado,
                "nivel_saida_solicitado_label": (
                    ROTULOS_NIVEIS[nivel_saida_solicitado]
                    if nivel_saida_solicitado
                    else None
                ),
                "nivel_alocacao_manual": (
                    nivel_quebra if metodo_distribuicao == "manual" else None
                ),
                "nivel_alocacao_manual_label": (
                    ROTULOS_NIVEIS[nivel_quebra]
                    if metodo_distribuicao == "manual"
                    else None
                ),
                "base_populacional": base_populacional,
                "modo": modo_montagem,
                "camadas": camadas_calculo if modo_montagem != "Clássico" else [],
                "fonte_dados": fonte_dados,
                "fonte_dados_label": fonte_dados_label,
                "recorte_descricao": recorte_descricao,
                "metodo_distribuicao": metodo_distribuicao,
                "metodo_distribuicao_label": metodo_distribuicao_label,
                "exibicao_cotas": exibicao_cotas,
                "exibicao_cotas_label": exibicao_cotas_label,
                "faixa_flex_pct": int(faixa_flex_pct),
            }
            st.session_state.ga_modo_resultado = modo_montagem
            st.session_state.ga_campanhas = None

            if metodo_distribuicao == "manual":
                st.session_state.ga_resultado_base_manual = resultado_base
                st.session_state.ga_calculo_base_manual = calculo_camadas
                st.session_state.ga_resultado = None
                st.session_state.ga_camadas_resumos = None
                st.session_state.ga_validacoes = None
                st.session_state.ga_manual_signature = None
                st.session_state.pop("ga_editor_amostra_manual", None)
            else:
                st.session_state.ga_resultado = resultado_base
                st.session_state.ga_resultado_base_manual = None
                st.session_state.ga_calculo_base_manual = None
                st.session_state.ga_alocacoes_manuais_salvas = []
                st.session_state.ga_manual_signature = None
                if calculo_camadas is not None:
                    st.session_state.ga_camadas_resumos = calculo_camadas.resumos
                    st.session_state.ga_validacoes = calculo_camadas.validacoes
                else:
                    st.session_state.ga_camadas_resumos = None
                    st.session_state.ga_validacoes = None
        except (ValueError, BancoAmostrasErro, AlocacaoManualErro) as exc:
            st.error(f"Erro no cálculo: {exc}")

    params_preparados = st.session_state.get("ga_params")
    if (
        params_preparados
        and params_preparados.get("metodo_distribuicao") == "manual"
        and st.session_state.get("ga_resultado_base_manual") is not None
    ):
        _renderizar_distribuicao_manual(
            params_preparados,
            int(params_preparados.get("faixa_flex_pct", faixa_flex_pct)),
        )

    if st.session_state.ga_resultado is None:
        if st.session_state.get("ga_resultado_base_manual") is None:
            st.info(
                "Ajuste os parâmetros na barra lateral e clique em **Gerar amostra** "
                "ou **Definir metas por unidade**."
            )
        return

    resultado = st.session_state.ga_resultado
    params = st.session_state.ga_params

    st.caption(
        f"Fonte usada no resultado: **{params['fonte_dados_label']}** · "
        f"Distribuição: **{params.get('metodo_distribuicao_label', 'Proporcional à população')}**"
    )
    if params.get("recorte_descricao"):
        st.caption(f"Recorte aplicado: **{params['recorte_descricao']}**")

    # -----------------------------------------------------------------
    # Métricas
    # -----------------------------------------------------------------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unidades territoriais", len(resultado))
    c2.metric("Total de entrevistas", int(resultado["amostra"].sum()))
    coluna_pop_base = "populacao_base" if "populacao_base" in resultado.columns else "populacao_total"
    c3.metric("População base", f"{int(resultado[coluna_pop_base].sum()):,}".replace(",", "."))
    c4.metric("Menor unidade", int(resultado["amostra"].min()))

    if "renda_disponivel" in resultado.columns:
        sem_renda = int((pd.to_numeric(resultado["renda_disponivel"], errors="coerce").fillna(0) == 0).sum())
        if sem_renda:
            st.warning(
                f"{sem_renda} unidade(s) não possuem cobertura completa de renda. "
                "As respectivas cotas de renda foram deixadas em branco, não zeradas."
            )

    # -----------------------------------------------------------------
    # Fechamento por camadas
    # -----------------------------------------------------------------
    if st.session_state.ga_camadas_resumos:
        st.subheader("Fechamento por camadas")
        validacoes = st.session_state.ga_validacoes or []
        if validacoes and all(v["ok"] for v in validacoes):
            st.success("Todas as camadas fecham no total e cada nível filho fecha com seu nível pai.")
        elif validacoes:
            st.error("Foi encontrada divergência de fechamento entre as camadas.")

        for nivel, resumo_camada in st.session_state.ga_camadas_resumos.items():
            zeros = int((resumo_camada["amostra"] == 0).sum())
            titulo = (
                f"{ROTULOS_NIVEIS[nivel]} — {len(resumo_camada)} unidade(s), "
                f"{int(resumo_camada['amostra'].sum())} entrevistas"
            )
            with st.expander(titulo, expanded=nivel == params["nivel_quebra"]):
                if zeros:
                    st.caption(
                        f"{zeros} unidade(s) receberam amostra 0 na distribuição "
                        f"{params.get('metodo_distribuicao_label', '').lower()}."
                    )
                cols_resumo = [
                    c
                    for c in [
                        "uf",
                        "regiao_intermediaria",
                        "regiao_imediata",
                        "municipio",
                        "distrito",
                        "populacao_base",
                        "percentual",
                        "percentual_no_pai",
                        "amostra_ideal",
                        "amostra_proporcional",
                        "amostra",
                        "percentual_amostra",
                        "diferenca_vs_proporcional",
                    ]
                    if c in resumo_camada.columns
                ]
                tabela_camada = resumo_camada[cols_resumo].copy().rename(
                    columns={
                        "uf": "UF",
                        "regiao_intermediaria": "Região Intermediária",
                        "regiao_imediata": "Região Imediata",
                        "municipio": "Município",
                        "distrito": "Distrito",
                        "populacao_base": "População-base",
                        "percentual": "% real do total",
                        "percentual_no_pai": "% real no pai",
                        "amostra_ideal": "Amostra ideal",
                        "amostra_proporcional": "Sugestão proporcional",
                        "amostra": "Amostra aplicada",
                        "percentual_amostra": "% da amostra",
                        "diferenca_vs_proporcional": "Diferença",
                    }
                )
                for col_pct in ["% real do total", "% real no pai", "% da amostra"]:
                    if col_pct in tabela_camada.columns:
                        tabela_camada[col_pct] = tabela_camada[col_pct].round(2)
                st.dataframe(tabela_camada, use_container_width=True, hide_index=True)

        with st.expander("Validações técnicas"):
            for validacao in validacoes:
                icone = "✅" if validacao["ok"] else "❌"
                st.write(f"{icone} {validacao['mensagem']}")

    # -----------------------------------------------------------------
    # Preview
    # -----------------------------------------------------------------
    st.subheader("Preview da amostra final")
    colunas_preview = {
        "regiao_intermediaria": "Meso (Interm.)",
        "regiao_imediata": "Micro (Imed.)",
        "municipio": "Município",
        "distrito": "Distrito",
        "populacao_base": "População-base",
        "populacao_total": "População total",
        "percentual": "% real",
        "amostra_ideal": "Amostra ideal",
        "amostra_proporcional": "Sugestão proporcional",
        "amostra": "Amostra aplicada",
        "percentual_amostra": "% da amostra",
        "diferenca_vs_proporcional": "Diferença",
        "minimo": "Mín.",
        "maximo": "Máx.",
        "cota_masc": "M",
        "cota_fem": "F",
        "cota_16_19": "16-19",
        "cota_20_29": "20-29",
        "cota_30_39": "30-39",
        "cota_40_49": "40-49",
        "cota_50mais": "50+",
        "cota_renda_ate_2sm": "≤2sm",
        "cota_renda_2_a_5sm": "2-5sm",
        "cota_renda_5_a_10sm": "5-10sm",
        "cota_renda_mais_10sm": ">10sm",
        "renda_disponivel": "Renda disponível",
        "renda_cobertura_percentual": "Cobertura renda (%)",
    }
    cols_existentes = [c for c in colunas_preview if c in resultado.columns]
    preview = resultado[cols_existentes].rename(columns=colunas_preview)
    for col_pct in ["% real", "% da amostra"]:
        if col_pct in preview.columns:
            preview[col_pct] = preview[col_pct].round(2)
    if "Amostra ideal" in preview.columns:
        preview["Amostra ideal"] = preview["Amostra ideal"].round(3)
    if "Cobertura renda (%)" in preview.columns:
        preview["Cobertura renda (%)"] = preview["Cobertura renda (%)"].round(1)
    st.dataframe(preview, use_container_width=True, hide_index=True)

    # -----------------------------------------------------------------
    # Campanhas
    # -----------------------------------------------------------------
    st.subheader("Divisão em campanhas (opcional)")
    usar_campanhas = st.checkbox(
        "Dividir amostra em campanhas",
        value=st.session_state.ga_campanhas is not None,
        key="ga_usar_campanhas",
    )

    if usar_campanhas:
        with st.form("ga_form_campanhas"):
            col_a, col_b = st.columns([1, 3])
            with col_a:
                num_campanhas = st.number_input(
                    "Nº de campanhas",
                    min_value=1,
                    max_value=20,
                    value=5,
                    step=1,
                    key="ga_num_campanhas",
                )
            with col_b:
                nomes_campanhas = []
                cols_nomes = st.columns(min(int(num_campanhas), 5))
                for i in range(int(num_campanhas)):
                    with cols_nomes[i % len(cols_nomes)]:
                        nome = st.text_input(
                            f"Nome {i + 1}",
                            value=f"CAMPANHA {i + 1}",
                            key=f"ga_nome_camp_{i}",
                        )
                        nomes_campanhas.append(nome)

            st.markdown(
                "**Agrupamento**: escolha a qual campanha cada unidade pertence. "
                f"Nível: **{params['nivel_quebra_label']}**."
            )
            st.caption("Deixe vazio para não incluir a unidade em nenhuma campanha.")
            col_nome_unidade = params["nivel_quebra"]
            editor_df = resultado[[col_nome_unidade, "amostra"]].copy()
            editor_df.columns = ["Unidade", "Amostra"]
            editor_df["Campanha"] = ""

            edited = st.data_editor(
                editor_df,
                column_config={
                    "Unidade": st.column_config.TextColumn("Unidade", disabled=True),
                    "Amostra": st.column_config.NumberColumn("Amostra", disabled=True),
                    "Campanha": st.column_config.SelectboxColumn(
                        "Campanha",
                        options=[""] + nomes_campanhas,
                        required=False,
                    ),
                },
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                key="ga_editor_campanhas",
            )
            submit_camp = st.form_submit_button(
                "✅ Aplicar campanhas",
                type="primary",
                use_container_width=True,
            )

        if submit_camp:
            resultado_temp = resultado.copy()
            resultado_temp["_campanha"] = edited["Campanha"].values
            dict_camp = {}
            for nome in nomes_campanhas:
                df_camp = resultado_temp[
                    resultado_temp["_campanha"] == nome
                ].drop(columns="_campanha").copy()
                if not df_camp.empty:
                    dict_camp[nome] = df_camp
            st.session_state.ga_campanhas = dict_camp if dict_camp else None

    if st.session_state.ga_campanhas:
        resumo = [
            {
                "Campanha": nome,
                "Unidades": len(df_c),
                "Amostra": int(df_c["amostra"].sum()),
                "% do total": round(
                    df_c["amostra"].sum() / resultado["amostra"].sum() * 100,
                    2,
                ),
            }
            for nome, df_c in st.session_state.ga_campanhas.items()
        ]
        st.markdown("**Resumo das campanhas**")
        st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True)
        total_atribuido = sum(r["Amostra"] for r in resumo)
        total_geral = int(resultado["amostra"].sum())
        nao_atribuido = total_geral - total_atribuido
        if nao_atribuido > 0:
            st.warning(
                f"{nao_atribuido} entrevistas "
                f"({nao_atribuido / total_geral * 100:.1f}%) não foram atribuídas."
            )

    # -----------------------------------------------------------------
    # Excel
    # -----------------------------------------------------------------
    st.subheader("Download")
    st.caption(
        "O Excel inclui a aba **Informações**, com população real, memória de cálculo, "
        "comparação entre sugestão proporcional e valor aplicado, validações e resumo "
        "das campanhas. A exibição atual das cotas é: "
        f"**{exibicao_cotas_label}**."
    )
    buffer = io.BytesIO()
    caminho_temporario = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            caminho_temporario = tmp.name
        informacoes_exportacao = {
            **params,
            "exibicao_cotas": exibicao_cotas,
            "exibicao_cotas_label": exibicao_cotas_label,
        }
        exportar_amostra(
            df_completo=resultado,
            dict_campanhas=st.session_state.ga_campanhas,
            resumos_camadas=st.session_state.ga_camadas_resumos,
            caminho_saida=caminho_temporario,
            informacoes=informacoes_exportacao,
            exibicao_cotas=exibicao_cotas,
        )
        with open(caminho_temporario, "rb") as arquivo_excel:
            buffer.write(arquivo_excel.read())
    finally:
        if caminho_temporario and os.path.exists(caminho_temporario):
            os.unlink(caminho_temporario)

    nome_arquivo = f"Amostra_{params['uf'] or 'Brasil'}_{params['amostra_total']}.xlsx"
    st.download_button(
        label="⬇️ Baixar Excel",
        data=buffer.getvalue(),
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        key="ga_download_button",
    )
