"""
Gerador de Amostra.

Gera amostras estratificadas por hierarquia territorial do IBGE (Região
Intermediária → Região Imediata → Município → Distrito), com distribuição
proporcional à população (Método do Maior Resto) e cotas demográficas de
sexo, idade e renda dentro de cada unidade.

Diferente das análises de correspondência, não depende do `dados`
carregado no início do app: usa as bases mestre pré-processadas do IBGE
em dados/gerador_amostra/ (não tem upload próprio nem depende de arquivo
carregado por fora).
"""
import io
import tempfile

import streamlit as st
import pandas as pd

from core.amostra_math import calcular_amostra, _carregar_municipios, _carregar_distritos
from core.amostra_exportador import exportar_amostra


def modulo_gerador_amostra():
    st.header("Gerador de Amostra")
    st.caption(
        "Gera amostras estratificadas por hierarquia territorial (IBGE), "
        "com distribuição proporcional à população e cotas demográficas "
        "de sexo, idade e renda dentro de cada unidade."
    )

    # -----------------------------------------------------------------
    # Estado da sessão (persiste entre re-execuções), namespaced com "ga_"
    # -----------------------------------------------------------------
    if "ga_resultado" not in st.session_state:
        st.session_state.ga_resultado = None
    if "ga_params" not in st.session_state:
        st.session_state.ga_params = None
    if "ga_campanhas" not in st.session_state:
        st.session_state.ga_campanhas = None

    # -----------------------------------------------------------------
    # Sidebar — Formulário de parâmetros (só executa ao clicar "Gerar amostra")
    # -----------------------------------------------------------------
    with st.sidebar.form("ga_form_parametros", clear_on_submit=False):
        st.header("Parâmetros da pesquisa")

        incluir_distritos = st.toggle(
            "Trabalhar em nível de distrito",
            value=False,
            help="Ativa para desagregar municípios em seus distritos.",
            key="ga_incluir_distritos",
        )

        # Carrega base para preencher UFs
        df_base = _carregar_distritos() if incluir_distritos else _carregar_municipios()
        ufs_disponiveis = sorted(df_base["uf"].unique().tolist())

        uf_opcao = st.selectbox(
            "Estado (UF)",
            options=["(Brasil inteiro)"] + ufs_disponiveis,
            index=(ufs_disponiveis.index("SP") + 1) if "SP" in ufs_disponiveis else 0,
            key="ga_uf_opcao",
        )

        # Regiões Intermediárias (mostradas com base no estado)
        df_para_ri = df_base if uf_opcao == "(Brasil inteiro)" else df_base[df_base["uf"] == uf_opcao]
        regioes_interm = sorted(df_para_ri["regiao_intermediaria"].dropna().unique().tolist())
        filtro_ri = st.multiselect(
            "Regiões Intermediárias (opcional)",
            options=regioes_interm,
            help="Deixe vazio para incluir todas.",
            key="ga_filtro_ri",
        )

        niveis_quebra_map = {
            "Região Intermediária": "regiao_intermediaria",
            "Região Imediata": "regiao_imediata",
            "Município": "municipio",
        }
        if incluir_distritos:
            niveis_quebra_map["Distrito"] = "distrito"
        nivel_quebra_label = st.selectbox(
            "Nível de quebra da amostra",
            options=list(niveis_quebra_map.keys()),
            index=0,
            key="ga_nivel_quebra_label",
        )

        amostra_total = st.number_input(
            "Total de entrevistas", min_value=1, value=1000, step=100,
            key="ga_amostra_total",
        )

        base_populacional_label = st.radio(
            "Base populacional",
            options=["16 anos ou mais (idade da pesquisa)", "Total absoluto (todas as idades)"],
            index=0,
            key="ga_base_populacional_label",
        )

        faixa_flex_pct = st.slider(
            "Faixa mínimo/máximo (%)",
            min_value=0, max_value=50, value=20, step=5,
            help="Ex: 20% em 17 → min 14, max 20.",
            key="ga_faixa_flex_pct",
        )

        submitted = st.form_submit_button(
            "🎯 Gerar amostra", type="primary", use_container_width=True
        )

    # Só recalcula se o botão foi clicado
    if submitted:
        uf = None if uf_opcao == "(Brasil inteiro)" else uf_opcao
        codigos_ri = None
        if filtro_ri:
            codigos_ri = df_para_ri[
                df_para_ri["regiao_intermediaria"].isin(filtro_ri)
            ]["cod_regiao_intermediaria"].unique().tolist()
        base_populacional = "16_mais" if base_populacional_label.startswith("16") else "total"

        try:
            st.session_state.ga_resultado = calcular_amostra(
                uf=uf,
                codigos_regioes_intermediarias=codigos_ri,
                amostra_total=int(amostra_total),
                nivel_quebra=niveis_quebra_map[nivel_quebra_label],
                base_populacional=base_populacional,
                incluir_distritos=incluir_distritos,
                faixa_flex=faixa_flex_pct / 100,
            )
            st.session_state.ga_params = dict(
                uf=uf, amostra_total=int(amostra_total),
                nivel_quebra=niveis_quebra_map[nivel_quebra_label],
                nivel_quebra_label=nivel_quebra_label,
            )
            st.session_state.ga_campanhas = None  # limpa campanhas antigas
        except ValueError as e:
            st.error(f"Erro no cálculo: {e}")

    # -----------------------------------------------------------------
    # Se ainda não gerou nada, mostra a mensagem inicial e encerra aqui
    # -----------------------------------------------------------------
    if st.session_state.ga_resultado is None:
        st.info("Ajuste os parâmetros na barra lateral e clique em **Gerar amostra**.")
        return

    resultado = st.session_state.ga_resultado
    params = st.session_state.ga_params

    # -----------------------------------------------------------------
    # Métricas resumo
    # -----------------------------------------------------------------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unidades territoriais", len(resultado))
    c2.metric("Total de entrevistas", int(resultado["amostra"].sum()))
    c3.metric("População base", f"{int(resultado['populacao_total'].sum()):,}".replace(",", "."))
    c4.metric("Menor unidade", int(resultado["amostra"].min()))

    # -----------------------------------------------------------------
    # Preview da tabela
    # -----------------------------------------------------------------
    st.subheader("Preview da amostra")

    colunas_preview = {
        "regiao_intermediaria": "Meso (Interm.)",
        "regiao_imediata": "Micro (Imed.)",
        "municipio": "Município",
        "distrito": "Distrito",
        "populacao_total": "População",
        "percentual": "%",
        "amostra": "Amostra",
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
    }
    cols_existentes = [c for c in colunas_preview.keys() if c in resultado.columns]
    preview = resultado[cols_existentes].rename(columns=colunas_preview)
    preview["%"] = preview["%"].round(2)
    st.dataframe(preview, use_container_width=True, hide_index=True)

    # -----------------------------------------------------------------
    # Divisão em campanhas (formulário separado)
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
                    "Nº de campanhas", min_value=1, max_value=20, value=5, step=1,
                    key="ga_num_campanhas",
                )

            with col_b:
                nomes_campanhas = []
                cols_nomes = st.columns(min(int(num_campanhas), 5))
                for i in range(int(num_campanhas)):
                    with cols_nomes[i % len(cols_nomes)]:
                        nome = st.text_input(
                            f"Nome {i+1}", value=f"CAMPANHA {i+1}", key=f"ga_nome_camp_{i}"
                        )
                        nomes_campanhas.append(nome)

            st.markdown(
                f"**Agrupamento**: escolha a qual campanha cada unidade pertence. "
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
                        "Campanha", options=[""] + nomes_campanhas, required=False,
                    ),
                },
                hide_index=True, use_container_width=True, num_rows="fixed",
                key="ga_editor_campanhas",
            )

            submit_camp = st.form_submit_button(
                "✅ Aplicar campanhas", type="primary", use_container_width=True
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

    # Mostra resumo das campanhas se já aplicadas
    if st.session_state.ga_campanhas:
        resumo = [
            {
                "Campanha": nome,
                "Unidades": len(df_c),
                "Amostra": int(df_c["amostra"].sum()),
                "% do total": round(
                    df_c["amostra"].sum() / resultado["amostra"].sum() * 100, 2
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
                f"⚠️ {nao_atribuido} entrevistas "
                f"({nao_atribuido/total_geral*100:.1f}%) não atribuídas a nenhuma campanha."
            )

    # -----------------------------------------------------------------
    # Download do Excel
    # -----------------------------------------------------------------
    st.subheader("Download")

    buffer = io.BytesIO()
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        exportar_amostra(
            df_completo=resultado,
            dict_campanhas=st.session_state.ga_campanhas,
            caminho_saida=tmp.name,
        )
        with open(tmp.name, "rb") as f:
            buffer.write(f.read())

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
