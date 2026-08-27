"""Painel Streamlit para consulta do banco territorial e demográfico."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core.banco_amostras import (
    CAMINHO_BANCO_PADRAO,
    ROTULOS_NIVEIS_BANCO,
    BancoAmostrasErro,
    buscar_territorios,
    carregar_base_amostragem,
    consultar_cobertura,
    listar_alertas,
    listar_fontes,
    listar_ufs,
    obter_metadados_banco,
    resumo_alertas,
    resumo_banco,
    verificar_integridade,
)

NIVEIS_EXPLORADOR = {
    "Município": "municipio",
    "Distrito": "distrito",
    "Subdistrito": "subdistrito",
    "Bairro": "bairro",
}


def _formatar_inteiro(valor) -> str:
    try:
        return f"{int(valor):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "—"


def _indicador_ativo(valor) -> bool:
    try:
        return int(valor) == 1
    except (TypeError, ValueError):
        return False


def _mostrar_detalhes_territorio(chave: str, nivel: str) -> None:
    detalhe = carregar_base_amostragem(
        nivel,
        chaves_territoriais=[chave],
        somente_com_demografia=False,
    )
    if detalhe.empty:
        st.warning("O território selecionado não foi encontrado na visão demográfica.")
        return

    linha = detalhe.iloc[0]
    st.markdown(f"### {linha['nome']}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Nível", ROTULOS_NIVEIS_BANCO.get(nivel, nivel))
    c2.metric("Código", str(linha["codigo"]))
    c3.metric("UF", linha.get("uf") or "—")
    c4.metric("População total", _formatar_inteiro(linha.get("populacao_total")))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("População 16+", _formatar_inteiro(linha.get("populacao_16_mais_referencia")))
    c6.metric("Masculino", _formatar_inteiro(linha.get("masculino_total")))
    c7.metric("Feminino", _formatar_inteiro(linha.get("feminino_total")))
    c8.metric(
        "Renda disponível",
        "Sim" if _indicador_ativo(linha.get("renda_disponivel")) else "Não",
    )

    with st.expander("Hierarquia e procedência", expanded=False):
        campos = {
            "Chave territorial": linha.get("chave_territorial"),
            "Grande região": linha.get("grande_regiao_nome"),
            "Estado": linha.get("nome_uf"),
            "Região intermediária": linha.get("nome_regiao_intermediaria"),
            "Região imediata": linha.get("nome_regiao_imediata"),
            "Município": linha.get("nome_municipio"),
            "Distrito": linha.get("nome_distrito"),
            "Método da população 16+": linha.get("metodo_populacao_16_mais_referencia"),
            "Ano da renda": linha.get("ano_referencia_renda"),
            "Observações": linha.get("observacoes"),
        }
        tabela = pd.DataFrame(
            [{"Campo": campo, "Valor": valor if pd.notna(valor) else "—"} for campo, valor in campos.items()]
        )
        st.dataframe(tabela, use_container_width=True, hide_index=True)

    with st.expander("Demografia", expanded=True):
        demografia = pd.DataFrame(
            [
                {"Indicador": "População total", "Valor": linha.get("populacao_total")},
                {"Indicador": "População 15+", "Valor": linha.get("populacao_15_mais")},
                {"Indicador": "População 16+ de referência", "Valor": linha.get("populacao_16_mais_referencia")},
                {"Indicador": "16 a 19", "Valor": linha.get("idade_16_19_referencia")},
                {"Indicador": "20 a 29", "Valor": linha.get("idade_20_29")},
                {"Indicador": "30 a 39", "Valor": linha.get("idade_30_39:1")},
                {"Indicador": "40 a 49", "Valor": linha.get("idade_40_49:1")},
                {"Indicador": "50 ou mais", "Valor": linha.get("idade_50_mais")},
            ]
        )
        demografia["Valor"] = demografia["Valor"].apply(_formatar_inteiro)
        st.dataframe(demografia, use_container_width=True, hide_index=True)

    with st.expander("Renda", expanded=False):
        renda = pd.DataFrame(
            [
                {"Faixa": "Até 2 salários mínimos", "População": linha.get("renda_ate_2sm")},
                {"Faixa": "De 2 a 5 salários mínimos", "População": linha.get("renda_2_a_5sm")},
                {"Faixa": "De 5 a 10 salários mínimos", "População": linha.get("renda_5_a_10sm")},
                {"Faixa": "Mais de 10 salários mínimos", "População": linha.get("renda_mais_10sm")},
            ]
        )
        renda["População"] = renda["População"].apply(_formatar_inteiro)
        st.dataframe(renda, use_container_width=True, hide_index=True)

    filhos = {
        "municipio": ["distrito", "bairro"],
        "distrito": ["subdistrito"],
    }.get(nivel, [])
    if filhos:
        st.markdown("#### Cobertura dos níveis inferiores")
        coberturas = []
        for nivel_filho in filhos:
            cobertura = consultar_cobertura(
                nivel_pai=nivel,
                codigo_pai=str(linha["codigo"]),
                nivel_filho=nivel_filho,
            )
            if not cobertura.empty:
                coberturas.append(cobertura)
        if coberturas:
            tabela_cobertura = pd.concat(coberturas, ignore_index=True)
            exibir = tabela_cobertura[
                [
                    "nivel_filho",
                    "quantidade_registros_filhos",
                    "quantidade_filhos_com_dados",
                    "populacao_pai",
                    "populacao_filhos",
                    "proporcao_cobertura",
                    "fecha_exatamente",
                    "observacoes",
                ]
            ].copy()
            exibir["nivel_filho"] = exibir["nivel_filho"].map(
                lambda valor: ROTULOS_NIVEIS_BANCO.get(valor, valor)
            )
            exibir["proporcao_cobertura"] = (
                pd.to_numeric(exibir["proporcao_cobertura"], errors="coerce") * 100
            ).round(2)
            exibir["fecha_exatamente"] = exibir["fecha_exatamente"].map(
                {1: "Sim", 0: "Não"}
            )
            exibir = exibir.rename(
                columns={
                    "nivel_filho": "Nível filho",
                    "quantidade_registros_filhos": "Registros",
                    "quantidade_filhos_com_dados": "Com demografia",
                    "populacao_pai": "População pai",
                    "populacao_filhos": "População dos filhos",
                    "proporcao_cobertura": "Cobertura (%)",
                    "fecha_exatamente": "Fecha exatamente",
                    "observacoes": "Observações",
                }
            )
            st.dataframe(exibir, use_container_width=True, hide_index=True)
        else:
            st.info("Não há medição de cobertura para os níveis inferiores desta unidade.")


def modulo_banco_amostras() -> None:
    st.header("Banco de Amostras")
    st.caption(
        "Consulte a base territorial e demográfica usada pelo Gerador. "
        "Nesta versão, o banco é somente leitura dentro do aplicativo."
    )

    try:
        integridade = verificar_integridade()
        resumo = resumo_banco()
        alertas_resumo = resumo_alertas()
    except BancoAmostrasErro as exc:
        st.error(str(exc))
        return

    if integridade["integridade"] != "ok" or integridade["chaves_estrangeiras_invalidas"]:
        st.error("O banco falhou na verificação de integridade. Não use os dados para calcular amostras.")
    else:
        st.success(
            f"Banco íntegro · esquema {integridade['versao_schema']} · "
            "chaves estrangeiras válidas"
        )

    aba_visao, aba_explorar, aba_qualidade, aba_fontes = st.tabs(
        ["Visão geral", "Explorar territórios", "Qualidade", "Fontes e backup"]
    )

    with aba_visao:
        total_territorios = int(resumo["quantidade"].sum())
        com_demografia = int(resumo["com_demografia"].fillna(0).sum())
        total_alertas = int(alertas_resumo["quantidade"].sum()) if not alertas_resumo.empty else 0
        tamanho_mb = CAMINHO_BANCO_PADRAO.stat().st_size / (1024 * 1024)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Territórios", _formatar_inteiro(total_territorios))
        c2.metric("Com demografia", _formatar_inteiro(com_demografia))
        c3.metric("Alertas registrados", _formatar_inteiro(total_alertas))
        c4.metric("Tamanho do banco", f"{tamanho_mb:.1f} MB")

        tabela_resumo = resumo.copy()
        tabela_resumo["nivel"] = tabela_resumo["nivel"].map(
            lambda valor: ROTULOS_NIVEIS_BANCO.get(valor, valor)
        )
        tabela_resumo = tabela_resumo.rename(
            columns={
                "nivel": "Nível",
                "quantidade": "Quantidade",
                "com_demografia": "Com demografia",
                "registros_com_codigo_duplicado": "Códigos duplicados preservados",
            }
        )
        st.dataframe(tabela_resumo, use_container_width=True, hide_index=True)
        st.info(
            "Município e distrito já podem alimentar o Gerador. Bairro e subdistrito "
            "ficam disponíveis para consulta enquanto definimos as regras de cobertura parcial."
        )

    with aba_explorar:
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            nivel_label = st.selectbox(
                "Nível territorial",
                options=list(NIVEIS_EXPLORADOR.keys()),
                key="ba_nivel_explorador",
            )
            nivel = NIVEIS_EXPLORADOR[nivel_label]
        with c2:
            ufs = listar_ufs()
            opcoes_uf = ["Todas"] + ufs["uf"].dropna().astype(str).tolist()
            uf = st.selectbox("UF", options=opcoes_uf, key="ba_uf_explorador")
        with c3:
            texto = st.text_input(
                "Buscar pelo nome",
                placeholder="Ex.: Copacabana, Campinas, Sé...",
                key="ba_texto_explorador",
            )

        col_a, col_b, col_c = st.columns([1, 1, 2])
        with col_a:
            somente_com_dados = st.checkbox(
                "Somente com demografia",
                value=True,
                key="ba_somente_com_demografia",
            )
        with col_b:
            limite = st.number_input(
                "Limite",
                min_value=10,
                max_value=1000,
                value=100,
                step=10,
                key="ba_limite_busca",
            )
        with col_c:
            buscar = st.button(
                "Pesquisar territórios",
                type="primary",
                use_container_width=True,
                key="ba_buscar_territorios",
            )

        if buscar:
            st.session_state["ba_resultados_busca"] = buscar_territorios(
                nivel,
                texto=texto or None,
                uf=None if uf == "Todas" else uf,
                somente_com_demografia=somente_com_dados,
                limite=int(limite),
            )
            st.session_state["ba_nivel_resultado"] = nivel

        resultados = st.session_state.get("ba_resultados_busca")
        nivel_resultado = st.session_state.get("ba_nivel_resultado")
        if isinstance(resultados, pd.DataFrame) and not resultados.empty:
            st.markdown(f"**{len(resultados)} território(s) encontrado(s)**")
            tabela = resultados[
                [
                    "chave_territorial",
                    "codigo",
                    "nome",
                    "uf",
                    "nome_municipio",
                    "nome_distrito",
                    "codigo_duplicado",
                    "dados_demograficos_disponiveis",
                ]
            ].copy()
            tabela = tabela.rename(
                columns={
                    "chave_territorial": "Chave",
                    "codigo": "Código",
                    "nome": "Nome",
                    "uf": "UF",
                    "nome_municipio": "Município",
                    "nome_distrito": "Distrito",
                    "codigo_duplicado": "Código duplicado",
                    "dados_demograficos_disponiveis": "Com demografia",
                }
            )
            st.dataframe(tabela, use_container_width=True, hide_index=True)

            opcoes = {
                f"{linha['nome']} · {linha.get('uf') or '—'} · {linha['codigo']} · {linha['chave_territorial']}": linha[
                    "chave_territorial"
                ]
                for _, linha in resultados.iterrows()
            }
            escolha = st.selectbox(
                "Abrir detalhes",
                options=list(opcoes.keys()),
                key="ba_territorio_selecionado",
            )
            _mostrar_detalhes_territorio(opcoes[escolha], nivel_resultado)
        elif buscar:
            st.info("Nenhum território encontrado com os filtros informados.")
        else:
            st.info("Defina os filtros e clique em **Pesquisar territórios**.")

    with aba_qualidade:
        st.markdown("### Resumo dos alertas")
        if alertas_resumo.empty:
            st.success("Nenhum alerta registrado.")
        else:
            resumo_exibir = alertas_resumo.rename(
                columns={
                    "severidade": "Severidade",
                    "categoria": "Categoria",
                    "quantidade": "Quantidade",
                }
            )
            st.dataframe(resumo_exibir, use_container_width=True, hide_index=True)

            severidades = ["Todas"] + alertas_resumo["severidade"].dropna().unique().tolist()
            categorias = ["Todas"] + alertas_resumo["categoria"].dropna().unique().tolist()
            c1, c2 = st.columns(2)
            with c1:
                severidade = st.selectbox(
                    "Severidade",
                    severidades,
                    key="ba_filtro_severidade",
                )
            with c2:
                categoria = st.selectbox(
                    "Categoria",
                    categorias,
                    key="ba_filtro_categoria",
                )
            alertas = listar_alertas(
                severidade=None if severidade == "Todas" else severidade,
                categoria=None if categoria == "Todas" else categoria,
                limite=5000,
            )
            st.dataframe(alertas, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Baixar alertas filtrados em CSV",
                data=alertas.to_csv(index=False).encode("utf-8-sig"),
                file_name="alertas_qualidade_banco_amostras.csv",
                mime="text/csv",
                use_container_width=True,
                key="ba_download_alertas",
            )

    with aba_fontes:
        st.markdown("### Fontes incorporadas")
        fontes = listar_fontes()
        st.dataframe(fontes, use_container_width=True, hide_index=True)

        st.markdown("### Metadados do banco")
        metadados = obter_metadados_banco()
        st.dataframe(metadados, use_container_width=True, hide_index=True)

        st.markdown("### Backup")
        st.warning(
            "O download abaixo contém o banco completo. Guarde-o como cópia de segurança; "
            "não edite o arquivo manualmente durante o uso do aplicativo."
        )
        st.download_button(
            "⬇️ Baixar banco SQLite",
            data=CAMINHO_BANCO_PADRAO.read_bytes(),
            file_name=CAMINHO_BANCO_PADRAO.name,
            mime="application/vnd.sqlite3",
            use_container_width=True,
            key="ba_download_sqlite",
        )
