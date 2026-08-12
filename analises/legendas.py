"""
Legendas.

Localiza, num relatório de tabelas, os blocos segmentados por território
(Regiões, Capital etc.) e insere a legenda correspondente (vinda de um
segundo arquivo) logo após a linha "Pergunta:" de cada bloco.

Casamento por CHAVE individual (código/rótulo de cada item, ex.: "Região
1", "Centro e Oeste"), não por tipo de tabela inteiro — uma tabela que só
usa parte das categorias (ex.: só as regiões 1, 3 e 7) recebe só os
itens de legenda correspondentes a essas, não a lista completa.

Diferente das análises de correspondência, não depende do `dados`
carregado no início do app: tem upload próprio (dois arquivos Excel).
"""
import io

import openpyxl
import streamlit as st

from core.legendas_math import parsear_legenda_por_chave, aplicar_legendas_por_chave


def modulo_legendas():
    st.header("Legendas")
    st.caption(
        "Insere automaticamente a legenda de blocos segmentados por "
        "território (Regiões, Capital etc.) logo após a linha 'Pergunta:' "
        "de cada tabela correspondente — só com os itens cujo código/rótulo "
        "realmente aparece no cabeçalho daquela tabela específica."
    )

    col1, col2 = st.columns(2)
    with col1:
        arquivo_tabela = st.file_uploader(
            "Relatório de tabelas (sem legenda)", type=["xlsx"], key="lg_upload_tabela"
        )
    with col2:
        arquivo_legenda = st.file_uploader(
            "Arquivo de legendas", type=["xlsx"], key="lg_upload_legenda"
        )

    if not arquivo_tabela or not arquivo_legenda:
        st.info("Envie os dois arquivos .xlsx para começar.")
        return

    if st.button("Calcular e gerar arquivo", key="lg_calcular"):
        with st.spinner("Lendo a referência e casando os cabeçalhos..."):
            try:
                dados_referencia = parsear_legenda_por_chave(arquivo_legenda)
            except Exception as e:
                st.error(f"Não foi possível ler o arquivo de legendas: {e}")
                return

            if not dados_referencia["mapa"]:
                st.error(
                    "Não encontrei nenhum item de legenda reconhecível no "
                    "arquivo de referência (cada item precisa de um código "
                    "numérico ou um rótulo antes de ':' logo abaixo de um "
                    "título como 'LEGENDA REGIÕES')."
                )
                return

            try:
                wb = openpyxl.load_workbook(arquivo_tabela, rich_text=True)
                ws = wb.active
                n_inseridas = aplicar_legendas_por_chave(ws, dados_referencia)
            except Exception as e:
                st.error(f"Não foi possível processar o relatório: {e}")
                return

            saida = io.BytesIO()
            wb.save(saida)

        st.session_state["lg_resultado_bytes"] = saida.getvalue()
        st.session_state["lg_resultado_n"] = n_inseridas
        st.session_state["lg_resultado_chaves"] = sorted(dados_referencia["mapa"].keys())
        if n_inseridas:
            st.success(f"Pronto! Legenda inserida em {n_inseridas} tabela(s).")
        else:
            st.warning(
                "Nenhuma tabela recebeu legenda — confira se os códigos/"
                "rótulos do cabeçalho do relatório batem com os da "
                "referência (chaves lidas na referência: "
                + ", ".join(sorted(dados_referencia["mapa"].keys())[:20])
                + ")."
            )

    if "lg_resultado_bytes" in st.session_state:
        st.markdown("---")
        st.subheader("Resultado")
        st.write(f"**{st.session_state['lg_resultado_n']}** tabela(s) receberam legenda.")

        st.download_button(
            "Baixar relatório com legendas",
            data=st.session_state["lg_resultado_bytes"],
            file_name="Relatorio_com_Legendas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="lg_download"
        )
