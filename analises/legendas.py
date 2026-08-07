"""
Legendas.

Localiza, num relatório de tabelas, os blocos segmentados por território
(Regiões, Bairros, Municípios etc.) e insere a legenda correspondente
(vinda de um segundo arquivo) logo após a linha "Pergunta:" de cada
bloco — mantendo a primeira linha em branco como está, adicionando
"LEGENDA" + os itens, e uma linha de espaçamento antes do restante do
relatório continuar normalmente.

Diferente das análises de correspondência, não depende do `dados`
carregado no início do app: tem upload próprio (dois arquivos Excel).
"""
import io

import streamlit as st

from core.legendas_math import (
    parsear_blocos_tabela,
    parsear_blocos_legenda,
    bloco_precisa_legenda,
    parear_blocos,
    gerar_workbook_com_legendas,
)


def modulo_legendas():
    st.header("Legendas")
    st.caption(
        "Insere automaticamente a legenda de blocos segmentados por "
        "território (Regiões, Bairros, Municípios, Sub-Regiões etc.) "
        "logo após a linha 'Pergunta:' de cada tabela correspondente."
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
        with st.spinner("Lendo e casando os blocos..."):
            try:
                blocos_tabela = parsear_blocos_tabela(arquivo_tabela)
                arquivo_legenda.seek(0)
                blocos_legenda = parsear_blocos_legenda(arquivo_legenda)
            except Exception as e:
                st.error(f"Não foi possível ler os arquivos: {e}")
                return

            blocos_elegiveis = [b for b in blocos_tabela if bloco_precisa_legenda(b)]

            if not blocos_elegiveis:
                st.warning(
                    "Não encontrei nenhum bloco segmentado por território "
                    "(Regiões, Bairros, Municípios etc.) no relatório de "
                    "tabelas. Confira se os nomes de categoria batem com "
                    "o esperado."
                )
                return
            if not blocos_legenda:
                st.error(
                    "Não encontrei nenhum bloco 'LEGENDA' no arquivo de "
                    "legendas. Confirme se esse é o arquivo certo."
                )
                return

            pares, avisos = parear_blocos(blocos_elegiveis, blocos_legenda)

            arquivo_tabela.seek(0)
            wb_novo = gerar_workbook_com_legendas(arquivo_tabela, pares)

            saida = io.BytesIO()
            wb_novo.save(saida)
            saida.seek(0)

        st.session_state["lg_resultado_bytes"] = saida.getvalue()
        st.session_state["lg_resultado_pares"] = [
            {"titulo": bt["titulo"], "n_itens": len(bl["itens"])}
            for bt, bl in pares
        ]
        st.session_state["lg_resultado_avisos"] = avisos
        st.success(
            f"Pronto! {len(pares)} bloco(s) de legenda inserido(s) em "
            f"{len(blocos_elegiveis)} tabela(s) territorial(is) encontrada(s)."
        )

    if "lg_resultado_bytes" in st.session_state:
        st.markdown("---")
        st.subheader("Resultado")

        avisos = st.session_state.get("lg_resultado_avisos", [])
        if avisos:
            for aviso in avisos:
                st.warning(aviso)
        else:
            st.success("Todos os blocos foram pareados sem inconsistências.")

        for item in st.session_state["lg_resultado_pares"]:
            st.write(f"**{item['titulo']}** — {item['n_itens']} item(ns) de legenda inserido(s)")

        st.download_button(
            "Baixar relatório com legendas",
            data=st.session_state["lg_resultado_bytes"],
            file_name="Relatorio_com_Legendas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="lg_download"
        )
