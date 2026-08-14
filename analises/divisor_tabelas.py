"""Streamlit screen for automatic table splitting in Excel reports."""
from __future__ import annotations

import io

import openpyxl
import pandas as pd
import streamlit as st

from core.divisor_tabelas_math import processar_workbook


def _result_rows(summary):
    rows = []
    for item in summary["results"]:
        rows.append({
            "Aba": item.sheet,
            "Tabela": item.title,
            "Linhas originais": f"{item.original_start}-{item.original_end}",
            "Labels": item.labels,
            "Partes": item.parts,
            "Distribuicao": " + ".join(str(x) for x in item.part_sizes),
        })
    return rows


def modulo_divisor_tabelas():
    st.header("Divisor de Tabelas de Regioes")
    st.write(
        "Envie um relatorio Excel finalizado. O modulo identifica tabelas de regioes "
        "pela presenca de um bloco de legenda (LEGENDA, LEGENDA REGIOES, LEGENDA CAPITAL etc.) e divide somente aquelas em que uma quebra "
        "de pagina atravesa qualquer ponto do bloco da tabela: cabecalho, labels, Base, "
        "Pergunta, espacamento ou a propria LEGENDA. A legenda e repetida integralmente "
        "em todas as partes."
    )
    st.info(
        "A analise considera as quebras manuais gravadas no arquivo e tambem "
        "estima as quebras automaticas pelas alturas das linhas, tamanho do "
        "papel, margens e escala de impressao."
    )

    uploaded = st.file_uploader(
        "Carregar relatorio Excel",
        type=["xlsx", "xlsm"],
        key="divisor_tabelas_upload",
    )
    if uploaded is None:
        return

    if st.button("Analisar e dividir tabelas", key="divisor_tabelas_processar"):
        try:
            raw = uploaded.getvalue()
            keep_vba = uploaded.name.lower().endswith(".xlsm")
            wb = openpyxl.load_workbook(io.BytesIO(raw), keep_vba=keep_vba, rich_text=True)

            with st.status("Analisando paginacao e tabelas...", expanded=True) as status:
                summary = processar_workbook(wb)
                status.update(label="Analise concluida", state="complete")

            out = io.BytesIO()
            wb.save(out)
            out_bytes = out.getvalue()

            st.session_state["divisor_tabelas_resultado"] = out_bytes
            st.session_state["divisor_tabelas_summary"] = summary
            st.session_state["divisor_tabelas_nome"] = uploaded.name
        except Exception as exc:
            st.exception(exc)
            return

    summary = st.session_state.get("divisor_tabelas_summary")
    out_bytes = st.session_state.get("divisor_tabelas_resultado")
    if not summary or out_bytes is None:
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Abas analisadas", summary["sheets_analyzed"])
    c2.metric("Tabelas com legenda", summary.get("tables_with_legend", 0))
    c3.metric("Tabelas divididas", summary["tables_split"])
    c4.metric("Partes geradas", summary["parts_created"])

    if summary["tables_split"]:
        st.success(
            f"{summary['tables_split']} tabela(s) precisaram de divisao e foram ajustadas."
        )
        rows = _result_rows(summary)
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.success(
            "Nenhuma tabela com legenda atravessa uma quebra de pagina segundo a configuracao atual. "
            "As demais tabelas nao sao alteradas por este modulo."
        )

    original_name = st.session_state.get("divisor_tabelas_nome", "relatorio.xlsx")
    if "." in original_name:
        stem, ext = original_name.rsplit(".", 1)
        output_name = f"{stem}_tabelas_divididas.{ext}"
    else:
        output_name = original_name + "_tabelas_divididas.xlsx"

    st.download_button(
        "Baixar relatorio processado",
        data=out_bytes,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="divisor_tabelas_download",
    )
