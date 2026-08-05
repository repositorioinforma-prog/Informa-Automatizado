"""
Processador de Base Reduzida.

Diferente das análises de correspondência, não depende do `dados`
(pandas) carregado no início do app: trabalha direto com o Workbook do
openpyxl (para preservar formatação original), então tem upload de
arquivo próprio.
"""
import io
import zipfile
from copy import copy

import openpyxl
import streamlit as st


def _copiar_celula_base_reduzida(origem, destino):
    """Copia valor e estilo (fonte, borda, preenchimento etc.) de uma célula para outra."""
    destino.value = origem.value
    if origem.has_style:
        destino.font = copy(origem.font)
        destino.border = copy(origem.border)
        destino.fill = copy(origem.fill)
        destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection)
        destino.alignment = copy(origem.alignment)


def _processar_tabelas_base_reduzida(wb, aba_original):
    """
    Percorre a aba selecionada procurando blocos de tabelas separados por
    linhas em branco. Em cada tabela, localiza a linha "Base reduzida" e
    remove as colunas cuja base seja menor que 20. Retorna:
    - um workbook consolidado com todas as tabelas já filtradas;
    - uma lista de (nome_arquivo, bytes) com cada tabela filtrada isolada.
    """
    ws = wb[aba_original]

    wb_consolidado = openpyxl.Workbook()
    ws_consolidado = wb_consolidado.active
    ws_consolidado.title = "Consolidado"

    linha_destino = 1
    max_col = ws.max_column
    max_row = ws.max_row

    linhas = list(ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col))
    i = 0
    arquivos_tabelas = []

    while i < len(linhas):
        inicio_bloco_vazio = all(cell.value is None for cell in linhas[i])
        proxima_preenchida = i + 1 < len(linhas) and any(
            cell.value is not None for cell in linhas[i + 1]
        )

        if inicio_bloco_vazio and proxima_preenchida:
            inicio_tabela = i + 1
            fim_tabela = inicio_tabela

            while fim_tabela < len(linhas) and any(
                cell.value is not None for cell in linhas[fim_tabela]
            ):
                fim_tabela += 1

            tabela = linhas[inicio_tabela:fim_tabela]

            idx_base = None
            for idx, linha in enumerate(tabela):
                if any(
                    'base reduzida' in str(cell.value).lower()
                    for cell in linha if cell.value
                ):
                    idx_base = idx
                    break

            if idx_base is not None:
                colunas_para_excluir = []
                base_linha = tabela[idx_base]
                for j in range(1, len(base_linha)):
                    valor = base_linha[j].value
                    try:
                        valor_str = str(valor).replace(',', '.')
                        limpo = ''.join(
                            ch for ch in valor_str if ch.isdigit() or ch in ['.', '-']
                        )
                        if limpo:
                            num = float(limpo)
                            if num < 20:
                                colunas_para_excluir.append(j)
                    except Exception:
                        continue

                for k in range(len(tabela)):
                    nova_linha = []
                    for j in range(len(tabela[k])):
                        if j not in colunas_para_excluir:
                            nova_linha.append(tabela[k][j])
                    tabela[k] = nova_linha

                # copia para o consolidado
                for linha in tabela:
                    for col_idx, cell in enumerate(linha):
                        destino = ws_consolidado.cell(row=linha_destino, column=col_idx + 1)
                        _copiar_celula_base_reduzida(cell, destino)
                    linha_destino += 1
                linha_destino += 1  # linha em branco entre tabelas

                # workbook separado só com esta tabela
                wb_tabela = openpyxl.Workbook()
                ws_tabela = wb_tabela.active
                for row_idx, linha in enumerate(tabela, start=1):
                    for col_idx, cell in enumerate(linha, start=1):
                        destino = ws_tabela.cell(row=row_idx, column=col_idx)
                        _copiar_celula_base_reduzida(cell, destino)

                tabela_bytes = io.BytesIO()
                wb_tabela.save(tabela_bytes)
                tabela_bytes.seek(0)
                arquivos_tabelas.append((f"tabela_{inicio_tabela}.xlsx", tabela_bytes))

            i = fim_tabela
        else:
            i += 1

    consolidado_bytes = io.BytesIO()
    wb_consolidado.save(consolidado_bytes)
    consolidado_bytes.seek(0)

    return consolidado_bytes, arquivos_tabelas


def processador_base_reduzida():
    """
    Tela do Processador de Base Reduzida: upload de um Excel com várias
    tabelas empilhadas numa aba, remoção das colunas com base < 20 e
    download da planilha consolidada e/ou das tabelas separadas em .zip.
    """
    st.header("Processador de Base Reduzida")
    st.caption(
        "Remove, de cada tabela do arquivo, as colunas cuja 'Base reduzida' "
        "seja menor que 20, preservando a formatação original."
    )

    uploaded_file = st.file_uploader(
        "Faça upload do arquivo Excel", type=["xlsx"], key="upload_base_reduzida"
    )

    if not uploaded_file:
        st.info("Envie um arquivo .xlsx para começar.")
        return

    wb = openpyxl.load_workbook(uploaded_file)
    abas = wb.sheetnames
    aba = st.selectbox("Selecione a aba para processar:", abas, key="aba_base_reduzida")

    if st.button("Processar", key="processar_base_reduzida"):
        with st.spinner("Processando tabelas..."):
            consolidado_bytes, arquivos_tabelas = _processar_tabelas_base_reduzida(wb, aba)

        st.session_state["br_consolidado_bytes"] = consolidado_bytes.getvalue()

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for nome, tabela_bytes in arquivos_tabelas:
                zip_file.writestr(nome, tabela_bytes.read())
        zip_buffer.seek(0)
        st.session_state["br_zip_bytes"] = zip_buffer.getvalue()

        st.success(f"Processamento concluído! {len(arquivos_tabelas)} tabela(s) encontrada(s).")

    if "br_consolidado_bytes" in st.session_state:
        st.download_button(
            label="Baixar planilha consolidada",
            data=st.session_state["br_consolidado_bytes"],
            file_name="Bases reduzidas - Consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_br_consolidado"
        )

    if "br_zip_bytes" in st.session_state:
        st.download_button(
            label="Baixar todas as tabelas processadas (ZIP)",
            data=st.session_state["br_zip_bytes"],
            file_name="Tabelas_processadas.zip",
            mime="application/zip",
            key="download_br_zip"
        )

