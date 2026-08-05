"""
Funções de carregamento e pré-processamento de dados (CSV, Excel, SPSS).
"""
import os
import tempfile

import pandas as pd
import pyreadstat
import streamlit as st


def carregar_dados(uploaded_file):
    """
    Carrega dados de CSV, Excel ou SPSS (SAV), preservando valores numéricos
    e associando labels apenas para exibição.
    """
    try:
        if uploaded_file.name.endswith(".csv"):
            try:
                dados = pd.read_csv(uploaded_file, encoding="utf-8")
            except UnicodeDecodeError:
                st.warning("Codificação 'utf-8' falhou. Tentando com 'latin1'.")
                dados = pd.read_csv(uploaded_file, encoding="latin1")

            st.success(f"Dados carregados com sucesso! Número de registros: {dados.shape[0]}")
            st.write("Visualização dos dados:", dados.head())
            return dados, dados

        elif uploaded_file.name.endswith(".xlsx"):
            dados = pd.read_excel(uploaded_file)
            st.success(f"Dados carregados com sucesso! Número de registros: {dados.shape[0]}")
            st.write("Visualização dos dados:", dados.head())
            return dados, dados

        elif uploaded_file.name.endswith(".sav"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            dados, meta = pyreadstat.read_sav(tmp_path)
            os.unlink(tmp_path)

            label_dict = {}
            for var in meta.variable_value_labels:
                label_dict[var] = {
                    int(k): v for k, v in meta.variable_value_labels[var].items()
                }

            dados_exibicao = dados.copy()
            for col in label_dict.keys():
                if col in dados.columns:
                    dados_exibicao[col] = (
                        dados[col].map(label_dict[col]).fillna(dados[col])
                    )

            st.success(f"Dados carregados com sucesso! Número de registros: {dados.shape[0]}")
            st.write("Visualização com labels:", dados_exibicao.head())
            st.write("Dados numéricos usados nos cálculos:", dados.head())

            return dados, dados_exibicao

        else:
            st.error("Formato de arquivo não suportado. Use CSV, Excel ou SPSS (.sav).")
            return None, None

    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return None, None


def validar_arquivo(arquivo):
    """
    Função para validar se o arquivo carregado contém dados válidos.
    """
    if arquivo is not None:
        try:
            dados = pd.read_csv(arquivo)
            if dados.empty:
                st.error("O arquivo está vazio. Por favor, carregue um arquivo válido.")
                return None
            return dados
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
            return None
    else:
        st.warning("Nenhum arquivo carregado.")
        return None


def preprocessar_dados(dados):
    for coluna in dados.columns:
        if dados[coluna].apply(
            lambda x: str(x).replace(".", "", 1).isdigit() if pd.notnull(x) else False
        ).all():
            dados[coluna] = pd.to_numeric(dados[coluna], errors="coerce")
        else:
            dados[coluna] = dados[coluna].astype(str)
    return dados

