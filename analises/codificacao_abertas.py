"""Interface Streamlit da Codificação Automática de Abertas (candidatos).

Sobe um banco `.sav` com uma pergunta aberta sobre candidato (Presidente,
Governador, Senador...) + uma base de candidatos (ex.: TSE), casa cada
resposta de texto livre contra os candidatos do cargo e gera a variável
codificada (Syntax.sps) + tabela de frequência — sem nenhuma chamada de
API, é fuzzy matching + regras determinísticas.
"""
from __future__ import annotations

import os
import tempfile

import pandas as pd
import streamlit as st

from core.codificacao_abertas_math import (
    ler_base_candidatos,
    filtrar_candidatos,
    classificar_banco,
    gerar_syntax,
)

LIMIAR_REVISAO = 0.85  # casos abaixo disso entram na tabela "pra revisar"


def _ler_sav_temp(uploaded_file):
    from core.pyreadstat_seguro import pyreadstat

    with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
        tmp.write(uploaded_file.getvalue())
        caminho = tmp.name
    try:
        df, meta = pyreadstat.read_sav(caminho, apply_value_formats=False)
        return df, meta
    finally:
        os.unlink(caminho)


def _variaveis_texto(df, meta):
    """Lista as colunas que parecem ser de texto livre (dtype object no
    pandas) — é o tipo que uma pergunta aberta do SPSS vira ao ler com
    pyreadstat. Mostra todas se nenhuma bater esse critério, pra não
    travar o fluxo por causa de um `.sav` fora do padrão."""
    candidatas = [c for c in df.columns if df[c].dtype == object]
    return candidatas or list(df.columns)


def modulo_codificacao_abertas():
    st.title("Codificação Automática de Abertas — Candidatos")
    st.caption(
        "Sobe o banco `.sav` com a pergunta aberta sobre candidato "
        "(Presidente, Governador, Senador, Deputado, Prefeito, "
        "Vereador...) e uma base de candidatos (ex.: TSE) — casa cada "
        "resposta de texto livre contra os candidatos do cargo por "
        "nome de urna, nome completo e nome social. Classifique uma "
        "variável de cada vez e adicione à fila; quando terminar, "
        "baixe tudo junto num único `Syntax.sps`. 100% determinístico "
        "(fuzzy matching + regras) — sem chamada de API."
    )

    for chave in ("cab_candidatos_por_cargo", "cab_cargos", "cab_ufs", "cab_resultado"):
        if chave not in st.session_state:
            st.session_state[chave] = None
    if "cab_fila" not in st.session_state:
        st.session_state["cab_fila"] = []

    col1, col2 = st.columns(2)
    with col1:
        arquivo_sav = st.file_uploader("Banco SPSS (.sav)", type=["sav"], key="cab_upload_sav")
    with col2:
        arquivo_tse = st.file_uploader(
            "Base de candidatos (ex.: TSE, .xlsx)", type=["xlsx"], key="cab_upload_tse"
        )

    if arquivo_tse is not None and st.session_state["cab_candidatos_por_cargo"] is None:
        try:
            candidatos_por_cargo, cargos, ufs = ler_base_candidatos(arquivo_tse)
            st.session_state["cab_candidatos_por_cargo"] = candidatos_por_cargo
            st.session_state["cab_cargos"] = cargos
            st.session_state["cab_ufs"] = ufs
        except Exception as e:
            st.error(f"Não consegui ler a base de candidatos: {e}")
            return

    if arquivo_sav is None or arquivo_tse is None:
        st.info("Envie os dois arquivos pra continuar.")
        return

    cargos = st.session_state["cab_cargos"] or []
    if not cargos:
        st.error("Não encontrei nenhum cargo na base de candidatos.")
        return

    cargo = st.selectbox("Cargo", cargos, key="cab_cargo")

    uf = None
    if cargo.upper() != "PRESIDENTE":
        ufs = st.session_state["cab_ufs"] or []
        st.warning(
            f"'{cargo}' é um cargo estadual/municipal — candidatos com o "
            "mesmo nome podem existir em estados diferentes sem relação "
            "nenhuma entre si. Filtre por UF abaixo (esta versão "
            "classifica contra a lista de UM estado por vez, não cruza "
            "automaticamente com a UF de cada respondente)."
        )
        uf = st.selectbox("UF", ["(todas — cuidado com ambiguidade)"] + ufs, key="cab_uf")
        if uf == "(todas — cuidado com ambiguidade)":
            uf = None

    candidatos = filtrar_candidatos(st.session_state["cab_candidatos_por_cargo"], cargo, uf)
    if not candidatos:
        st.error(f"Não encontrei candidatos pra {cargo}" + (f" em {uf}" if uf else "") + ".")
        return
    st.caption(f"{len(candidatos)} candidato(s) carregado(s) pra {cargo}" + (f" ({uf})" if uf else "") + ".")

    try:
        df, meta = _ler_sav_temp(arquivo_sav)
    except Exception as e:
        st.error(f"Não consegui ler o banco .sav: {e}")
        return

    variaveis = _variaveis_texto(df, meta)
    col3, col4 = st.columns(2)
    with col3:
        var_origem = st.selectbox("Variável com a resposta aberta (ex.: Q1)", variaveis, key="cab_var_origem")
    with col4:
        var_destino = st.text_input(
            "Nome da variável codificada", value=f"{var_origem}A", key="cab_var_destino"
        )

    with st.expander("Parâmetros avançados"):
        limiar_confianca = st.slider(
            "Limiar de confiança do fuzzy matching", min_value=0.5, max_value=1.0,
            value=0.72, step=0.01, key="cab_limiar",
            help="Respostas com o melhor candidato abaixo desse score caem em 'Outros'.",
        )
        base_minima = st.number_input(
            "Base mínima pra criar categoria própria", min_value=1, value=3, key="cab_base_minima",
            help="Candidato/Família com menos casos que isso é agrupado em 'Outros'.",
        )

    if st.button("Classificar e gerar Syntax", key="cab_classificar"):
        with st.status("Classificando as respostas...", expanded=True) as status:
            valores = df[var_origem].tolist()
            status.update(label=f"Casando {len(valores)} respostas contra {len(candidatos)} candidatos...")
            resultado = classificar_banco(
                valores, candidatos, limiar_confianca=limiar_confianca, base_minima=base_minima,
            )
            status.update(label="Gerando a Syntax...")
            syntax_texto = gerar_syntax(
                var_origem, var_destino, resultado["mapa_valor_categoria"], resultado["codigos"],
                rotulo_cargo=cargo + (f" ({uf})" if uf else ""),
            )
            status.update(label="Concluído!", state="complete")

        st.session_state["cab_resultado"] = resultado
        st.session_state["cab_syntax_texto"] = syntax_texto

    resultado = st.session_state.get("cab_resultado")
    if resultado:
        st.subheader("Frequência")
        df_freq = pd.DataFrame(resultado["frequencia"])
        df_freq.columns = ["Categoria", "Código", "N", "%"]
        st.dataframe(df_freq, use_container_width=True, hide_index=True)

        ja_na_fila = any(item["variavel_destino"] == var_destino for item in st.session_state["cab_fila"])
        rotulo_botao = "🔄 Atualizar na fila" if ja_na_fila else "➕ Adicionar à fila"
        if st.button(rotulo_botao, key="cab_adicionar_fila"):
            n_categorias = sum(
                1 for f in resultado["frequencia"]
                if f["categoria"] not in ("Não sabe", "Branco/Nulo", "Não irá votar", "Outros")
            )
            novo_item = {
                "variavel_origem": var_origem,
                "variavel_destino": var_destino,
                "cargo_label": cargo + (f" ({uf})" if uf else ""),
                "syntax": st.session_state["cab_syntax_texto"],
                "n_casos": len(resultado["resultado_por_caso"]),
                "n_categorias": n_categorias,
            }
            st.session_state["cab_fila"] = [
                item for item in st.session_state["cab_fila"] if item["variavel_destino"] != var_destino
            ]
            st.session_state["cab_fila"].append(novo_item)
            st.success(f"{var_destino} {'atualizado na' if ja_na_fila else 'adicionado à'} fila.")

        st.subheader("Casos pra revisar manualmente")
        casos_revisao = [
            r for r in resultado["resultado_por_caso"]
            if r["confianca"] < LIMIAR_REVISAO and r["categoria"] not in ("Não sabe", "Branco/Nulo", "Não irá votar")
        ]
        if casos_revisao:
            st.caption(
                f"{len(casos_revisao)} caso(s) com confiança abaixo de {LIMIAR_REVISAO:.0%} — "
                "vale conferir antes de adicionar à fila."
            )
            df_revisao = pd.DataFrame([
                {"Resposta original": r["valor_original"], "Categoria sugerida": r["categoria"],
                 "Confiança": r["confianca"], "Motivo": r["motivo"]}
                for r in casos_revisao
            ]).drop_duplicates().sort_values("Confiança")
            st.dataframe(df_revisao, use_container_width=True, hide_index=True)
        else:
            st.success("Nenhum caso abaixo do limiar de revisão.")

    # A fila aparece sempre — independente de ter classificado algo nesta
    # execução — senão ela "sumiria" da tela toda vez que a pessoa só
    # trocasse a variável/cargo sem apertar "Classificar" de novo.
    st.divider()
    st.subheader(f"Fila de códigos ({len(st.session_state['cab_fila'])})")
    st.caption(
        "Classifique quantas variáveis quiser (troque a variável/cargo acima e clique "
        "em 'Adicionar à fila' de novo) — quando terminar, baixe tudo junto num único "
        "arquivo .sps."
    )

    if not st.session_state["cab_fila"]:
        st.info("Fila vazia — classifique uma variável acima e adicione à fila.")
        return

    for i, item in enumerate(st.session_state["cab_fila"]):
        col_info, col_remover = st.columns([6, 1])
        with col_info:
            st.write(
                f"**{item['variavel_destino']}** — {item['variavel_origem']} → "
                f"{item['cargo_label']} · {item['n_categorias']} categoria(s) · "
                f"{item['n_casos']} caso(s)"
            )
        with col_remover:
            if st.button("Remover", key=f"cab_remover_fila_{i}"):
                st.session_state["cab_fila"].pop(i)
                st.rerun()

    col_baixar, col_limpar = st.columns(2)
    with col_baixar:
        syntax_completa = "\n\n".join(item["syntax"] for item in st.session_state["cab_fila"])
        st.download_button(
            "⬇️ Baixar tudo (Syntax.sps)", data=syntax_completa, file_name="Syntax_Codificacao.sps",
            mime="text/plain", key="cab_download_fila", use_container_width=True,
        )
    with col_limpar:
        if st.button("🗑️ Limpar fila", key="cab_limpar_fila", use_container_width=True):
            st.session_state["cab_fila"] = []
            st.rerun()
