"""
Relatório Automatizado.

Orquestra, em sequência e com as perguntas necessárias em cada etapa, o
fluxo de finalização de um relatório descritivo — equivalente aos 15
códigos VBA usados manualmente hoje (ver conversa de origem). Trabalha
sempre em cima de UM ÚNICO arquivo, que vai sendo atualizado passo a
passo conforme a pessoa avança no assistente.

ESTADO ATUAL (fase 1 — esqueleto): o fluxo completo (perguntas,
ramificações, uploads) já funciona de ponta a ponta. Os códigos que ainda
não foram portados para Python entram como "placeholder" — aparecem no
log como executados, mas não alteram o arquivo ainda. Duas etapas já
usam lógica real, reaproveitada de módulos já prontos do app:
  - "Legendas" (chama a mesma lógica de core/legendas_math.py)
  - a pergunta de Base Reduzida (etapa do código 05) já grava o parâmetro
    escolhido, mas a exclusão de colunas em si ainda é placeholder.

Cada código será substituído por lógica real nas próximas etapas deste
projeto, uma de cada vez.

Não depende do `dados` carregado no início do app: tem upload próprio.
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


# =========================================================
# Utilidades de estado do assistente
# =========================================================
def _log(msg):
    st.session_state.setdefault("ra_log", []).append(msg)


def _ir_para(etapa):
    st.session_state["ra_step"] = etapa
    st.rerun()


def _reiniciar():
    for chave in list(st.session_state.keys()):
        if chave.startswith("ra_"):
            del st.session_state[chave]
    st.rerun()


def _stub(numero, nome):
    """
    Placeholder para um código ainda não portado para Python: não altera
    o arquivo, só registra no log que essa etapa foi (ou seria) aplicada.
    Será substituído pela lógica real de cada código, um de cada vez.
    """
    _log(f"Código {numero} ({nome}): placeholder — arquivo não alterado ainda.")


# =========================================================
# Assistente
# =========================================================
def modulo_relatorio_automatizado():
    st.header("Relatório Automatizado")

    if "ra_step" not in st.session_state:
        st.session_state["ra_step"] = "aviso"
    if "ra_log" not in st.session_state:
        st.session_state["ra_log"] = []

    if st.session_state["ra_step"] != "aviso":
        if st.button("🔄 Reiniciar do zero", key="ra_reiniciar"):
            _reiniciar()

    etapa = st.session_state["ra_step"]

    # ---------------------------------------------------------------
    if etapa == "aviso":
        st.warning(
            "Antes de começar: importe aqui o **relatório descritivo já "
            "com as bases nas múltiplas aplicadas** (rode o módulo "
            "\"Base nas Múltiplas\" antes, se ainda não tiver feito isso)."
        )
        if st.button("Entendi, continuar", key="ra_aviso_continuar"):
            _ir_para("upload")
        return

    # ---------------------------------------------------------------
    if etapa == "upload":
        arquivo = st.file_uploader(
            "Relatório descritivo (.xlsx)", type=["xlsx"], key="ra_upload_relatorio"
        )
        if arquivo is not None:
            st.session_state["ra_wb_bytes"] = arquivo.read()
            st.session_state["ra_nome_arquivo"] = arquivo.name
            _log("Relatório carregado.")
            _ir_para("origem")
        return

    # A partir daqui, sempre existe um arquivo em andamento em
    # st.session_state["ra_wb_bytes"] — mostramos isso pra situar a pessoa
    st.caption(f"Arquivo em andamento: **{st.session_state.get('ra_nome_arquivo', '?')}**")

    # ---------------------------------------------------------------
    if etapa == "origem":
        origem = st.radio(
            "O relatório foi gerado no SPSS Relatoria ou no SPSS Inovação?",
            ["SPSS Relatoria", "SPSS Inovação"],
            key="ra_origem"
        )
        if st.button("Continuar", key="ra_origem_continuar"):
            if origem == "SPSS Inovação":
                _stub("01", "Preenche células do SPSS pirata")
                _stub("02", "Muda Bordas")
            else:
                _log("SPSS Relatoria selecionado — códigos 01 e 02 pulados.")
            _ir_para("novos_termos")
        return

    # ---------------------------------------------------------------
    if etapa == "novos_termos":
        st.subheader("Ordenar tabelas (código 03)")
        precisa = st.radio(
            "O relatório precisa de novos termos na lista de exclusão da "
            "ordenação (termos que NÃO devem ser reordenados, ex.: 'Base', "
            "'Não sabe', nomes de candidato)?",
            ["Não", "Sim"],
            key="ra_precisa_termos"
        )
        termos_novos = ""
        if precisa == "Sim":
            termos_novos = st.text_area(
                "Cole aqui os termos novos (um por linha) a adicionar à "
                "lista de exclusão:",
                key="ra_termos_novos",
                height=150,
            )
            st.caption(
                "⚠️ Por enquanto isso só fica registrado — a etapa de "
                "ordenação em si ainda é um placeholder (código 03 ainda "
                "não foi portado)."
            )
        if st.button("Continuar", key="ra_termos_continuar"):
            if precisa == "Sim":
                lista = [t.strip() for t in termos_novos.splitlines() if t.strip()]
                st.session_state["ra_termos_novos_lista"] = lista
                _log(f"Código 03 (Ordenar tabelas): {len(lista)} termo(s) novo(s) registrado(s) — placeholder.")
            else:
                _stub("03", "Ordenar tabelas")
            _ir_para("codigo_04")
        return

    # ---------------------------------------------------------------
    if etapa == "codigo_04":
        _stub("04", "Formatação de rótulos/quebras de texto")
        _ir_para("base_reduzida")
        return

    # ---------------------------------------------------------------
    if etapa == "base_reduzida":
        st.subheader("Formata layout / Base reduzida (código 05)")
        limite = st.number_input(
            "Qual o limite mínimo de Base reduzida? "
            "(colunas com base menor que esse valor serão removidas)",
            min_value=1, value=25, step=1, key="ra_limite_base_reduzida"
        )
        if st.button("Aplicar e continuar", key="ra_base_reduzida_continuar"):
            st.session_state["ra_limite_base_reduzida_escolhido"] = limite
            _log(
                f"Código 05 (Formata layout / Base reduzida): parâmetro "
                f"escolhido = {limite} — placeholder (exclusão de colunas "
                "ainda não aplicada de verdade)."
            )
            _ir_para("codigo_06_07")
        return

    # ---------------------------------------------------------------
    if etapa == "codigo_06_07":
        _stub("06", "Ajuste DIN 9 e autoajuste da linha")
        _stub("07", "Adiciona quebra de página")
        _ir_para("legendas")
        return

    # ---------------------------------------------------------------
    if etapa == "legendas":
        st.subheader("Legendas")
        st.caption(
            "Envie o arquivo com a(s) legenda(s) a serem usadas neste "
            "relatório (mesmo formato do módulo \"Legendas\")."
        )
        arquivo_legenda = st.file_uploader(
            "Arquivo de legendas (.xlsx)", type=["xlsx"], key="ra_upload_legenda"
        )

        col1, col2 = st.columns([1, 1])
        with col1:
            aplicar = st.button("Aplicar legendas e continuar", key="ra_legendas_aplicar")
        with col2:
            pular = st.button("Este relatório não precisa de legenda", key="ra_legendas_pular")

        if pular:
            _log("Legendas: etapa pulada (relatório sem tabelas territoriais).")
            _ir_para("codigo_08_12")
            return

        if aplicar:
            if not arquivo_legenda:
                st.warning("Envie o arquivo de legendas antes de continuar.")
                return
            with st.spinner("Aplicando legendas..."):
                try:
                    wb_atual_bytes = io.BytesIO(st.session_state["ra_wb_bytes"])
                    blocos_tabela = parsear_blocos_tabela(wb_atual_bytes)
                    blocos_elegiveis = [b for b in blocos_tabela if bloco_precisa_legenda(b)]

                    if not blocos_elegiveis:
                        st.info(
                            "Não encontrei nenhum bloco segmentado por território "
                            "neste relatório — seguindo sem aplicar legenda."
                        )
                        _log("Legendas: nenhum bloco territorial encontrado, nada a fazer.")
                        _ir_para("codigo_08_12")
                        return

                    blocos_legenda = parsear_blocos_legenda(arquivo_legenda)
                    pares, avisos = parear_blocos(blocos_elegiveis, blocos_legenda)

                    wb_atual_bytes = io.BytesIO(st.session_state["ra_wb_bytes"])
                    wb_novo = gerar_workbook_com_legendas(wb_atual_bytes, pares)

                    saida = io.BytesIO()
                    wb_novo.save(saida)
                    st.session_state["ra_wb_bytes"] = saida.getvalue()
                except Exception as e:
                    st.error(f"Não foi possível aplicar as legendas: {e}")
                    return

            for aviso in avisos:
                st.warning(aviso)
            _log(f"Legendas: {len(pares)} tabela(s) receberam legenda.")
            _ir_para("codigo_08_12")
        return

    # ---------------------------------------------------------------
    if etapa == "codigo_08_12":
        _stub("08", "Base pequena")
        _stub("09", "Ajusta a altura das labels")
        _stub("10", "Ajusta o título da Renda")
        _stub("11", "Ajusta a altura dos Títulos")
        _stub("12", "Ajuste da altura das Perguntas")
        _ir_para("etapa_13")
        return

    # ---------------------------------------------------------------
    if etapa == "etapa_13":
        st.subheader("Corrigir cabeçalhos repetidos (código 13)")
        st.caption(
            "Esta etapa ainda está em desenho — no VBA original, você "
            "seleciona com o mouse um cabeçalho já corrigido e um antigo, "
            "e o código replica a correção em todos os blocos iguais. "
            "Ainda vamos decidir juntos como isso funciona no app."
        )
        col1, col2 = st.columns([1, 1])
        with col1:
            pular13 = st.button("Pular esta etapa por enquanto", key="ra_pular_13")
        if pular13:
            _log("Código 13 (Duplo click / cabeçalhos repetidos): pulado (etapa ainda em desenho).")
            _ir_para("codigo_14")
        return

    # ---------------------------------------------------------------
    if etapa == "codigo_14":
        _stub("14", "AplicarCabecalho")
        _ir_para("codigo_15")
        return

    # ---------------------------------------------------------------
    if etapa == "codigo_15":
        _stub("15", "InserirCapasResultados")
        _ir_para("final")
        return

    # ---------------------------------------------------------------
    if etapa == "final":
        st.success("Fluxo concluído!")
        st.subheader("Log das etapas")
        for linha in st.session_state.get("ra_log", []):
            st.write(f"- {linha}")

        st.download_button(
            "Baixar relatório processado",
            data=st.session_state["ra_wb_bytes"],
            file_name=f"processado_{st.session_state.get('ra_nome_arquivo', 'relatorio.xlsx')}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="ra_download_final"
        )
        return
