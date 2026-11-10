"""
Padronização de Títulos — consulta ao catálogo de títulos padrão por tipo de pergunta.

A pessoa procura o tipo de pergunta (busca por palavra ou por grupo), preenche as partes que variam
(cargo, candidato, cidade...) e a tela mostra o título padrão já com o sufixo do tipo de resposta, a pergunta
correspondente e as duas linhas prontas para o .sps (`/TITLE` e `/CAPTION`). Também lista como aquele título
já apareceu escrito nos relatórios e baixa o catálogo completo em Excel.

Por ora é só consulta; aplicar os padrões nos arquivos fica para uma etapa seguinte.
Toda a lógica mora em core/padronizacao_titulos_math.py. Chaves de session_state com prefixo `pt_`.
"""
import pandas as pd
import streamlit as st

from core.padronizacao_titulos_math import (
    GRUPOS,
    PADROES,
    buscar,
    catalogo_como_linhas,
    catalogo_excel,
    linha_caption_sps,
    linha_title_sps,
    montar_pergunta,
    montar_titulo,
)

_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_TODOS = "Todos os grupos"
_OUTRO = "Outro (digitar)"


def _valor_do_campo(campo, chave_base: str) -> str:
    """Campo com lista usual vira seleção (com «Outro» para digitar); os demais, texto livre."""
    if campo.opcoes:
        escolha = st.selectbox(
            campo.rotulo, list(campo.opcoes) + [_OUTRO], key=f"{chave_base}_sel",
        )
        if escolha != _OUTRO:
            return escolha
        return st.text_input(f"{campo.rotulo} (digitar)", key=f"{chave_base}_txt", placeholder=campo.exemplo)
    return st.text_input(campo.rotulo, key=f"{chave_base}_txt", placeholder=f"Ex.: {campo.exemplo}")


def modulo_padronizacao_titulos():
    st.caption(
        "Consulte o título padrão de cada tipo de pergunta, já com «(Estimulada e Única)», «(Espontânea e Única)» "
        "ou «(Estimulada e Múltipla)». Levantado nos Relatórios Descritivos de 6 projetos; onde eles divergiam, "
        "vale a forma mais frequente (/2026, Cenário Único, Presidente da República...)."
    )

    # ---------------- 1. procurar o tipo de pergunta
    st.subheader("1. Qual é a pergunta?")
    c1, c2 = st.columns([3, 2])
    consulta = c1.text_input(
        "Buscar", key="pt_busca", placeholder="Ex.: rejeição, ipv, governo estadual, rádio, segundo turno",
    )
    grupo = c2.selectbox("Grupo", [_TODOS] + list(GRUPOS), key="pt_grupo")
    achados = buscar(consulta, None if grupo == _TODOS else grupo)
    if not achados:
        st.warning("Nenhum tipo de pergunta encontrado. Tente outra palavra ou escolha «Todos os grupos».")
    else:
        st.caption(f"{len(achados)} tipo(s) de pergunta" + (f" para «{consulta.strip()}»" if consulta.strip() else "") + ".")
        ids = [p.id for p in achados]
        mapa = {p.id: p for p in achados}
        escolhido = st.selectbox(
            "Tipo de pergunta", ids, key="pt_tipo",
            format_func=lambda i: f"{mapa[i].grupo} · {mapa[i].nome}",
        )
        padrao = mapa[escolhido]

        # ---------------- 2. preencher o que varia
        st.subheader("2. Preencha o que muda")
        st.caption("Modelo: " + montar_titulo(padrao, {}, parcial=True))
        valores = {}
        if padrao.campos:
            colunas = st.columns(min(3, len(padrao.campos)))
            for k, campo in enumerate(padrao.campos):
                with colunas[k % len(colunas)]:
                    valores[campo.nome] = _valor_do_campo(campo, f"pt_{padrao.id}_{campo.nome}")
        else:
            st.caption("Este título não tem parte variável.")

        # ---------------- 3. resultado
        st.subheader("3. Título padrão")
        titulo = montar_titulo(padrao, valores, parcial=True)
        pergunta = montar_pergunta(padrao, valores, parcial=True)
        faltam = [c.rotulo for c in padrao.campos if not str(valores.get(c.nome) or "").strip()]
        if faltam:
            st.info("Falta preencher: " + ", ".join(faltam) + " (os «‹campo›» abaixo mostram onde entra).")
        st.markdown("**Título**")
        st.code(titulo, language=None)
        if pergunta:
            st.markdown("**Pergunta**")
            st.code(pergunta, language=None)
            st.markdown("**Linhas para o .sps**")
            st.code(linha_title_sps(titulo) + "\n" + linha_caption_sps(pergunta), language=None)
        else:
            st.markdown("**Linha para o .sps**")
            st.code(linha_title_sps(titulo), language=None)
        if padrao.observacao:
            st.caption("Obs.: " + padrao.observacao)

        if padrao.variantes or padrao.perguntas_vistas:
            with st.expander("Como já apareceu nos relatórios"):
                if padrao.variantes:
                    st.markdown("**Outras formas do título** (não são o padrão)")
                    st.dataframe(
                        pd.DataFrame(padrao.variantes, columns=["Onde", "Título como estava"]),
                        hide_index=True, use_container_width=True,
                    )
                if padrao.perguntas_vistas:
                    st.markdown("**Outras redações da pergunta**")
                    st.dataframe(
                        pd.DataFrame(padrao.perguntas_vistas, columns=["Onde", "Pergunta"]),
                        hide_index=True, use_container_width=True,
                    )

    # ---------------- catálogo completo
    st.subheader("Catálogo completo")
    st.caption(f"{len(PADROES)} tipos de pergunta, em {len(GRUPOS)} grupos.")
    st.download_button(
        "Baixar catálogo (Excel)", data=catalogo_excel(), file_name="Padrões de Título.xlsx",
        mime=_MIME_XLSX, key="pt_download",
    )
    tabela = pd.DataFrame(catalogo_como_linhas(achados if achados else None))
    if not tabela.empty:
        st.dataframe(
            tabela[["Grupo", "Tipo de pergunta", "Título padrão (modelo)", "Exemplo de título"]],
            hide_index=True, use_container_width=True,
        )
    st.caption("A tabela acompanha a busca e o grupo escolhidos acima; o Excel traz todos, com a pergunta e as outras formas já vistas.")
