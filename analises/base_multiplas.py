"""
Base nas Múltiplas.

Tabelas de perguntas do tipo Múltipla (Resposta Múltipla) não saem do
SPSS com a linha de Base, já que o percentual de cada opção não segue a
lógica de uma tabela de resposta única. Como a base de entrevistas por
segmento é a mesma em qualquer pergunta do mesmo projeto, este módulo
pega a linha "Base" de um relatório de qualquer outra pergunta (que TEM
base) e a copia para as tabelas de múltiplas, casando cada coluna pelo
texto do cabeçalho (grupo + categoria) — não pela posição, já que a
ordem das colunas pode ser diferente entre os dois arquivos.

Diferente das análises de correspondência, não depende do `dados`
carregado no início do app: tem upload próprio (dois arquivos Excel).
"""
import io

import streamlit as st

from core.base_multiplas_math import (
    parsear_blocos,
    extrair_bases,
    calcular_linhas_base,
    gerar_workbook_com_base,
    bloco_eh_religiao,
)


def modulo_base_multiplas():
    st.header("Base nas Múltiplas")
    st.caption(
        "Copia a linha de Base de um relatório com base (qualquer pergunta "
        "do projeto) para as tabelas de Múltipla, casando cada coluna pelo "
        "texto do cabeçalho (ex.: 'Masculino' só recebe a base de uma "
        "coluna 'Masculino')."
    )

    col1, col2 = st.columns(2)
    with col1:
        arquivo_multiplas = st.file_uploader(
            "Relatório de Múltiplas (sem base)", type=["xlsx"], key="bm_upload_multiplas"
        )
    with col2:
        arquivo_bases = st.file_uploader(
            "Relatório com base (qualquer pergunta do mesmo projeto)",
            type=["xlsx"], key="bm_upload_bases"
        )

    if not arquivo_multiplas or not arquivo_bases:
        st.info("Envie os dois arquivos .xlsx para começar.")
        return

    if st.button("Calcular e gerar arquivo", key="bm_calcular"):
        with st.spinner("Lendo e casando as tabelas..."):
            try:
                blocos_multiplas = parsear_blocos(arquivo_multiplas)
                arquivo_bases.seek(0)
                blocos_bases = parsear_blocos(arquivo_bases)
            except Exception as e:
                st.error(f"Não foi possível ler os arquivos: {e}")
                return

            if not blocos_multiplas:
                st.error(
                    "Não encontrei nenhum bloco 'Titulo:' no relatório de "
                    "múltiplas. Confirme se o arquivo segue o formato "
                    "padrão de exportação do SPSS."
                )
                return
            if not blocos_bases:
                st.error(
                    "Não encontrei nenhum bloco 'Titulo:' no relatório com "
                    "base. Confirme se o arquivo segue o formato padrão de "
                    "exportação do SPSS."
                )
                return

            base_total, por_par, por_categoria, por_grupo_sem_categoria = extrair_bases(blocos_bases)

            if base_total is None:
                st.error(
                    "Não encontrei nenhuma linha 'Base' no relatório com "
                    "base. Confirme se esse é o arquivo certo."
                )
                return

            # índice restrito, usado só para os blocos de religião: busca
            # apenas dentro de outros blocos de religião do arquivo de
            # bases, pra nunca casar uma categoria ambígua (ex.: "Não
            # sabe") com algo de uma pergunta sem relação nenhuma
            blocos_religiao_bases = [b for b in blocos_bases if bloco_eh_religiao(b)]
            indice_religiao = None
            if blocos_religiao_bases:
                indice_religiao = extrair_bases(blocos_bases, apenas_blocos=blocos_religiao_bases)[1:]

            linhas_base = calcular_linhas_base(
                blocos_multiplas, base_total, por_par, por_categoria, por_grupo_sem_categoria,
                indice_religiao=indice_religiao,
            )

            arquivo_multiplas.seek(0)
            wb_novo = gerar_workbook_com_base(arquivo_multiplas, blocos_multiplas, linhas_base)

            saida = io.BytesIO()
            wb_novo.save(saida)
            saida.seek(0)

        st.session_state["bm_resultado_bytes"] = saida.getvalue()
        st.session_state["bm_resultado_resumo"] = [
            {
                "titulo": b["titulo"].replace("Titulo:", "").strip(),
                "colunas_preenchidas": len(lb["valores"]),
                "nao_encontradas": lb["nao_encontradas"],
                "ja_tinha_base": lb.get("ja_tinha_base", False),
                "eh_religiao": lb.get("eh_religiao", False),
            }
            for b, lb in zip(blocos_multiplas, linhas_base)
        ]
        n_ja_tinham = sum(1 for lb in linhas_base if lb.get("ja_tinha_base"))
        n_religiao = sum(1 for lb in linhas_base if lb.get("eh_religiao"))
        msg = f"Pronto! {len(blocos_multiplas)} tabela(s) no total"
        if n_ja_tinham:
            msg += f", {n_ja_tinham} já tinha(m) base e foi(ram) deixada(s) intocada(s)"
        if n_religiao:
            msg += f", {n_religiao} identificada(s) como religião (casamento restrito)"
        st.success(msg + ".")

    if "bm_resultado_bytes" in st.session_state:
        st.markdown("---")
        st.subheader("Resultado")

        total_nao_encontradas = sum(
            len(item["nao_encontradas"]) for item in st.session_state["bm_resultado_resumo"]
        )
        if total_nao_encontradas > 0:
            st.warning(
                f"{total_nao_encontradas} coluna(s) não encontraram base "
                "correspondente — confira a lista abaixo e complete manualmente "
                "no arquivo gerado, se precisar."
            )
        else:
            st.success("Todas as colunas encontraram base correspondente.")

        for item in st.session_state["bm_resultado_resumo"]:
            if item["ja_tinha_base"]:
                st.write(f"**{item['titulo']}** — já tinha Base, não foi alterada")
                continue
            selo = " 🕊️ (religião — casamento restrito)" if item["eh_religiao"] else ""
            linha = f"**{item['titulo']}**{selo} — {item['colunas_preenchidas']} coluna(s) preenchida(s)"
            if item["nao_encontradas"]:
                linha += f" | ⚠️ sem base: {', '.join(item['nao_encontradas'])}"
            st.write(linha)

        st.download_button(
            "Baixar relatório de Múltiplas com Base",
            data=st.session_state["bm_resultado_bytes"],
            file_name="Relatorio_Multiplas_com_Base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="bm_download"
        )
