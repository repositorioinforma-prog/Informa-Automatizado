"""Interface Streamlit do Revisor de Syntax SPSS."""
from __future__ import annotations

import os
import tempfile

import pandas as pd
from core.pyreadstat_seguro import pyreadstat
import streamlit as st

from core.revisor_syntax_math import review_syntax


def _ler_sav(uploaded_file):
    if uploaded_file is None:
        return None, None
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
        tmp.write(uploaded_file.getvalue())
        path = tmp.name
    try:
        df, meta = pyreadstat.read_sav(path, apply_value_formats=False)
        return df, meta
    finally:
        os.unlink(path)


def _df_findings(findings):
    if not findings:
        return pd.DataFrame(columns=["Nível", "Linha", "Variável", "Problema", "Detalhe", "Sugestão"])
    labels = {
        "critico": "🔴 Erro crítico",
        "possivel_erro": "🟠 Possível erro",
        "revisar": "🟡 Revisar",
        "informativo": "🔵 Informativo",
    }
    return pd.DataFrame([
        {
            "Nível": labels.get(f["severity"], f["severity"]),
            "Linha": f.get("line") or "",
            "Variável": f.get("variable") or "",
            "Problema": f.get("title") or "",
            "Detalhe": f.get("message") or "",
            "Sugestão": f.get("suggestion") or "",
        }
        for f in findings
    ])


def modulo_revisor_syntax():
    st.title("Revisor de Syntax SPSS")
    st.write(
        "Audita uma syntax `.sps` e, opcionalmente, cruza as recodificações com o banco `.sav` original "
        "e com um `.sav` já processado. O foco é localizar erros lógicos que podem passar sem erro de execução no SPSS."
    )

    with st.expander("O que é verificado nesta versão", expanded=False):
        st.markdown(
            """
- encoding declarado versus encoding real do arquivo;
- comandos aparentemente incompletos;
- regras `RECODE` duplicadas ou conflitantes;
- `VALUE LABELS` que não podem ser produzidos pelo `RECODE`;
- seções `Q...`/`P...` marcadas mas sem programação;
- variáveis de origem inexistentes no banco;
- valores que existem no `.sav` mas caem em `ELSE=SYSMIS` ou não são contemplados;
- variáveis usadas em `MRSETS`/`/MCGROUP` que não existem;
- variáveis criadas pela syntax que não aparecem no banco processado opcional;
- **inversão semântica**: segue cada código do banco original por
  VALUE LABEL → RECODE → VALUE LABEL de destino e confere se a
  categoria de chegada é objetivamente OUTRA categoria que já existia
  no banco (não confunde com simples reformulação de redação, ex.:
  "40 a 49 anos" → "40 a 49" não gera alerta);
- em variáveis de perfil/segmento (P1-P8, sexo, idade, renda,
  escolaridade, raça, religião, região...), só aparece alerta quando é
  uma **inversão real** — o resto do ruído comum nesse tipo de
  variável fica oculto de propósito;
- escala 0–10 recodificada em Aprovação/Regular/Reprovação que não
  segue o padrão 0–3/4–6/7–10;
- (opcional) cruzamento com o programa Total + Completo: perguntas que
  aparecem num e não no outro, e variáveis usadas na programação que
  não existem em lugar nenhum.
            """
        )

    syntax_file = st.file_uploader("1. Syntax SPSS (.sps)", type=["sps"], key="revisor_syntax_sps")
    original_sav = st.file_uploader(
        "2. Banco original (.sav) — recomendado",
        type=["sav"], key="revisor_syntax_original",
        help="Permite validar se todas as categorias reais do banco foram contempladas nas recodificações.",
    )
    processed_sav = st.file_uploader(
        "3. Banco processado (.sav) — opcional",
        type=["sav"], key="revisor_syntax_processado",
        help="Permite conferir se variáveis que a syntax cria realmente aparecem no banco depois do processamento.",
    )
    program_file = st.file_uploader(
        "4. Programa Total + Completo (.sps) — opcional",
        type=["sps"], key="revisor_syntax_programa",
        help=(
            "Precisa ter um cabeçalho tipo '*****COMPLETO*****' separando as duas "
            "partes. Permite conferir se toda pergunta do Total também está no "
            "Completo (e vice-versa) e se as variáveis usadas na programação existem."
        ),
    )

    if syntax_file is None:
        st.info("Carregue uma syntax `.sps` para iniciar a revisão.")
        return

    if st.button("Revisar Syntax", type="primary"):
        try:
            with st.status("Revisando syntax...", expanded=True) as status:
                st.write("Lendo e interpretando comandos SPSS...")
                df_original = meta_original = None
                df_processado = None

                if original_sav is not None:
                    st.write("Lendo banco original e levantando categorias reais...")
                    df_original, meta_original = _ler_sav(original_sav)
                if processed_sav is not None:
                    st.write("Lendo banco processado para conferência final...")
                    df_processado, _ = _ler_sav(processed_sav)

                st.write("Cruzando RECODE, labels, MRSETS e estrutura da syntax...")
                result = review_syntax(
                    syntax_file.getvalue(), dataframe=df_original,
                    metadata=meta_original, processed_dataframe=df_processado,
                    program_bytes=program_file.getvalue() if program_file is not None else None,
                )
                st.session_state["revisor_syntax_result"] = result
                status.update(label="Revisão concluída", state="complete", expanded=False)
        except Exception as exc:
            st.exception(exc)
            return

    result = st.session_state.get("revisor_syntax_result")
    if not result:
        return

    st.subheader("Resumo")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔴 Críticos", result["counts"]["critico"])
    c2.metric("🟠 Possíveis erros", result["counts"]["possivel_erro"])
    c3.metric("🟡 Revisar", result["counts"]["revisar"])
    c4.metric("RECODEs", result["recodes"])

    st.caption(
        f"Encoding lido: {result['encoding']} · {result['commands']} comandos · "
        f"{result['variable_labels']} VARIABLE LABELS · {result['value_label_variables']} variáveis com VALUE LABELS"
    )

    findings = result["findings"]
    if not findings:
        st.success("Nenhuma inconsistência foi encontrada pelas regras desta versão do revisor.")
        return

    df = _df_findings(findings)
    st.dataframe(df, use_container_width=True, hide_index=True)

    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar relatório da revisão (.csv)", data=csv_bytes,
        file_name="revisao_syntax_spss.csv", mime="text/csv",
    )

    st.subheader("Detalhes")
    for idx, f in enumerate(findings, 1):
        label = f"{idx}. {f['severity_label']} — {f['title']}"
        if f.get("variable"):
            label += f" [{f['variable']}]"
        if f.get("line"):
            label += f" · linha {f['line']}"
        with st.expander(label):
            st.write(f["message"])
            if f.get("suggestion"):
                st.info(f"Sugestão: {f['suggestion']}")
            if f.get("snippet"):
                st.text_area(
                    "Trecho", value=f["snippet"], height=150, disabled=True,
                    label_visibility="collapsed", key=f"revisor_syntax_snippet_{idx}",
                )
