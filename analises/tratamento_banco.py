"""
Tratamento do Banco — etapa 1.

Sobe a Syntax (.sps) e o banco original (.sav). O app roda a recodificação (COMPUTE, RECODE, rótulos,
conjuntos múltiplos), confere se deu certo e, se o banco já tratado pelo SPSS for enviado, compara variável
a variável. Quando está tudo certo, baixa o banco tratado em .sav.

As etapas seguintes (rodar o Total e o Total e Completo e salvar o Excel) entram depois.
Toda a lógica mora em core/tratamento_banco_math.py e core/syntax_executor_math.py.
Chaves de session_state com prefixo `tb_`.
"""
import hashlib

import pandas as pd
import streamlit as st

from core.analitica_dados import carregar_banco
from core.revisor_syntax_math import SEVERITY_LABELS
from core.tratamento_banco_math import (
    exportar_sav,
    nome_arquivo_tratado,
    tratar_banco,
)

_ICONE = {"critico": "🔴", "possivel_erro": "🟠", "revisar": "🟡", "informativo": "🔵"}
_STATUS = {"ok": "✅ rodou", "ignorado": "⚪ sem efeito", "nao_suportado": "⚠️ não executado", "erro": "❌ erro"}


def _hash(*partes: bytes) -> str:
    h = hashlib.md5()
    for p in partes:
        h.update(p or b"")
    return h.hexdigest()


def _tabela_achados(achados):
    linhas = []
    for f in achados:
        linhas.append({
            "Gravidade": f"{_ICONE.get(f.severity, '')} {SEVERITY_LABELS.get(f.severity, f.severity)}",
            "Linha": f.line if f.line else "",
            "Variável": f.variable or "",
            "O que foi encontrado": f.title,
            "Detalhe": f.message,
            "Como resolver": f.suggestion or "",
        })
    return pd.DataFrame(linhas)


def _mostrar_achados(achados, vazio, chave):
    if not achados:
        st.success(vazio)
        return
    niveis = [k for k in ("critico", "possivel_erro", "revisar", "informativo") if any(f.severity == k for f in achados)]
    escolhidos = st.multiselect(
        "Mostrar", niveis, default=[n for n in niveis if n != "informativo"] or niveis,
        format_func=lambda k: f"{_ICONE[k]} {SEVERITY_LABELS[k]} ({sum(1 for f in achados if f.severity == k)})",
        key=chave,
    )
    sel = [f for f in achados if f.severity in escolhidos]
    if sel:
        st.dataframe(_tabela_achados(sel), use_container_width=True, hide_index=True)
    else:
        st.caption("Nada nos níveis escolhidos.")


def modulo_tratamento_banco():
    st.caption(
        "Roda a Syntax no banco original (recodificações e rótulos), confere se a recodificação ficou certa e, "
        "se você enviar o banco que o SPSS gerou, compara variável a variável. Tudo certo, baixa o banco tratado."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        sps = st.file_uploader("1) Syntax (.sps)", type=["sps"], key="tb_up_sps")
    with c2:
        sav = st.file_uploader("2) Banco original (.sav)", type=["sav"], key="tb_up_sav")
    with c3:
        sav_spss = st.file_uploader("3) Banco tratado pelo SPSS (.sav) — opcional, para comparar",
                                    type=["sav"], key="tb_up_spss")

    zero = st.checkbox(
        "Rodar do zero: se o banco já tiver as variáveis que a Syntax cria, refazê-las",
        value=True, key="tb_zero",
        help="Assim a conferência vale de verdade, mesmo que o banco enviado já venha tratado.",
    )

    if not (sps and sav):
        st.info("Envie a Syntax e o banco original para começar.")
        return

    if st.button("Rodar tratamento", type="primary", key="tb_rodar"):
        try:
            with st.spinner("Rodando a Syntax e conferindo..."):
                b_sps, b_sav = sps.getvalue(), sav.getvalue()
                banco = carregar_banco(b_sav, sav.name, user_missing=True)
                spss = carregar_banco(sav_spss.getvalue(), sav_spss.name, user_missing=True) if sav_spss else None
                res = tratar_banco(b_sps, banco, spss, partir_do_zero=zero)
            st.session_state["tb_res"] = res
            st.session_state["tb_nome"] = sav.name
            st.session_state["tb_chave"] = _hash(b_sps, b_sav, sav_spss.getvalue() if sav_spss else b"", bytes([zero]))
            st.session_state.pop("tb_sav_bytes", None)
        except Exception as e:  # noqa: BLE001
            st.error(f"Não foi possível rodar o tratamento: {e}")
            return

    res = st.session_state.get("tb_res")
    if res is None:
        return

    exe = res.execucao
    cont = res.contagem

    # ---- veredito --------------------------------------------------------
    if res.aprovado:
        if res.pendencias_da_revisao:
            st.warning(
                f"Recodificação conferida, sem erro na execução. A revisão da Syntax tem "
                f"{res.pendencias_da_revisao} pendência(s) para você olhar antes de seguir."
            )
        else:
            st.success("Recodificação conferida: sem erros. O banco tratado está pronto.")
    else:
        st.error("Há problemas na recodificação. Veja a aba «Conferências» antes de usar o banco tratado.")

    m = st.columns(5)
    m[0].metric("Casos", f"{len(exe.df):,}".replace(",", "."))
    m[1].metric("Variáveis criadas", len(exe.criadas))
    m[2].metric("Erros / possíveis erros", cont.get("critico", 0) + cont.get("possivel_erro", 0))
    m[3].metric("A revisar", cont.get("revisar", 0))
    if res.comparacao is not None:
        m[4].metric("Iguais ao SPSS", f"{len(res.comparacao) - len(res.divergentes)}/{len(res.comparacao)}")
    else:
        m[4].metric("Iguais ao SPSS", "—")
    if res.erro_comparacao:
        st.error(res.erro_comparacao)

    abas = ["Conferências", "Revisão da Syntax", "Comandos executados", "Variáveis criadas"]
    if res.comparacao is not None:
        abas.insert(1, "Comparação com o SPSS")
    tabs = st.tabs(abas)
    por_nome = dict(zip(abas, tabs))

    with por_nome["Conferências"]:
        _mostrar_achados(res.achados_do_tratamento, "Nenhum problema encontrado na execução e nas conferências.", "tb_filtro_conf")

    if "Comparação com o SPSS" in por_nome:
        with por_nome["Comparação com o SPSS"]:
            linhas = []
            for c in res.comparacao:
                linhas.append({
                    "Variável": c.variavel,
                    "Resultado": "✅ igual" if c.igual else "❌ diferente",
                    "Casos diferentes": c.casos_diferentes,
                    "Rótulos de valor": "" if c.rotulos_de_valor_iguais is None else ("✅" if c.rotulos_de_valor_iguais else "❌"),
                    "Rótulo da variável": "" if c.rotulo_de_variavel_igual is None else ("✅" if c.rotulo_de_variavel_igual else "❌"),
                    "Detalhe": (c.detalhe + " " + c.exemplo).strip(),
                })
            df = pd.DataFrame(linhas)
            if res.divergentes:
                st.warning(f"{len(res.divergentes)} variável(is) diferente(s) do banco do SPSS.")
                df = df.sort_values("Resultado", ascending=False)
            else:
                st.success("Todas as variáveis criadas aqui são idênticas às do banco do SPSS (valores e rótulos).")
            st.dataframe(df, use_container_width=True, hide_index=True)

    with por_nome["Revisão da Syntax"]:
        st.caption("Mesma análise do módulo Revisor de Syntax, feita sobre o banco enviado.")
        _mostrar_achados(res.achados_da_revisao, "A revisão não encontrou nada.", "tb_filtro_rev")

    with por_nome["Comandos executados"]:
        cont_st = exe.contagem_por_status()
        st.caption(" · ".join(f"{_STATUS.get(k, k)}: {v}" for k, v in cont_st.items()))
        ver = st.radio("Mostrar", ["Só RECODE/COMPUTE", "Todos"], horizontal=True, key="tb_ver_cmds")
        regs = [r for r in exe.registros if ver == "Todos" or r.tipo in ("RECODE", "COMPUTE")]
        st.dataframe(pd.DataFrame([{
            "Linha": r.linha, "Comando": r.tipo, "Status": _STATUS.get(r.status, r.status),
            "Resumo": r.resumo, "Resultado": r.detalhe,
        } for r in regs]), use_container_width=True, hide_index=True)

    with por_nome["Variáveis criadas"]:
        linhas = []
        for nome in exe.criadas:
            col = next((c for c in exe.df.columns if str(c).upper() == nome.upper()), None)
            if col is None:
                continue
            serie = exe.df[col]
            linhas.append({
                "Variável": nome,
                "Rótulo": exe.var_labels.get(nome.upper(), ""),
                "Casos com valor": int(serie.notna().sum()),
                "Códigos": ", ".join(
                    (str(int(v)) if float(v).is_integer() else str(v)) for v in sorted(serie.dropna().unique())[:12]
                ) if pd.api.types.is_numeric_dtype(serie) else "(texto)",
            })
        st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

    # ---- download --------------------------------------------------------
    st.divider()
    if not res.aprovado:
        st.caption("O banco só deve ser usado depois de resolver os problemas acima.")
    if st.button("Preparar banco tratado (.sav)", key="tb_preparar"):
        try:
            st.session_state["tb_sav_bytes"] = exportar_sav(exe)
        except Exception as e:  # noqa: BLE001
            st.error(f"Não foi possível gerar o .sav: {e}")
    dados = st.session_state.get("tb_sav_bytes")
    if dados:
        st.download_button(
            "Baixar banco tratado (.sav)", data=dados,
            file_name=nome_arquivo_tratado(st.session_state.get("tb_nome", "Banco.sav")),
            mime="application/octet-stream", key="tb_baixar",
        )
