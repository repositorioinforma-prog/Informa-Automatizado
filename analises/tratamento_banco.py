"""
Tratamento do Banco — etapa 1.

Sobe a Syntax (.sps) e o banco original (.sav). O app roda a recodificação (COMPUTE, RECODE, rótulos,
conjuntos múltiplos), confere se deu certo e, se o banco já tratado pelo SPSS for enviado, compara variável
a variável. Quando está tudo certo, baixa o banco tratado em .sav.

Etapa 2: com o banco já com peso, calcula as tabelas do Total e Completo (TABLES/CTABLES), confere cada uma
e, se o Relatório Descritivo for enviado, compara número a número. O resultado sai em Excel no layout que o SPSS
exporta (download + pré-visualização na tela, que é a mesma aba do arquivo).
Toda a lógica mora em core/tratamento_banco_math.py, core/syntax_executor_math.py e core/total_completo_math.py.
Chaves de session_state com prefixo `tb_` (etapa 1) e `tc_` (etapa 2).
"""
import hashlib

import pandas as pd
import streamlit as st

from core.analise_perguntas_math import html_sobre_fundo_branco
from core.analitica_dados import carregar_banco, eh_nome_de_peso
from core.planilha_utils import worksheet_para_html
from core.revisor_syntax_math import SEVERITY_LABELS
from core.syntax_executor_math import executar_syntax
from core.total_completo_excel_math import (
    SECOES,
    montar_excel_total_completo,
    nao_geradas,
    nome_arquivo_total_completo,
    planilha_tabelas,
)
from core.total_completo_math import (
    PESO_PADRAO,
    achados_da_comparacao_tc,
    calcular_total_e_completo,
    comparar_com_descritivo,
    ler_descritivo,
)
from core.tratamento_banco_math import (
    exportar_sav,
    nome_arquivo_tratado,
    tratar_banco,
)

_ICONE = {"critico": "🔴", "possivel_erro": "🟠", "revisar": "🟡", "informativo": "🔵"}
_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_MAX_TABELAS_PREVIA = 25      # a prévia "de todas" mostra as primeiras tabelas; o arquivo baixado tem todas
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


def _etapa_recodificacao():
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
            st.session_state["tb_sps_bytes"] = b_sps
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


# ---------------------------------------------------------------------------
# Etapa 2 — Total e Completo
# ---------------------------------------------------------------------------

def _banco_da_etapa2(arq):
    conteudo = arq.getvalue()
    chave = _hash(conteudo)
    if st.session_state.get("tc_chave_banco") == chave:
        return st.session_state["tc_banco"]
    banco = carregar_banco(conteudo, arq.name)
    st.session_state["tc_chave_banco"] = chave
    st.session_state["tc_banco"] = banco
    st.session_state.pop("tc_res", None)
    return banco


def _df_tabela(calc):
    colunas = pd.MultiIndex.from_tuples([(c.grupo, c.rotulo) for c in calc.colunas])
    linhas = [[("" if pd.isna(v) else f"{v:.1f}") for v in calc.valores[i]] for i in range(len(calc.rotulos))]
    linhas.append([f"{c.base:.0f}" for c in calc.colunas])
    return pd.DataFrame(linhas, index=list(calc.rotulos) + ["Base"], columns=colunas)


def _previa_excel(calcs):
    """Mostra as tabelas como ficam no Excel (mesma função que escreve o arquivo)."""
    ws = planilha_tabelas(calcs, como_texto=True)
    html, _trunc_l, _trunc_c = worksheet_para_html(ws, max_linhas=400)
    altura = int(min(760, max(240, ws.max_row * 21 + 50)))
    st.components.v1.html(html_sobre_fundo_branco(html), height=altura, scrolling=True)


def _bloco_excel(res):
    """Escolha das seções, geração e download do Excel do Total e Completo."""
    n_por = {s: sum(1 for c in res.calculadas if c is not None and c.tabela.secao == s) for s in SECOES}
    st.markdown("**Excel (layout do SPSS)**")
    secoes = st.multiselect(
        "Seções no Excel", list(SECOES), default=[s for s in SECOES if n_por[s]],
        format_func=lambda s: f"{s} ({n_por[s]} tabelas)", key="tc_secoes",
    )
    if not secoes:
        st.caption("Escolha ao menos uma seção.")
        return
    faltam = nao_geradas(res, secoes)
    if faltam:
        st.warning(f"{len(faltam)} tabela(s) não puderam ser calculadas e ficam fora da aba principal; "
                   "elas aparecem na aba «Não geradas» do Excel, com o motivo.")
    chave = (id(res), tuple(secoes))
    if st.button("Preparar Excel", key="tc_preparar_excel"):
        try:
            with st.spinner("Montando o Excel..."):
                st.session_state["tc_excel"] = (chave, montar_excel_total_completo(res, secoes))
        except Exception as e:  # noqa: BLE001
            st.error(f"Não foi possível gerar o Excel: {e}")
    pronto = st.session_state.get("tc_excel")
    if pronto and pronto[0] == chave:
        st.download_button(
            "Baixar Excel", data=pronto[1],
            file_name=nome_arquivo_total_completo(st.session_state.get("tc_nome_sps", ""), secoes),
            mime=_MIME_XLSX, key="tc_baixar_excel",
        )
    elif pronto:
        st.caption("A escolha mudou: clique em «Preparar Excel» de novo.")


def _etapa_total_completo():
    st.caption(
        "Calcula, no banco com peso, as tabelas do Total e Completo (percentual da coluna e «Base»), confere cada "
        "uma e, se você enviar o Relatório Descritivo, compara número a número."
    )
    c1, c2 = st.columns(2)
    with c1:
        tc = st.file_uploader("1) Total e Completo (.sps)", type=["sps"], key="tc_up_sps")
        sav = st.file_uploader("2) Banco com peso (.sav)", type=["sav"], key="tc_up_sav")
    with c2:
        sx = st.file_uploader(
            "3) Syntax (.sps) — para as perguntas de múltipla resposta",
            type=["sps"], key="tc_up_syntax",
            help="Se você já rodou a etapa 1 nesta sessão, a Syntax dela é usada quando este campo fica vazio.",
        )
        desc = st.file_uploader("4) Relatório Descritivo (.xlsx) — opcional, para comparar", type=["xlsx", "xls"], key="tc_up_desc")

    if not (tc and sav):
        st.info("Envie o Total e Completo e o banco com peso para começar.")
        return

    try:
        banco = _banco_da_etapa2(sav)
    except Exception as e:  # noqa: BLE001
        st.error(f"Não foi possível ler o banco: {e}")
        return

    candidatos = [c for c in banco.df.columns if eh_nome_de_peso(c)]
    opcoes = candidatos + ["Sem peso"]
    padrao = next((i for i, c in enumerate(opcoes) if str(c).lower() == PESO_PADRAO), 0)
    peso_sel = st.selectbox("Variável de peso", opcoes, index=padrao, key="tc_peso",
                            help="O SPSS usa o peso ligado na sessão; aqui você escolhe qual.")
    peso = None if peso_sel == "Sem peso" else peso_sel

    if st.button("Calcular Total e Completo", type="primary", key="tc_rodar"):
        try:
            with st.spinner("Calculando as tabelas e conferindo..."):
                b_tc = tc.getvalue()
                b_sx = sx.getvalue() if sx else st.session_state.get("tb_sps_bytes")
                mrsets = executar_syntax(b_sx, banco.df, banco.var_labels, banco.value_labels).mrsets if b_sx else {}
                res = calcular_total_e_completo(b_tc, banco.df, banco.value_labels, banco.var_labels, mrsets, peso)
                if desc:
                    res.comparacao = comparar_com_descritivo(res, ler_descritivo(desc.getvalue()))
                    res.findings.extend(achados_da_comparacao_tc(res.comparacao))
            st.session_state["tc_res"] = res
            st.session_state["tc_nome_sps"] = tc.name
            st.session_state.pop("tc_excel", None)
        except Exception as e:  # noqa: BLE001
            st.error(f"Não foi possível calcular: {e}")
            return

    res = st.session_state.get("tc_res")
    if res is None:
        return

    from core.revisor_syntax_math import SEVERITY_ORDER
    cont = {k: sum(1 for f in res.findings if f.severity == k) for k in SEVERITY_ORDER}
    graves = cont["critico"] + cont["possivel_erro"]
    n_dif = sum(1 for c in (res.comparacao or []) if c.status == "diferente")
    if res.comparacao is not None:
        n_igual = sum(1 for c in res.comparacao if c.status == "igual")

    if graves == 0 and res.n_calculadas == len(res.tabelas):
        st.success(f"{res.n_calculadas} tabela(s) calculada(s), sem erros."
                   + (f" {n_igual} igual(is) ao Descritivo." if res.comparacao is not None and not n_dif else ""))
    elif graves == 0:
        st.warning(f"{res.n_calculadas} de {len(res.tabelas)} tabela(s) calculada(s).")
    else:
        st.error("Há problemas nas tabelas. Veja a aba «Conferências».")

    m = st.columns(5)
    m[0].metric("Tabelas na programação", len(res.tabelas))
    m[1].metric("Calculadas", res.n_calculadas)
    m[2].metric("Erros / possíveis erros", graves)
    m[3].metric("A revisar", cont["revisar"])
    m[4].metric("Iguais ao Descritivo", "—" if res.comparacao is None else f"{n_igual}/{len(res.tabelas)}")

    _bloco_excel(res)

    abas = ["Conferências", "Tabelas", "Resumo"]
    if res.comparacao is not None:
        abas.insert(1, "Comparação com o Descritivo")
    tabs = dict(zip(abas, st.tabs(abas)))

    with tabs["Conferências"]:
        _mostrar_achados(sorted(res.findings, key=lambda f: (SEVERITY_ORDER.get(f.severity, 9), f.line or 10**9)),
                         "Nenhum problema encontrado.", "tc_filtro_conf")

    if "Comparação com o Descritivo" in tabs:
        with tabs["Comparação com o Descritivo"]:
            rot = {"igual": "✅ igual", "diferente": "❌ diferente", "so_no_app": "🟡 só na programação",
                   "so_no_descritivo": "🟡 só no Descritivo"}
            linhas = [{"Tabela": c.titulo, "Resultado": rot.get(c.status, c.status), "Células": c.celulas,
                       "Diferentes": c.celulas_diferentes, "Maior diferença": round(c.maior_diferenca, 2),
                       "Detalhe": (c.detalhe + " " + "; ".join(c.exemplos)).strip()} for c in res.comparacao]
            df_c = pd.DataFrame(linhas)
            if n_dif == 0 and all(c.status != "so_no_descritivo" for c in res.comparacao):
                st.success("Todas as tabelas calculadas batem com o Descritivo (diferença máxima de 0,05 ponto).")
            apenas = st.checkbox("Mostrar só o que não está igual", value=True, key="tc_so_dif")
            if apenas:
                df_c = df_c[df_c["Resultado"] != rot["igual"]]
            st.dataframe(df_c, use_container_width=True, hide_index=True)

    with tabs["Tabelas"]:
        itens = [(t, c) for t, c in zip(res.tabelas, res.calculadas) if c is not None]
        if not itens:
            st.info("Nenhuma tabela foi calculada.")
        else:
            busca = st.text_input("Buscar pela pergunta ou variável", key="tc_busca")
            filtrados = [(t, c) for t, c in itens
                         if not busca or busca.lower() in (t.titulo + " " + t.var_linha + " " + t.pergunta).lower()]
            if not filtrados:
                st.caption("Nada encontrado.")
            else:
                nomes = [f"#{t.ordem} · {t.secao} · {t.var_linha} × {', '.join(t.colunas)} · {t.titulo}" for t, _ in filtrados]
                todas = f"Todas as tabelas ({len(filtrados)}) — como ficam no Excel"
                opcoes = [todas] + nomes
                esc = st.selectbox("Tabela", range(len(opcoes)), format_func=lambda i: opcoes[i], key="tc_tabela_sel2")
                if esc == 0:
                    mostrar = [c for _, c in filtrados][:_MAX_TABELAS_PREVIA]
                    _previa_excel(mostrar)
                    if len(filtrados) > _MAX_TABELAS_PREVIA:
                        st.caption(f"Pré-visualização das {_MAX_TABELAS_PREVIA} primeiras tabelas; escolha uma tabela na "
                                   "lista (ou use a busca) para ver as outras. O Excel baixado tem todas.")
                else:
                    t, c = filtrados[esc - 1]
                    if t.pergunta:
                        st.caption(t.pergunta)
                    _previa_excel([c])
                    with st.expander("Ver como tabela simples"):
                        st.dataframe(_df_tabela(c), use_container_width=True)
                st.caption("Mesma aba do Excel: % da coluna com o peso (a tela mostra 1 casa; o arquivo guarda o valor inteiro) "
                           "e «Base» sem ponderar. O arquivo baixado é a fonte da verdade.")

    with tabs["Resumo"]:
        st.dataframe(pd.DataFrame([{
            "#": t.ordem, "Seção": t.secao, "Comando": t.comando, "Variável": ("$" if t.multipla else "") + t.var_linha,
            "Colunas": ", ".join(t.colunas), "Título": t.titulo,
            "Situação": "✅ calculada" if c is not None else ("❌ " + res.erros.get(t.ordem, t.motivo or "não calculada")),
        } for t, c in zip(res.tabelas, res.calculadas)]), use_container_width=True, hide_index=True)


def modulo_tratamento_banco():
    etapa = st.radio(
        "Etapa", ["1 · Recodificação do banco", "2 · Total e Completo"], horizontal=True,
        key="tb_etapa", label_visibility="collapsed",
    )
    if etapa.startswith("1"):
        _etapa_recodificacao()
    else:
        _etapa_total_completo()
