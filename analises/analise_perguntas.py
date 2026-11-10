"""
Análise das Perguntas.

Sobe a programação do SPSS (Total, ou Total + Completo — só o Total é lido) e o banco
(.sav). O app lista as variáveis que viram tabela no Total, confere se cada uma tem label
no banco (se não tiver, usa o título da tabela do Total) e cruza cada pergunta pelas
variáveis de cruzamento: as pré-definidas (P1 P2 P3 P4 P5 juncaoreligiao P7 P7_C VN, só as
que existem no banco) e as outras que a pessoa quiser marcar. O resultado sai calculado
aqui dentro e baixa em Excel no layout da exportação do SPSS já formatada (uma aba com todas as
tabelas empilhadas, % da coluna + resíduo ajustado, já pintado como no VBA de destaque) + aba «Resumo».

Não depende do `dados` carregado no início do app: tem upload próprio.
Toda a lógica mora em core/analise_perguntas_math.py. Chaves de session_state com prefixo `ap_`.
"""
import hashlib

import pandas as pd
import streamlit as st

from core.analise_perguntas_math import (
    CRUZAMENTOS_PADRAO,
    FONTE_BANCO,
    FONTE_SEM_LABEL,
    FONTE_TITULO,
    LIMIAR_Z_PADRAO,
    ROTULO_FONTE,
    calcular_analise,
    coluna_do_banco,
    cruzamentos_padrao_presentes,
    ler_perguntas_total,
    bloco_como_planilha,
    montar_excel,
    resolver_labels,
    separar_total,
)
from core.analitica_dados import carregar_banco, eh_nome_de_peso
from core.planilha_utils import worksheet_para_html
from core.revisor_syntax_math import decode_sps

_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_SEM_PESO = "Sem peso"


def _hash(conteudo: bytes) -> str:
    return hashlib.md5(conteudo).hexdigest()


def _limpar_resultado():
    for k in ("ap_resultado", "ap_excel"):
        st.session_state.pop(k, None)


def _carregar_arquivos(arq_sps, arq_sav):
    """Lê os dois arquivos uma vez por conteúdo (guarda em session_state)."""
    b_sps, b_sav = arq_sps.getvalue(), arq_sav.getvalue()
    chave = (_hash(b_sps), _hash(b_sav))
    if st.session_state.get("ap_chave") == chave:
        return st.session_state["ap_dados"]
    texto, _codificacao, _achados = decode_sps(b_sps)
    _, tem_completo = separar_total(texto.replace("\r\n", "\n"))
    tabelas = ler_perguntas_total(texto)
    banco = carregar_banco(b_sav, arq_sav.name)
    dados = {"tabelas": tabelas, "banco": banco, "tem_completo": tem_completo}
    st.session_state["ap_chave"] = chave
    st.session_state["ap_dados"] = dados
    for k in list(st.session_state.keys()):
        if str(k).startswith(("ap_lab_", "ap_perg", "ap_cruz", "ap_extra", "ap_peso")):
            del st.session_state[k]
    _limpar_resultado()
    return dados


def modulo_analise_perguntas():
    st.caption(
        "Cruza cada pergunta do Total pelas variáveis de cruzamento e baixa tudo em Excel, no layout "
        "da exportação do SPSS (% da coluna e resíduo ajustado, já pintado). A conta segue o CROSSTABS "
        "do SPSS com /COUNT ROUND CELL."
    )

    c1, c2 = st.columns(2)
    arq_sps = c1.file_uploader("Programação do SPSS (Total ou Total + Completo)", type=["sps"], key="ap_up_sps")
    arq_sav = c2.file_uploader("Banco de dados (.sav)", type=["sav"], key="ap_up_sav")
    if not arq_sps or not arq_sav:
        st.info("Suba a programação (.sps) e o banco (.sav) para começar.")
        return

    try:
        dados = _carregar_arquivos(arq_sps, arq_sav)
    except Exception as e:  # leitura de arquivo: mostra o motivo em vez de quebrar a tela
        st.error(f"Não consegui ler os arquivos: {e}")
        return

    tabelas, banco = dados["tabelas"], dados["banco"]
    df = banco.df
    colunas = [str(c) for c in df.columns]
    if not tabelas:
        st.error("Não encontrei nenhuma tabela (TABLES/CTABLES) no Total dessa programação.")
        return
    if dados["tem_completo"]:
        st.caption("A programação tem Total e Completo: só o Total foi lido.")

    unicas = [t for t in tabelas if not t.multipla and coluna_do_banco(t.var, colunas)]
    multiplas = [t.var for t in tabelas if t.multipla]
    fora_banco = [t.var for t in tabelas if not t.multipla and not coluna_do_banco(t.var, colunas)]
    m1, m2, m3 = st.columns(3)
    m1.metric("Variáveis no Total", len(tabelas))
    m2.metric("Cruzáveis (no banco)", len(unicas))
    m3.metric("Casos no banco", f"{len(df):,}".replace(",", "."))
    if multiplas:
        st.warning(
            "Múltipla resposta não é cruzada nesta versão (não existe como coluna no banco): "
            + ", ".join(multiplas) + "."
        )
    if fora_banco:
        st.warning("Variáveis do Total que não estão no banco: " + ", ".join(fora_banco) + ".")
    if not unicas:
        return

    # ---------------- 1. labels
    st.subheader("1. Labels das perguntas")
    preferir = st.checkbox(
        "Usar sempre o título do Total quando ele existir",
        value=False,
        key="ap_pref_titulo",
        help="Padrão: vale o label do banco; só se o banco não tiver label entra o título do Total.",
    )
    labels = {lv.var: lv for lv in resolver_labels(unicas, banco.var_labels, preferir_titulo=preferir)}
    n_banco = sum(1 for lv in labels.values() if lv.fonte == FONTE_BANCO)
    n_titulo = sum(1 for lv in labels.values() if lv.fonte == FONTE_TITULO)
    n_sem = sum(1 for lv in labels.values() if lv.fonte == FONTE_SEM_LABEL)
    n_dif = sum(1 for lv in labels.values() if lv.difere)
    st.caption(
        f"{n_banco} com label do banco · {n_titulo} usando o título do Total · {n_sem} sem nenhum texto"
        + (f" · {n_dif} com texto diferente entre banco e Total" if n_dif else "")
    )
    base_ed = pd.DataFrame(
        [
            {
                "Variável": lv.var,
                "Label usada": lv.label_usada,
                "Fonte": ROTULO_FONTE[lv.fonte],
                "Label no banco": lv.label_banco,
                "Título no Total": lv.titulo_total,
            }
            for lv in labels.values()
        ]
    )
    # a chave inclui a escolha acima: ao mudá-la, a tabela volta aos valores novos
    editado = st.data_editor(
        base_ed,
        key=f"ap_lab_{int(preferir)}",
        hide_index=True,
        use_container_width=True,
        disabled=["Variável", "Fonte", "Label no banco", "Título no Total"],
        column_config={"Label usada": st.column_config.TextColumn("Label usada (editável)", width="large")},
    )
    label_final = dict(zip(editado["Variável"], editado["Label usada"]))
    for lv in labels.values():
        txt = " ".join(str(label_final.get(lv.var) or "").split())
        lv.label_usada = txt or lv.var

    # ---------------- 2. perguntas e cruzamentos
    st.subheader("2. O que cruzar")
    todas = [t.var for t in unicas]
    perguntas = st.multiselect("Perguntas (linhas)", todas, default=todas, key="ap_perguntas")

    presentes, ausentes = cruzamentos_padrao_presentes(colunas)
    st.caption("Cruzamentos pré-definidos: " + ", ".join(CRUZAMENTOS_PADRAO))
    if ausentes:
        st.info("Não estão no banco e ficam de fora: " + ", ".join(ausentes) + ".")
    cruz_padrao = st.multiselect("Cruzamentos pré-definidos (colunas)", presentes, default=presentes, key="ap_cruz")

    padrao_min = {p.lower() for p in CRUZAMENTOS_PADRAO}
    pesos_cand = [c for c in colunas if eh_nome_de_peso(c)]
    candidatos_extra = [c for c in colunas if c.lower() not in padrao_min and c not in pesos_cand and c not in todas]
    extras = st.multiselect(
        "Outros cruzamentos (opcional)",
        candidatos_extra,
        default=[],
        key="ap_extra",
        format_func=lambda c: f"{c} — {banco.var_labels.get(c) or ''}".rstrip(" —"),
    )
    cruzamentos = list(dict.fromkeys(cruz_padrao + extras))

    # ---------------- 3. peso e limiar
    st.subheader("3. Peso e destaque do resíduo")
    st.caption(
        "Nem o banco nem o Total trazem um peso ativo. O padrão é a variável «peso», que reproduziu a saída do "
        "SPSS do RJ Deputados; escolha «Sem peso» se o seu SPSS foi rodado sem WEIGHT BY."
    )
    opcoes_peso = [_SEM_PESO] + pesos_cand
    idx = opcoes_peso.index("peso") if "peso" in opcoes_peso else 0
    peso_sel = st.selectbox("Variável de peso", opcoes_peso, index=idx, key="ap_peso")
    peso = None if peso_sel == _SEM_PESO else peso_sel
    limiar = st.number_input(
        "Pintar o resíduo ajustado a partir de ±",
        min_value=0.5, max_value=5.0, value=float(LIMIAR_Z_PADRAO), step=0.05, key="ap_limiar",
        help="Mesmo limite do VBA de destaque e da Análise Residual (A.R): 1,85. Vermelho (#C00000) para negativo e azul (#002060) para positivo.",
    )

    incluir_contagem = st.checkbox(
        "Incluir a linha de contagem (Count) no Excel",
        value=True,
        key="ap_contagem",
        help="Como no CROSSTABS do SPSS: Count, % within e Adjusted Residual por alternativa. "
             "Desmarque para sair só com % e resíduo, como no arquivo formatado de Teresópolis.",
    )

    # ---------------- 4. rodar
    st.subheader("4. Calcular")
    if not perguntas or not cruzamentos:
        st.info("Marque ao menos uma pergunta e um cruzamento.")
        return
    st.caption(f"{len(perguntas)} pergunta(s) × {len(cruzamentos)} cruzamento(s).")
    if st.button("Calcular análise", type="primary", key="ap_rodar"):
        with st.spinner("Calculando os cruzamentos..."):
            res = calcular_analise(df, perguntas, cruzamentos, peso, banco.value_labels)
            rot_cruz = {c: (banco.var_labels.get(c) or c) for c in cruzamentos}
            avisos = [f"Variáveis do Total fora do banco (ignoradas): {', '.join(fora_banco)}."] if fora_banco else []
            if multiplas:
                avisos.append(f"Múltipla resposta não cruzada: {', '.join(multiplas)}.")
            if ausentes:
                avisos.append(f"Cruzamentos pré-definidos ausentes no banco: {', '.join(ausentes)}.")
            xlsx = montar_excel(
                res, [labels[p] for p in perguntas], rot_cruz,
                peso=peso, limiar_z=float(limiar), perguntas_total=tabelas,
                nome_banco=arq_sav.name, nome_syntax=arq_sps.name, avisos=avisos,
                incluir_contagem=incluir_contagem,
            )
        st.session_state["ap_resultado"] = res
        st.session_state["ap_excel"] = xlsx
        st.session_state["ap_snap"] = {   # o que foi usado no cálculo: a pré-visualização mostra o mesmo
            "labels": {p: labels[p].label_usada for p in perguntas},
            "rot_cruz": rot_cruz,
            "limiar": float(limiar),
            "contagem": bool(incluir_contagem),
        }

    res = st.session_state.get("ap_resultado")
    if res is None:
        return
    st.success(f"{res.n_tabelas} tabela(s) calculada(s).")
    if res.sem_dados:
        st.warning("Sem nenhum caso válido (não geradas): " + ", ".join(f"{p} × {c}" for p, c in res.sem_dados[:20]))
    st.download_button(
        "Baixar Excel",
        data=st.session_state["ap_excel"],
        file_name="Analise_das_Perguntas.xlsx",
        mime=_MIME_XLSX,
        key="ap_download",
    )

    # ---------------- pré-visualização (o mesmo bloco que vai para o Excel)
    st.subheader("Pré-visualização")
    com_tabela = [p for p, lista in res.por_pergunta.items() if lista]
    if not com_tabela:
        return
    snap = st.session_state["ap_snap"]
    p_sel = st.selectbox("Pergunta", com_tabela, key="ap_prev_p")
    lista = res.por_pergunta[p_sel]
    c_sel = st.selectbox("Cruzamento", [cz.coluna for cz in lista], key="ap_prev_c")
    cz = next(x for x in lista if x.coluna == c_sel)
    ws = bloco_como_planilha(
        cz, snap["labels"].get(p_sel, p_sel), snap["rot_cruz"].get(c_sel, c_sel),
        snap["limiar"], snap["contagem"], como_texto=True,
    )
    html, _trunc_l, _trunc_c = worksheet_para_html(ws, max_linhas=800)
    altura = int(min(760, max(240, ws.max_row * 21 + 50)))
    st.components.v1.html(html, height=altura, scrolling=True)
    st.caption(
        "Mesma tabela que vai para o Excel (mesclagens, cores do destaque e formatos). "
        "O arquivo baixado é a fonte da verdade."
    )
