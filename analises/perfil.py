"""
Perfil.

Sobe o banco (.sav) e o app gera as tabelas do Perfil — Sexo (P1), Idade (P2), Renda Familiar (P3),
Escolaridade (P4) e Religiões (juncaoreligiao) — com a variável de peso ligada (padrão: «peso»).
Em projeto presencial (rua) a Idade sai também com a «Média» (variável MEP2); em projeto online, não.

O resultado baixa em Excel no formato do «Perfil_<Cidade>_<UF>_<MêsAno>», o mesmo que o preenchimento do
slide de perfil já lê. Não depende do `dados` carregado no início do app: tem upload próprio.
Toda a lógica mora em core/perfil_tabelas_math.py. Chaves de session_state com prefixo `pf_`.
"""
import hashlib

import streamlit as st

from core.analise_perguntas_math import coluna_do_banco, html_sobre_fundo_branco
from core.analitica_dados import carregar_banco, eh_nome_de_peso
from core.perfil_tabelas_math import (
    MODO_PRESENCIAL,
    PESO_PADRAO,
    VAR_MEDIA_IDADE,
    VARIAVEIS_PERFIL,
    calcular_perfil,
    montar_excel_perfil,
    nome_arquivo_perfil,
    planilha_perfil,
    sugerir_identificacao,
)
from core.planilha_utils import worksheet_para_html

_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_SEM_PESO = "Sem peso"
_PRESENCIAL = "Presencial (rua)"
_ONLINE = "Online"


def _hash(conteudo: bytes) -> str:
    return hashlib.md5(conteudo).hexdigest()


def _carregar_banco(arq_sav):
    """Lê o banco uma vez por conteúdo (guarda em session_state)."""
    conteudo = arq_sav.getvalue()
    chave = _hash(conteudo)
    if st.session_state.get("pf_chave") == chave:
        return st.session_state["pf_banco"], chave
    banco = carregar_banco(conteudo, arq_sav.name)
    st.session_state["pf_chave"] = chave
    st.session_state["pf_banco"] = banco
    for k in ("pf_resultado", "pf_excel", "pf_snap"):
        st.session_state.pop(k, None)
    return banco, chave


def modulo_perfil():
    st.caption(
        "Gera as tabelas do Perfil (Sexo, Idade, Renda Familiar, Escolaridade e Religiões) direto do banco, "
        "com o peso ligado, e baixa em Excel no formato do Perfil do SPSS."
    )

    arq_sav = st.file_uploader("Banco de dados (.sav)", type=["sav"], key="pf_up_sav")
    if not arq_sav:
        st.info("Suba o banco (.sav) para começar.")
        return
    try:
        banco, chave = _carregar_banco(arq_sav)
    except Exception as e:  # leitura de arquivo: mostra o motivo em vez de quebrar a tela
        st.error(f"Não consegui ler o banco: {e}")
        return

    df = banco.df
    colunas = [str(c) for c in df.columns]
    c1, c2 = st.columns(2)
    c1.metric("Casos no banco", f"{len(df):,}".replace(",", "."))
    achadas = [t for _, v, t in VARIAVEIS_PERFIL if coluna_do_banco(v, colunas)]
    c2.metric("Variáveis do Perfil no banco", f"{len(achadas)} de {len(VARIAVEIS_PERFIL)}")
    faltam = [f"{v} ({t})" for _, v, t in VARIAVEIS_PERFIL if not coluna_do_banco(v, colunas)]
    if faltam:
        st.warning("Não estão no banco e ficam de fora: " + ", ".join(faltam) + ".")

    # ---------------- 1. tipo de projeto
    st.subheader("1. Tipo de projeto")
    tem_media = bool(coluna_do_banco(VAR_MEDIA_IDADE, colunas))
    st.caption(
        "Presencial (rua) leva a «Média» de idade na tabela de Idade (variável "
        f"{VAR_MEDIA_IDADE}); online não leva."
        + ("" if tem_media else f" Este banco não tem a variável {VAR_MEDIA_IDADE}.")
    )
    tipo = st.radio(
        "Tipo de projeto",
        [_PRESENCIAL, _ONLINE],
        index=0 if tem_media else 1,
        horizontal=True,
        key=f"pf_tipo_{chave}",
        label_visibility="collapsed",
    )
    presencial = tipo == _PRESENCIAL

    # ---------------- 2. peso
    st.subheader("2. Peso")
    pesos_cand = [c for c in colunas if eh_nome_de_peso(c)]
    opcoes_peso = [_SEM_PESO] + pesos_cand
    idx = opcoes_peso.index(PESO_PADRAO) if PESO_PADRAO in opcoes_peso else 0
    if PESO_PADRAO not in opcoes_peso:
        st.warning(f"A variável «{PESO_PADRAO}» não está no banco: escolha outra ou siga sem peso.")
    peso_sel = st.selectbox("Variável de peso", opcoes_peso, index=idx, key=f"pf_peso_{chave}")
    peso = None if peso_sel == _SEM_PESO else peso_sel

    # ---------------- 3. identificação do arquivo
    st.subheader("3. Nome do arquivo")
    uf_sug, mes_sug = sugerir_identificacao(arq_sav.name)
    n1, n2, n3 = st.columns([3, 1, 1])
    cidade = n1.text_input("Cidade", value="", key=f"pf_cidade_{chave}", placeholder="Ex.: Paty do Alferes")
    uf = n2.text_input("UF", value=uf_sug, max_chars=2, key=f"pf_uf_{chave}", placeholder="RJ")
    mes_ano = n3.text_input("Mês/Ano", value=mes_sug, key=f"pf_mes_{chave}", placeholder="Set26")
    st.caption("Arquivo: " + nome_arquivo_perfil(cidade, uf, mes_ano))

    # ---------------- 4. gerar
    st.subheader("4. Gerar")
    if st.button("Gerar Perfil", type="primary", key="pf_rodar"):
        try:
            with st.spinner("Calculando as tabelas..."):
                res = calcular_perfil(df, banco.value_labels, peso, presencial)
                ident = {"Cidade": cidade.strip(), "UF": uf.strip().upper(), "Mês/Ano": mes_ano.strip()}
                xlsx = montar_excel_perfil(res, arq_sav.name, ident)
        except Exception as e:
            st.error(f"Não consegui gerar o Perfil: {e}")
            return
        st.session_state["pf_resultado"] = res
        st.session_state["pf_excel"] = xlsx
        st.session_state["pf_snap"] = {"arquivo": nome_arquivo_perfil(cidade, uf, mes_ano)}

    res = st.session_state.get("pf_resultado")
    if res is None:
        return
    if not res.blocos:
        st.error("Nenhuma tabela foi gerada: o banco não tem nenhuma variável do Perfil com casos válidos.")
        for a in res.avisos:
            st.warning(a)
        return
    st.success(f"{len(res.blocos)} tabela(s) gerada(s) — " + ("presencial (rua)" if res.modo == MODO_PRESENCIAL else "online") + ".")
    for a in res.avisos:
        st.warning(a)
    st.download_button(
        "Baixar Excel",
        data=st.session_state["pf_excel"],
        file_name=st.session_state["pf_snap"]["arquivo"],
        mime=_MIME_XLSX,
        key="pf_download",
    )

    # ---------------- pré-visualização (a mesma aba que vai para o Excel)
    st.subheader("Pré-visualização")
    ws = planilha_perfil(res, como_texto=True)
    html, _trunc_l, _trunc_c = worksheet_para_html(ws, max_linhas=400)
    altura = int(min(760, max(240, ws.max_row * 21 + 50)))
    st.components.v1.html(html_sobre_fundo_branco(html), height=altura, scrolling=True)
    st.caption(
        "Mesma aba «Perfil» do Excel: % do total ponderado (1 casa), «Base» sem ponderar e, no presencial, "
        "a «Média» de idade. O arquivo baixado é a fonte da verdade."
    )
