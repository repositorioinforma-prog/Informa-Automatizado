"""
Exclusão de Colunas por Base.

Apaga, de cada tabela do relatório, as colunas de segmento cuja linha 'Base'
e/ou 'Base reduzida' atende às regras escolhidas na tela — preservando a
formatação do Excel (bordas, cores, fontes, mesclagens de cabeçalho).

Diferente do Processador de Base Reduzida (que só olha 'Base reduzida' e só
aceita "menor que"), aqui a pessoa escolhe a linha de base (Base, Base
reduzida ou as duas), o operador (< ou ≤), o valor de corte e, se quiser,
restringe a regra a certos grupos de segmento (ex.: só "Região"). Dá para
somar várias regras; a coluna sai se QUALQUER uma pegar.

A tela mostra, antes de apagar, exatamente quais colunas de quais tabelas
vão sair. Depois de apagar, uma etapa final revisa o arquivo: tabela que ainda
tem base marcada com "*" mantém a nota "*Base pequena para análise
estatística"; tabela sem nenhuma base com "*" tem a linha da nota removida.
Toda a lógica mora em core/exclusao_colunas_math.py.

Não depende do `dados` carregado no início do app: tem upload próprio.
Chaves de session_state com prefixo `ec_`.
"""
import hashlib
import io
import os

import openpyxl
import pandas as pd
import streamlit as st

from core.exclusao_colunas_math import (
    TIPO_BASE,
    TIPO_BASE_REDUZIDA,
    Regra,
    aplicar_exclusoes,
    escanear_planilha,
    linhas_do_plano,
    listar_grupos,
    planejar_exclusoes,
)
from core.planilha_utils import worksheet_para_html

_OPCOES_LINHA = {
    "Base e Base reduzida": (TIPO_BASE, TIPO_BASE_REDUZIDA),
    "Somente Base": (TIPO_BASE,),
    "Somente Base reduzida": (TIPO_BASE_REDUZIDA,),
}
_OPCOES_OPERADOR = {"menor que (<)": "<", "menor ou igual a (≤)": "<="}
_PREFIXOS_DAS_REGRAS = ("ec_linhas_", "ec_op_", "ec_limite_", "ec_grupos_", "ec_rm_")
_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _limpar_estado_das_regras():
    for chave in list(st.session_state.keys()):
        if str(chave).startswith(_PREFIXOS_DAS_REGRAS):
            del st.session_state[chave]
    st.session_state["ec_ids"] = [0]
    st.session_state["ec_prox_id"] = 1


def _carregar_scans(dados):
    """Lê o arquivo uma única vez por upload (as telas seguintes só usam o scan)."""
    assinatura = hashlib.md5(dados).hexdigest()
    if st.session_state.get("ec_assinatura") != assinatura:
        try:
            with st.spinner("Lendo as tabelas do arquivo..."):
                wb = openpyxl.load_workbook(io.BytesIO(dados), rich_text=True)
                scans = {ws.title: escanear_planilha(ws) for ws in wb.worksheets}
        except Exception as e:  # arquivo corrompido, protegido por senha etc.
            st.error(f"Não foi possível ler o arquivo: {e}")
            return None
        _limpar_estado_das_regras()
        st.session_state["ec_assinatura"] = assinatura
        st.session_state["ec_scans"] = scans
        st.session_state.pop("ec_resultado", None)
    return st.session_state["ec_scans"]


def _desenhar_regras(grupos_disponiveis):
    """Uma caixa por regra (linha de base + operador + valor + grupos). Devolve a lista de Regra."""
    ids = st.session_state.setdefault("ec_ids", [0])
    st.session_state.setdefault("ec_prox_id", 1)

    regras = []
    remover = None
    for n, rid in enumerate(ids, start=1):
        with st.container(border=True):
            st.markdown(f"**Regra {n}**")
            c1, c2, c3 = st.columns([3, 3, 2])
            linhas_txt = c1.selectbox(
                "Considerar a linha", list(_OPCOES_LINHA), key=f"ec_linhas_{rid}",
                help="Qual linha da tabela tem o número que decide se a coluna sai.",
            )
            operador_txt = c2.selectbox(
                "Excluir a coluna quando o valor for", list(_OPCOES_OPERADOR), key=f"ec_op_{rid}",
            )
            limite = c3.number_input(
                "Valor", min_value=0, value=25, step=1, key=f"ec_limite_{rid}",
                help="Ex.: 'menor que 25' apaga as colunas com base 24 ou menos; 25 fica.",
            )

            chave_grupos = f"ec_grupos_{rid}"
            if chave_grupos in st.session_state:  # grupo que não existe mais (outra aba/arquivo) sai da seleção
                st.session_state[chave_grupos] = [g for g in st.session_state[chave_grupos] if g in grupos_disponiveis]
            grupos = st.multiselect(
                "Só nos grupos de segmento (vazio = todos os segmentos)", grupos_disponiveis,
                key=chave_grupos, placeholder="Todos os segmentos",
                help="O grupo é o título que fica acima das colunas no cabeçalho (ex.: Sexo, Idade, Região).",
            )
            if len(ids) > 1 and st.button("Remover esta regra", key=f"ec_rm_{rid}"):
                remover = rid

        regras.append(Regra(
            limite=float(limite), linhas=_OPCOES_LINHA[linhas_txt],
            operador=_OPCOES_OPERADOR[operador_txt], grupos=tuple(grupos),
        ))

    if remover is not None:
        ids.remove(remover)
        st.rerun()

    if st.button("➕ Adicionar outra regra", key="ec_add_regra"):
        ids.append(st.session_state["ec_prox_id"])
        st.session_state["ec_prox_id"] += 1
        st.rerun()

    return regras


def _numero_amigavel(v):
    return int(v) if float(v).is_integer() else v


def modulo_exclusao_colunas():
    st.header("Exclusão de Colunas por Base")
    st.caption(
        "Apaga, de cada tabela do relatório, as colunas de segmento cuja linha 'Base' e/ou 'Base reduzida' "
        "atende às regras abaixo — sem perder a formatação do Excel (bordas, cores, fontes e mesclagens do "
        "cabeçalho). Antes de apagar, a tela mostra exatamente o que vai sair."
    )

    arquivo = st.file_uploader("Arquivo Excel (.xlsx)", type=["xlsx"], key="ec_upload")
    if not arquivo:
        st.info("Envie um arquivo .xlsx para começar.")
        return

    dados = arquivo.getvalue()
    scans = _carregar_scans(dados)
    if scans is None:
        return

    # ------------------------------------------------------------------ o que há no arquivo
    resumo_abas = pd.DataFrame([
        {"Aba": aba, "Tabelas com base": len(s["tabelas"]),
         "Linhas 'Base'": s["n_base"], "Linhas 'Base reduzida'": s["n_base_reduzida"]}
        for aba, s in scans.items()
    ])
    with st.expander("O que foi encontrado no arquivo"):
        st.dataframe(resumo_abas, use_container_width=True, hide_index=True)

    desconhecidos = sorted({r for s in scans.values() for r in s["rotulos_base_desconhecidos"]})
    if desconhecidos:
        st.info(
            "Linhas que começam com 'Base' mas não são 'Base' nem 'Base reduzida' (ficam de fora das regras): "
            + ", ".join(f"'{r}'" for r in desconhecidos)
        )

    abas_com_tabelas = [aba for aba, s in scans.items() if s["tabelas"]]
    if not abas_com_tabelas:
        st.warning(
            "Não encontrei nenhuma linha 'Base' ou 'Base reduzida' na coluna A. "
            "Confirme se este é o relatório certo."
        )
        return

    if len(abas_com_tabelas) > 1:
        abas = st.multiselect("Abas a processar", abas_com_tabelas, default=abas_com_tabelas, key="ec_abas")
    else:
        abas = abas_com_tabelas
        st.caption(f"Aba: **{abas[0]}**")
    if not abas:
        st.info("Escolha ao menos uma aba.")
        return

    # ------------------------------------------------------------------ regras e opções
    st.subheader("Regras de exclusão")
    grupos_disponiveis = listar_grupos([scans[a] for a in abas])
    regras = _desenhar_regras(grupos_disponiveis)

    with st.expander("Opções avançadas"):
        proteger_total = st.checkbox(
            "Nunca apagar a coluna Total (coluna B)", value=True, key="ec_proteger_total",
            help="Só vale quando o cabeçalho da coluna B é 'Total' (ou está sem rótulo).",
        )
        vazias_como_zero = st.checkbox(
            "Tratar base vazia como 0", value=False, key="ec_vazias_zero",
            help="Por padrão, coluna com a célula de base vazia é mantida (não dá para saber o valor).",
        )
        revisar_notas = st.checkbox(
            "Revisar a nota '*Base pequena para análise estatística' no final", value=True, key="ec_revisar_notas",
            help="Depois de apagar as colunas, confere cada tabela: se sobrou alguma base marcada com '*', a nota "
                 "fica; se não sobrou nenhuma, a linha da nota daquela tabela é removida.",
        )
        ajustar_larguras = st.checkbox(
            "Acompanhar as larguras das colunas", value=True, key="ec_ajustar_larguras",
            help="Só é feito quando todas as tabelas da aba perdem exatamente as mesmas colunas e não há "
                 "outro conteúdo na aba; senão as larguras ficam como estão (as tabelas dividem as colunas).",
        )

    # ------------------------------------------------------------------ prévia (instantânea: não relê o Excel)
    planos = {a: planejar_exclusoes(scans[a], regras, proteger_total, vazias_como_zero) for a in abas}
    total_colunas = sum(p["total_colunas"] for p in planos.values())
    tabelas_afetadas = sum(p["tabelas_afetadas"] for p in planos.values())
    tabelas_analisadas = sum(len(scans[a]["tabelas"]) for a in abas)

    st.subheader("O que vai ser apagado")
    notas_a_remover = [
        {"Aba": a, "Tabela": i["titulo"], "Linha do título": i["linha_titulo"], "Linha da nota": i["linha"]}
        for a in abas for i in planos[a]["revisao_notas"]["remover"]
    ] if revisar_notas else []
    notas_mantidas = sum(planos[a]["revisao_notas"]["mantidas"] for a in abas) if revisar_notas else 0
    tabelas_com_asterisco_sem_nota = [
        t for a in abas for t in planos[a]["revisao_notas"]["sem_nota"]
    ] if revisar_notas else []

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Colunas a excluir", total_colunas)
    m2.metric("Tabelas afetadas", tabelas_afetadas)
    m3.metric("Tabelas analisadas", tabelas_analisadas)
    m4.metric("Notas de base pequena a remover", len(notas_a_remover))

    formulas_na_base = sum(b["formulas"] for a in abas for t in scans[a]["tabelas"] for b in t["bases"])
    if formulas_na_base:
        st.warning(
            f"{formulas_na_base} célula(s) de base são fórmulas — o valor delas não pôde ser lido e essas "
            "colunas ficam de fora das regras."
        )
    sem_segmentos = [t for p in planos.values() for t in p["tabelas_sem_segmentos"]]
    if sem_segmentos:
        exemplos = "; ".join(f"{t['aba']} (linha {t['linha_titulo'] or t['linha_ini']}): {t['titulo']}"
                             for t in sem_segmentos[:5])
        st.warning(
            f"{len(sem_segmentos)} tabela(s) ficariam só com a coluna Total, sem nenhum segmento: {exemplos}"
            + (" ..." if len(sem_segmentos) > 5 else "")
        )

    if tabelas_com_asterisco_sem_nota:
        exemplos = "; ".join(t["titulo"] for t in tabelas_com_asterisco_sem_nota[:5])
        st.info(
            f"{len(tabelas_com_asterisco_sem_nota)} tabela(s) têm base marcada com '*' mas não têm a nota "
            f"'*Base pequena para análise estatística' (nada é inserido): {exemplos}"
            + (" ..." if len(tabelas_com_asterisco_sem_nota) > 5 else "")
        )

    if total_colunas == 0 and not notas_a_remover:
        st.success("Nenhuma coluna atende às regras e nenhuma nota precisa sair — o arquivo ficaria igual ao original.")
        return

    if total_colunas:
        linhas = [l for a in abas for l in linhas_do_plano(planos[a])]
        st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

    if notas_a_remover:
        st.markdown(
            f"**Notas '*Base pequena para análise estatística' que saem** — depois das exclusões, essas tabelas "
            f"ficam sem nenhuma base marcada com '*' ({notas_mantidas} nota(s) continuam, pois ainda há base com '*'):"
        )
        st.dataframe(pd.DataFrame(notas_a_remover), use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------ aplicar
    assinatura_plano = repr((tuple(abas), tuple(regras), proteger_total, vazias_como_zero, ajustar_larguras,
                             revisar_notas))

    if st.button("Aplicar exclusões e gerar arquivo", key="ec_aplicar"):
        with st.spinner("Apagando as colunas..."):
            wb = openpyxl.load_workbook(io.BytesIO(dados), rich_text=True)
            resumos = {a: aplicar_exclusoes(wb[a], planos[a], ajustar_larguras, revisar_notas) for a in abas}
            saida = io.BytesIO()
            wb.save(saida)
        st.session_state["ec_resultado"] = {
            "bytes": saida.getvalue(),
            "nome": os.path.splitext(arquivo.name)[0] + " - colunas excluidas.xlsx",
            "resumos": resumos,
            "assinatura": assinatura_plano,
            "abas": list(abas),
        }

    resultado = st.session_state.get("ec_resultado")
    if not resultado:
        return

    st.markdown("---")
    st.subheader("Resultado")
    if resultado["assinatura"] != assinatura_plano:
        st.warning("As regras ou opções mudaram depois da última geração — clique em "
                   "'Aplicar exclusões e gerar arquivo' de novo para atualizar o arquivo abaixo.")

    n_excl = sum(r["colunas_excluidas"] for r in resultado["resumos"].values())
    n_tab = sum(r["tabelas_afetadas"] for r in resultado["resumos"].values())
    n_notas = sum(r["notas_removidas"] for r in resultado["resumos"].values())
    n_notas_ficaram = sum(r["notas_mantidas"] for r in resultado["resumos"].values())
    texto = f"Pronto! {n_excl} coluna(s) excluída(s) em {n_tab} tabela(s)."
    if n_notas or n_notas_ficaram:
        texto += (f" Notas '*Base pequena para análise estatística': {n_notas} removida(s), "
                  f"{n_notas_ficaram} mantida(s) (tabelas que ainda têm base com '*').")
    st.success(texto)
    for aba, r in resultado["resumos"].items():
        if r["larguras_ajustadas"]:
            st.caption(f"Aba '{aba}': larguras das colunas acompanharam o deslocamento.")
        if r["tabelas_com_asterisco_sem_nota"]:
            st.info(f"Aba '{aba}': tabela(s) com base '*' e sem a nota (nada foi inserido): "
                    + "; ".join(r["tabelas_com_asterisco_sem_nota"][:5]))
        for aviso in r["avisos"]:
            st.warning(f"Aba '{aba}': {aviso}")

    st.download_button(
        "Baixar arquivo com as colunas excluídas", data=resultado["bytes"], file_name=resultado["nome"],
        mime=_MIME_XLSX, key="ec_download",
    )

    if st.checkbox("Mostrar pré-visualização rápida do resultado", key="ec_mostrar_preview"):
        wb_preview = openpyxl.load_workbook(io.BytesIO(resultado["bytes"]))
        aba_prev = (
            st.selectbox("Aba", resultado["abas"], key="ec_preview_aba")
            if len(resultado["abas"]) > 1 else resultado["abas"][0]
        )
        html, truncou_linhas, truncou_colunas = worksheet_para_html(wb_preview[aba_prev])
        st.components.v1.html(html, height=620, scrolling=True)
        st.caption(
            "Pré-visualização aproximada (reproduz mesclagem, cores, fonte, alinhamento e bordas); "
            "o arquivo baixado é a fonte da verdade."
            + (" Mostrando só o começo da aba." if truncou_linhas or truncou_colunas else "")
        )
