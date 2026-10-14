"""
Controle Geral — o app "Controle-Automatico" (Gerador de Controle Geral) dentro
do Informa.

O motor fica no pacote controle/ (ver controle/__init__.py). Esta tela foi
reorganizada em DOIS FLUXOS (v32), com a mesma lógica do app original:

- Recodificação geográfica, em etapas: Arquivos → Variáveis → Processar →
  Revisar → Exportar (trilha clicável no topo e Voltar/Avançar no fim);
- Checagens: Duplicidades, Validação eleitoral e Exclusões consolidadas, com
  um resumo de cada uma no topo.

Visual: faixa de status do projeto (banco, syntax, Base Brasil, estados, ID),
cartões coloridos por status, tabelas enxutas (colunas essenciais, com "Mostrar
todas as colunas"), status coloridos nas tabelas e downloads agrupados em
cartões com a descrição de cada arquivo. As escolhas da tela (etapa, estados,
variáveis, filtros, cortes) são guardadas ao trocar de etapa ou de ferramenta.

Mudanças em relação ao app original: estado da sessão com prefixo "ctl_",
Base Brasil em controle/data/, .sav pelo pyreadstat seguro; o banco e a syntax
carregam separados (as checagens só precisam do banco) e trocar de banco limpa
os resultados antigos. Correções da v31 (marcadas com "[Informa]" em
controle/): realocação sem inconsistências, hierarquia de Campo Grande/MS só
com MS na abrangência, telefones de preenchimento, duplicidade "possível",
.xlsx em cache e .sav gravado sem pyreadstat.
"""
from __future__ import annotations

import html
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from controle import PASTA_DADOS

from controle.core.duplicates import (
    DATE_LABEL_PATTERNS, DuplicateConfig, NAME_LABEL_PATTERNS, PHONE_LABEL_PATTERNS,
    find_column_by_label, find_duplicates,
)
from controle.core.geographic_validation import build_geographic_coherence_report, build_cidade_bairro_realocation
from controle.core.history_learning import build_text_to_code_lookup, apply_learned_lookup
from controle.core.electoral_validation import detect_nao_vota_codes, validate_electoral
from controle.core.recoding import ControlVariableConfig, recode_dataframe
from controle.core.value_labels import parse_all_value_labels_sps, decode_sps_bytes
from controle.exporters.sps import make_exclusion_syntax, make_exclusion_syntax_with_reasons, make_value_labels_syntax
from controle.exporters.sav import write_sav_bytes
from controle.geography.database import load_geography_excel, UF_OPTIONS, uf_option_to_sigla
from controle.geography.hierarquia_bairros import get_hierarchy
from controle.importers.sav import read_sav_bytes, variable_catalog


# --- Caches: sem isso, o Streamlit reprocessa banco, syntax e Base Brasil do
# zero a cada interação.
@st.cache_data(show_spinner="Lendo banco .sav...")
def _cached_read_sav(data: bytes):
    return read_sav_bytes(data)


@st.cache_data(show_spinner="Lendo VALUE LABELS...")
def _cached_parse_sps(text: str):
    return parse_all_value_labels_sps(text)


@st.cache_resource(show_spinner="Carregando Base Brasil interna...")
def _cached_geography_from_path(path_str: str):
    path = Path(path_str)
    return load_geography_excel(path, source_name=path.name)


@st.cache_data(show_spinner="Carregando Base Brasil...")
def _cached_geography_from_bytes(data: bytes, name: str):
    return load_geography_excel(data, source_name=name)


def _load_internal_geography():
    candidates = [
        PASTA_DADOS / "Cidades_Bairros_Distritos - Brasil.xlsx",
        PASTA_DADOS / "Cidades_Bairros_Distritos - Brasil(1).xlsx",
    ]
    for path in candidates:
        if path.exists():
            return _cached_geography_from_path(str(path))
    return None


class _EstadoControle:
    """st.session_state com o prefixo "ctl_" (ss.x, ss.get("x"), ss.pop("x"), "x" in ss)."""
    _P = "ctl_"

    def __getattr__(self, nome):
        try:
            return st.session_state[self._P + nome]
        except KeyError:
            raise AttributeError(nome) from None

    def __setattr__(self, nome, valor):
        st.session_state[self._P + nome] = valor

    def __contains__(self, nome):
        return (self._P + nome) in st.session_state

    def get(self, nome, padrao=None):
        return st.session_state.get(self._P + nome, padrao)

    def pop(self, nome, *padrao):
        return st.session_state.pop(self._P + nome, *padrao)


ss = _EstadoControle()

# resultados que dependem do banco/da configuração
_DERIVADOS = ["audit_df", "output_df", "generated_label_sets", "coherence_df", "realocacao_df", "sav_bytes",
              "xlsx_cache"]
_DERIVADOS_CHECAGENS = ["duplicates_df", "duplicates_id_column", "eleitoral_df"]

# =====================================================================
# Visual
# =====================================================================
#: status da auditoria -> (fundo, texto, borda)
CORES_STATUS = {
    "CONFIRMADO": ("#dcfce7", "#166534", "#86efac"),
    "AJUSTADO AUTOMATICAMENTE": ("#e0f2fe", "#075985", "#7dd3fc"),
    "APRENDIDO (banco anterior)": ("#e0e7ff", "#3730a3", "#a5b4fc"),
    "ACEITO EM LOTE": ("#ede9fe", "#5b21b6", "#c4b5fd"),
    "REALOCADO (Cidade × Bairro)": ("#f3e8ff", "#6b21a8", "#d8b4fe"),
    "SUGESTÃO": ("#fef3c7", "#92400e", "#fcd34d"),
    "AMBÍGUO": ("#ffedd5", "#9a3412", "#fdba74"),
    "FORA DA AMOSTRA": ("#fee2e2", "#991b1b", "#fca5a5"),
    "NÃO IDENTIFICADO": ("#fce7f3", "#9d174d", "#f9a8d4"),
    "VAZIO": ("#f3f4f6", "#4b5563", "#d1d5db"),
    # checagens
    "VOTA NO ESTADO DA PESQUISA": ("#dcfce7", "#166534", "#86efac"),
    "VOTA EM OUTRO ESTADO": ("#fee2e2", "#991b1b", "#fca5a5"),
    "NÃO VOTA": ("#fef3c7", "#92400e", "#fcd34d"),
    "INCONSISTENTE": ("#fee2e2", "#991b1b", "#fca5a5"),
    "certa": ("#fee2e2", "#991b1b", "#fca5a5"),
    "provavel": ("#ffedd5", "#9a3412", "#fdba74"),
    "possivel": ("#fef3c7", "#92400e", "#fcd34d"),
    "manter": ("#dcfce7", "#166534", "#86efac"),
    "excluir": ("#fee2e2", "#991b1b", "#fca5a5"),
}
_NEUTRO = ("#f3f4f6", "#374151", "#d1d5db")

ORDEM_STATUS = ["CONFIRMADO", "AJUSTADO AUTOMATICAMENTE", "APRENDIDO (banco anterior)", "ACEITO EM LOTE",
                "REALOCADO (Cidade × Bairro)", "SUGESTÃO", "AMBÍGUO", "FORA DA AMOSTRA", "NÃO IDENTIFICADO",
                "VAZIO"]
PENDENTES = {"SUGESTÃO", "AMBÍGUO", "FORA DA AMOSTRA", "NÃO IDENTIFICADO"}
NOMES_STATUS = {
    "CONFIRMADO": "Confirmado", "AJUSTADO AUTOMATICAMENTE": "Ajustado automaticamente",
    "APRENDIDO (banco anterior)": "Aprendido (banco anterior)", "ACEITO EM LOTE": "Aceito em lote",
    "REALOCADO (Cidade × Bairro)": "Realocado (Cidade × Bairro)", "SUGESTÃO": "Sugestão",
    "AMBÍGUO": "Ambíguo", "FORA DA AMOSTRA": "Fora da amostra", "NÃO IDENTIFICADO": "Não identificado",
    "VAZIO": "Vazio", "INCONSISTENTE": "Realocação sugerida", "excluir": "A excluir",
    "VOTA NO ESTADO DA PESQUISA": "Vota no estado", "VOTA EM OUTRO ESTADO": "Vota em outro estado",
    "NÃO VOTA": "Não vota",
}

CSS = """
<style>
.ctl-faixa{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 6px}
.ctl-chip{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:999px;font-size:13px;
  line-height:1.3;border:1px solid #e5e7eb;background:#f9fafb;color:#374151}
.ctl-chip b{font-weight:700}
.ctl-chip.ok{background:#ecfdf3;border-color:#bbf7d0;color:#166534}
.ctl-chip.pend{background:#fff7ed;border-color:#fed7aa;color:#9a3412}
.ctl-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin:4px 0 14px}
.ctl-card{border-radius:12px;padding:10px 12px;border:1px solid}
.ctl-card .n{font-size:26px;font-weight:800;line-height:1.15}
.ctl-card .t{font-size:12.5px;font-weight:700}
.ctl-card .p{font-size:11.5px;opacity:.8}
.ctl-resumo{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin:4px 0 14px}
.ctl-res{border:1px solid #e5e7eb;border-radius:14px;padding:12px 14px;background:#fff}
.ctl-res .h{font-size:12px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:#6b7280}
.ctl-res .n{font-size:28px;font-weight:800;line-height:1.2;color:#111827}
.ctl-res .d{font-size:12.5px;color:#4b5563}
.ctl-res.vazio .n{font-size:15px;font-weight:600;color:#9ca3af;padding:6px 0 3px}
.ctl-res.alerta{border-color:#fca5a5;background:#fef2f2}
.ctl-res.ok{border-color:#86efac;background:#f0fdf4}
.ctl-etapa{font-size:13px;color:#6b7280;margin:-4px 0 10px}
.ctl-dl-t{font-weight:700;font-size:15px}
.ctl-dl-d{font-size:12.5px;color:#6b7280;min-height:36px}
.ctl-var{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.ctl-tag{display:inline-block;padding:2px 9px;border-radius:999px;font-size:12px;font-weight:700;
  background:#eef2ff;color:#3730a3;border:1px solid #c7d2fe}
.ctl-tag.cinza{background:#f3f4f6;color:#374151;border-color:#e5e7eb}

/* O tema do Streamlit é escuro e o app.py força texto escuro: os componentes
   que o app.py não pinta (seletor de etapas e cabeçalho dos expanders) ficavam
   texto escuro sobre fundo escuro. Aqui eles ganham fundo claro. */
.stApp [data-testid="stButtonGroup"] button,
.stApp button[kind^="segmented_control"],.stApp button[data-testid^="stBaseButton-segmented_control"]{
  background:#ffffff !important;border:1.5px solid #cbd5e1 !important;color:#1f2937 !important;
  font-weight:600 !important;box-shadow:none !important}
.stApp [data-testid="stButtonGroup"] button *,
.stApp button[kind^="segmented_control"] *{color:#1f2937 !important}
.stApp [data-testid="stButtonGroup"] button:hover{background:#eff6ff !important;border-color:#2563eb !important}
.stApp button[kind="segmented_controlActive"],.stApp button[data-testid="stBaseButton-segmented_controlActive"],
.stApp [data-testid="stButtonGroup"] button[data-testid$="Active"],
.stApp [data-testid="stButtonGroup"] button[kind$="Active"],
.stApp [data-testid="stButtonGroup"] button[aria-checked="true"],
.stApp [data-testid="stButtonGroup"] button[aria-pressed="true"]{
  background:#2563eb !important;border-color:#2563eb !important;color:#ffffff !important}
.stApp button[kind="segmented_controlActive"] *,.stApp button[data-testid="stBaseButton-segmented_controlActive"] *,
.stApp [data-testid="stButtonGroup"] button[data-testid$="Active"] *,
.stApp [data-testid="stButtonGroup"] button[kind$="Active"] *,
.stApp [data-testid="stButtonGroup"] button[aria-checked="true"] *,
.stApp [data-testid="stButtonGroup"] button[aria-pressed="true"] *{color:#ffffff !important}
.stApp [data-testid="stExpander"] details,
.stApp [data-testid="stExpander"] summary{background:#f1f5f9 !important;color:#111827 !important}
.stApp [data-testid="stExpander"] summary *{color:#111827 !important}
.stApp [data-testid="stExpander"] summary svg{fill:#111827 !important}
.stApp [data-testid="stExpander"] [data-testid="stExpanderDetails"]{background:#f8fafc !important}
.stApp [data-testid="stToggle"] label *,.stApp [data-testid="stCheckbox"] label *{color:#111827 !important}
</style>
"""


def _esc(x) -> str:
    return html.escape(str(x))


def _chip(rotulo: str, valor: str, ok: bool | None = True) -> str:
    classe = "ok" if ok else ("pend" if ok is False else "")
    icone = "✓" if ok else ("•" if ok is None else "!")
    return f'<span class="ctl-chip {classe}">{icone} {_esc(rotulo)}: <b>{_esc(valor)}</b></span>'


def _cartoes_status(contagens: dict, ordem=None, total: int | None = None) -> None:
    ordem = ordem or [s for s in ORDEM_STATUS if contagens.get(s)] + [s for s in contagens if s not in ORDEM_STATUS]
    total = total if total is not None else sum(contagens.values())
    partes = []
    for status in ordem:
        n = int(contagens.get(status, 0))
        fundo, texto, borda = CORES_STATUS.get(status, _NEUTRO)
        pct = f"{100 * n / total:.0f}%" if total else ""
        partes.append(
            f'<div class="ctl-card" style="background:{fundo};color:{texto};border-color:{borda}">'
            f'<div class="t">{_esc(NOMES_STATUS.get(status, status))}</div>'
            f'<div class="n">{f"{n:,}".replace(",", ".")}</div><div class="p">{pct}</div></div>'
        )
    st.markdown(f'<div class="ctl-cards">{"".join(partes)}</div>', unsafe_allow_html=True)


def _estilo_status(valor) -> str:
    fundo, texto, _ = CORES_STATUS.get(str(valor), (None, None, None))
    return f"background-color:{fundo};color:{texto};font-weight:600" if fundo else ""


def _tabela(dados: pd.DataFrame, essenciais: list[str] | None, key: str, colorir=("status",)) -> None:
    """Tabela enxuta: só as colunas essenciais, com "Mostrar todas as colunas"."""
    if dados is None or dados.empty:
        st.caption("Nada para mostrar.")
        return
    mostrar = dados
    if essenciais:
        uteis = [c for c in essenciais if c in dados.columns]
        if len(uteis) < len(dados.columns):
            _padrao(key, False)
            if st.toggle(f"Mostrar todas as colunas ({len(dados.columns)})", key=key):
                uteis = list(dados.columns)
        mostrar = dados[uteis]
    cols = [c for c in colorir if c in mostrar.columns]
    if cols and mostrar.size <= 150_000:
        estilo = mostrar.style
        aplicar = getattr(estilo, "map", None) or estilo.applymap
        st.dataframe(aplicar(_estilo_status, subset=cols), hide_index=True, use_container_width=True)
    else:
        st.dataframe(mostrar, hide_index=True, use_container_width=True)


def _downloads(itens: list[dict], colunas: int = 3) -> None:
    """Downloads agrupados em cartões: {titulo, desc, dados, arquivo, mime, icone}."""
    itens = [i for i in itens if i]
    for ini in range(0, len(itens), colunas):
        linha = st.columns(colunas)
        for col, item in zip(linha, itens[ini:ini + colunas]):
            with col:
                with st.container(border=True):
                    st.markdown(f'<div class="ctl-dl-t">{item.get("icone", "📄")} {_esc(item["titulo"])}</div>'
                                f'<div class="ctl-dl-d">{_esc(item["desc"])}</div>', unsafe_allow_html=True)
                    if item.get("acao"):
                        item["acao"]()
                    else:
                        vazio = not item.get("dados")
                        st.download_button(
                            f"Baixar {item['arquivo'].rsplit('.', 1)[-1].upper()}", item.get("dados") or b"",
                            item["arquivo"], item.get("mime", "application/octet-stream"),
                            key=f"ctl_dl_{item['arquivo']}", disabled=vazio, use_container_width=True,
                            type="primary" if item.get("principal") else "secondary",
                        )
                        if vazio and item.get("vazio_msg"):
                            st.caption(item["vazio_msg"])


# =====================================================================
# Estado das escolhas (guardado ao trocar de etapa ou de ferramenta)
# =====================================================================
# O Streamlit apaga o valor de um widget que não foi desenhado na execução.
# Guardamos uma cópia ("<key>__salvo") e devolvemos quando o widget volta.
_SUFIXO = "__salvo"


def _restaurar() -> None:
    for k in [k for k in list(st.session_state.keys()) if str(k).startswith("ctl_") and str(k).endswith(_SUFIXO)]:
        alvo = k[: -len(_SUFIXO)]
        if alvo not in st.session_state:
            st.session_state[alvo] = st.session_state[k]


# escolhas da tela que são guardadas (arquivos enviados e botões, não)
_GUARDAR_PREFIXOS = ("ctl_etapa", "ctl_area", "ctl_checagem", "ctl_id_column", "ctl_abrangencia", "ctl_estados",
                     "ctl_nova_", "ctl_hist_map_", "ctl_proc_var", "ctl_lote_", "ctl_realoc_child", "ctl_exp_status",
                     "ctl_dup_", "ctl_ufs_eleitoral", "ctl_eleitoral_", "ctl_cons_", "ctl_previa_")


def _guardar() -> None:
    for k in list(st.session_state.keys()):
        k = str(k)
        if k.endswith(_SUFIXO) or not (k.startswith(_GUARDAR_PREFIXOS) or (k.startswith("ctl_")
                                                                          and k.endswith("_todas"))):
            continue
        valor = st.session_state[k]
        if isinstance(valor, (str, int, float, bool, list)) or valor is None:
            st.session_state[k + _SUFIXO] = valor


def _padrao(key: str, valor) -> None:
    """Valor inicial de um widget sem passar `default`/`index` (evita o aviso
    do Streamlit quando o valor também vem do session_state)."""
    if key not in st.session_state:
        st.session_state[key] = valor


def _saneia(key: str, opcoes: list, multiplo: bool = False, padrao=None) -> None:
    """Valor guardado que não existe mais nas opções (outro banco, outra syntax) volta ao padrão."""
    if key not in st.session_state:
        if padrao is not None:
            st.session_state[key] = padrao
        return
    v = st.session_state[key]
    if multiplo:
        novo = [x for x in (v or []) if x in opcoes]
        if novo != v:
            st.session_state[key] = novo
    elif v not in opcoes:
        if padrao is not None and padrao in opcoes:
            st.session_state[key] = padrao
        elif opcoes:
            st.session_state[key] = opcoes[0]
        else:
            del st.session_state[key]


def _seletor(rotulo: str, opcoes: list, key: str, formato=str, padrao=None):
    """Seletor horizontal (pílulas no Streamlit novo; botões de rádio no antigo)."""
    _saneia(key, opcoes, padrao=padrao if padrao is not None else opcoes[0])
    if hasattr(st, "segmented_control"):
        valor = st.segmented_control(rotulo, opcoes, format_func=formato, key=key, label_visibility="collapsed")
        if valor is None:  # clicou de novo na opção já marcada
            valor = st.session_state.get(key + _SUFIXO, opcoes[0])
        return valor
    return st.radio(rotulo, opcoes, format_func=formato, key=key, horizontal=True, label_visibility="collapsed")


def _ir(key: str, valor):
    def _cb():
        st.session_state[key] = valor
    return _cb


# =====================================================================
# Faixa de status do projeto
# =====================================================================
def _faixa_status() -> None:
    chips = []
    if "df" in ss:
        nome = ss.get("sav_original_filename", "banco.sav")
        chips.append(_chip("Banco", f"{nome} · {len(ss.df):,} entrevistas".replace(",", "."), True))
    else:
        chips.append(_chip("Banco", "não carregado", False))
    if "label_sets" in ss:
        chips.append(_chip("Syntax", f"{len(ss.label_sets)} conjunto(s) de VALUE LABELS", bool(ss.label_sets)))
    else:
        chips.append(_chip("Syntax", "não carregada", False))
    if "geography_db" in ss:
        ufs = ss.get("ufs_pesquisa") or []
        chips.append(_chip("Base Brasil", ", ".join(ufs) if ufs else "Brasil inteiro", True))
    else:
        chips.append(_chip("Base Brasil", "não carregada", False))
    if "df" in ss and st.session_state.get("ctl_id_column"):
        chips.append(_chip("ID", st.session_state["ctl_id_column"], None))
    n_cfg = len(ss.get("control_configs") or [])
    chips.append(_chip("Variáveis de controle", str(n_cfg), True if n_cfg else None))
    st.markdown(f'<div class="ctl-faixa">{"".join(chips)}</div>', unsafe_allow_html=True)


# =====================================================================
# Fluxo 1 — Recodificação geográfica
# =====================================================================
ETAPAS = ["Arquivos", "Variáveis", "Processar", "Revisar", "Exportar"]
DESC_ETAPAS = [
    "Banco da pesquisa, VALUE LABELS do projeto, Base Brasil e abrangência da pesquisa.",
    "Quais variáveis de controle criar e de onde vêm as respostas.",
    "Interpreta as respostas na Base Brasil e confere contra os VALUE LABELS.",
    "Casos pendentes, aceitação em lote e realocação Cidade × Bairro.",
    "Resultado final e arquivos para baixar.",
]


def _pendentes() -> int:
    a = ss.get("audit_df")
    return 0 if a is None else int(a["status"].isin(PENDENTES).sum())


def _etapa_feita(i: int) -> bool:
    if i == 0:
        return "df" in ss and bool(ss.get("label_sets"))
    if i == 1:
        return bool(ss.get("control_configs"))
    if i == 2:
        return "audit_df" in ss
    if i == 3:
        return "audit_df" in ss and _pendentes() == 0
    return "sav_bytes" in ss


def _rotulo_etapa(i: int) -> str:
    if i == 3 and "audit_df" in ss and _pendentes():
        return f"{i + 1} · {ETAPAS[i]} ({_pendentes()})"
    return f"{'✓' if _etapa_feita(i) else i + 1} · {ETAPAS[i]}"


def _fluxo_recodificacao() -> None:
    etapa = _seletor("Etapa", list(range(len(ETAPAS))), "ctl_etapa", _rotulo_etapa, padrao=0)
    st.markdown(f'<div class="ctl-etapa">{_esc(DESC_ETAPAS[etapa])}</div>', unsafe_allow_html=True)

    [_etapa_arquivos, _etapa_variaveis, _etapa_processar, _etapa_revisar, _etapa_exportar][etapa]()

    st.divider()
    c1, _, c3 = st.columns([1, 3, 1])
    with c1:
        if etapa > 0:
            st.button(f"◀ {ETAPAS[etapa - 1]}", key="ctl_bt_voltar", use_container_width=True,
                      on_click=_ir("ctl_etapa", etapa - 1))
    with c3:
        if etapa < len(ETAPAS) - 1:
            st.button(f"{ETAPAS[etapa + 1]} ▶", key="ctl_bt_avancar", use_container_width=True,
                      type="primary" if _etapa_feita(etapa) else "secondary",
                      on_click=_ir("ctl_etapa", etapa + 1))


def _ir_arquivos():
    st.session_state["ctl_etapa"] = 0
    st.session_state["ctl_area"] = "recod"


def _precisa_arquivos(syntax: bool = True) -> bool:
    falta = "df" not in ss or (syntax and not ss.get("label_sets"))
    if falta:
        o_que = "o banco .sav e a syntax .sps" if syntax else "o banco .sav"
        st.info(f"Carregue {o_que} na etapa **Arquivos** da Recodificação geográfica.")
        st.button("Ir para Arquivos", key="ctl_bt_ir_arquivos", on_click=_ir_arquivos)
    return falta


def _carregar_banco(arquivo) -> None:
    dados = arquivo.getvalue()
    assinatura = (arquivo.name, len(dados))
    df, meta = _cached_read_sav(dados)
    if ss.get("banco_assinatura") != assinatura:
        # outro banco: resultados antigos não valem mais
        for k in _DERIVADOS + _DERIVADOS_CHECAGENS:
            ss.pop(k, None)
        ss.banco_assinatura = assinatura
    ss.df = df
    ss.meta = meta
    ss.sav_original_filename = arquivo.name


def _etapa_arquivos() -> None:
    c1, c2 = st.columns(2)
    with c1:
        with st.container(border=True):
            st.markdown("**Banco da pesquisa (.sav)**")
            sav_file = st.file_uploader("Banco da pesquisa (.sav)", type=["sav"], key="ctl_sav",
                                        label_visibility="collapsed")
            if sav_file:
                with st.spinner("Lendo banco..."):
                    _carregar_banco(sav_file)
            if "df" in ss:
                st.success(f"{ss.get('sav_original_filename', 'banco.sav')}: {len(ss.df):,} entrevistas e "
                           f"{len(ss.df.columns)} variáveis.".replace(",", "."))
    with c2:
        with st.container(border=True):
            st.markdown("**VALUE LABELS do projeto (.sps)**")
            sps_file = st.file_uploader("VALUE LABELS do projeto (.sps)", type=["sps"], key="ctl_sps",
                                        label_visibility="collapsed")
            if sps_file:
                ss.label_sets = _cached_parse_sps(decode_sps_bytes(sps_file.getvalue()))
                ss.sps_nome = sps_file.name
            if "label_sets" in ss:
                if ss.label_sets:
                    st.success(f"{ss.get('sps_nome', 'syntax')}: {len(ss.label_sets)} conjunto(s) de VALUE LABELS.")
                else:
                    st.error("Nenhum bloco VALUE LABELS válido foi encontrado na syntax.")

    if "df" in ss:
        df, meta = ss.df, ss.meta
        columns = list(df.columns)
        with st.container(border=True):
            st.markdown("**Projeto**")
            col_id, col_abr = st.columns([1, 2])
            with col_id:
                id_default = next((c for c in columns if c.lower() in {"respondent_id", "responde", "id"}),
                                  columns[0])
                _saneia("ctl_id_column", columns, padrao=id_default)
                st.selectbox("Variável identificadora (ID)", columns, key="ctl_id_column")
            with col_abr:
                _padrao("ctl_abrangencia_geografica", "Um ou mais estados específicos")
                abrangencia = st.radio("A pesquisa cobre:", ["Um ou mais estados específicos", "Brasil inteiro"],
                                       horizontal=True, key="ctl_abrangencia_geografica")
            if abrangencia == "Um ou mais estados específicos":
                _saneia("ctl_estados_pesquisa", UF_OPTIONS, multiplo=True)
                estados = st.multiselect(
                    "Estado(s) da pesquisa", UF_OPTIONS, key="ctl_estados_pesquisa",
                    help="Ex.: uma pesquisa no Rio de Janeiro deve marcar só 'RJ — Rio de Janeiro'.",
                    placeholder="Escolha o(s) estado(s)",
                )
                ss.ufs_pesquisa = [uf_option_to_sigla(o) for o in estados]
                if not estados:
                    st.caption("⚠️ Sem estado marcado, a Base Brasil do país inteiro é usada — mais ambiguidade "
                               "(ex.: 'Centro' existe em quase toda cidade).")
            else:
                ss.ufs_pesquisa = []
            if "geography_db" in ss and ss.get("ufs_pesquisa"):
                base_geo = ss.geography_db
                efetiva = base_geo.filter_by_uf(ss.ufs_pesquisa)
                st.caption(f"Base Brasil restrita a {', '.join(ss.ufs_pesquisa)}: "
                           f"{efetiva.summary()['registros']:,} de {base_geo.summary()['registros']:,} registros."
                           .replace(",", "."))
        with st.expander("Ver as variáveis do banco"):
            st.dataframe(pd.DataFrame(variable_catalog(df, meta)), hide_index=True, use_container_width=True)
    if ss.get("label_sets"):
        with st.expander("Ver os VALUE LABELS encontrados na syntax"):
            st.dataframe(pd.DataFrame([{"Conjunto": n, "Categorias": len(l)} for n, l in ss.label_sets.items()]),
                         hide_index=True, use_container_width=True)

    _bloco_base_brasil()
    _bloco_banco_anterior()


def _bloco_base_brasil() -> None:
    internal_geo = _load_internal_geography()
    with st.expander("Base Brasil (municípios, bairros e distritos)", expanded="geography_db" not in ss
                     and internal_geo is None):
        if internal_geo is not None:
            if "geo_alternativa" not in ss:
                ss.geography_db = internal_geo
            st.caption(f"Base interna: {internal_geo.source_name} — já vem com o app, não precisa enviar.")
            geo_file_override = st.file_uploader("Usar outra planilha só desta vez (opcional)",
                                                 type=["xlsx", "xls"], key="ctl_geo_override")
            if geo_file_override:
                try:
                    ss.geography_db = _cached_geography_from_bytes(geo_file_override.getvalue(),
                                                                   geo_file_override.name)
                    ss.geo_alternativa = geo_file_override.name
                except Exception as exc:
                    st.error(f"Não foi possível ler a Base Brasil enviada: {exc}")
            elif "geo_alternativa" in ss:
                ss.pop("geo_alternativa")
                ss.geography_db = internal_geo
            if "geo_alternativa" in ss:
                st.info(f"Usando a planilha enviada ({ss.geo_alternativa}) no lugar da base interna.")
        else:
            st.info("A Base Brasil não está embutida nesta cópia do app. Envie a planilha abaixo "
                    "(ou coloque-a em controle/data/ para não precisar enviar de novo).")
            geo_file = st.file_uploader("Base Brasil — Cidades, Bairros e Distritos (.xlsx/.xls)",
                                        type=["xlsx", "xls"], key="ctl_geo")
            if geo_file:
                try:
                    ss.geography_db = _cached_geography_from_bytes(geo_file.getvalue(), geo_file.name)
                except Exception as exc:
                    st.error(f"Não foi possível ler a Base Brasil: {exc}")
        if "geography_db" in ss:
            resumo = ss.geography_db.summary()
            for col, chave, titulo in zip(st.columns(5), ["registros", "municipios", "distritos", "bairros", "ufs"],
                                          ["Registros", "Municípios", "Distritos", "Bairros", "UFs"]):
                col.metric(titulo, f"{resumo[chave]:,}".replace(",", "."))
            if st.toggle("Prévia da Base Brasil", key="ctl_previa_base"):
                st.dataframe(ss.geography_db.dataframe().head(500), hide_index=True, use_container_width=True)


def _bloco_banco_anterior() -> None:
    with st.expander("Aprendizado com banco anterior (opcional)", expanded="historical_df" in ss):
        st.caption(
            "Em pesquisas de acompanhamento contínuo (tracking), o mesmo apelido, erro de digitação ou nome "
            "popular de bairro tende a se repetir. Com o banco de uma onda anterior já codificado, o app "
            "reaproveita essas decisões para os casos que a Base Brasil sozinha não resolve — só complementa o "
            "que ficaria pendente, nunca substitui uma decisão já confirmada nesta rodada."
        )
        historical_file = st.file_uploader("Banco anterior já codificado (.sav)", type=["sav"],
                                           key="ctl_historical_sav")
        if historical_file:
            with st.spinner("Lendo banco anterior..."):
                hist_df, hist_meta = _cached_read_sav(historical_file.getvalue())
            ss.historical_df = hist_df
            ss.historical_meta = hist_meta
            st.success(f"Banco anterior carregado: {len(hist_df):,} entrevistas, {len(hist_df.columns)} variáveis."
                       .replace(",", "."))
        elif "historical_df" in ss:
            st.info(f"Usando o banco anterior já carregado ({len(ss.historical_df):,} entrevistas)."
                    .replace(",", "."))
            if st.button("Remover banco anterior", key="ctl_bt_rem_hist"):
                for key in ["historical_df", "historical_meta", "historical_code_columns"]:
                    ss.pop(key, None)
                st.rerun()


TIPOS = {"Município": "municipio", "Bairro": "bairro", "Distrito": "distrito", "Não geográfica": None}
TIPO_NOME = {v: k for k, v in TIPOS.items()}


def _etapa_variaveis() -> None:
    if _precisa_arquivos():
        return
    df, label_sets = ss.df, ss.label_sets
    columns = list(df.columns)
    label_names = list(label_sets)
    col_labels = getattr(ss.meta, "column_names_to_labels", {}) or {}

    if ss.pop("limpar_form", False):
        for k in ("ctl_nova_nome", "ctl_nova_fontes"):
            st.session_state.pop(k, None)
            st.session_state.pop(k + _SUFIXO, None)

    configs = ss.control_configs
    if configs:
        st.markdown("**Variáveis de controle configuradas**")
        for i, c in enumerate(configs):
            with st.container(border=True):
                a, b = st.columns([6, 1])
                with a:
                    tipo = TIPO_NOME.get(c.get("geographic_type"), "Não geográfica")
                    st.markdown(
                        f'<div class="ctl-var"><b style="font-size:16px">{_esc(c["output_name"])}</b>'
                        f'<span class="ctl-tag">{_esc(tipo)}</span>'
                        f'<span class="ctl-tag cinza">VALUE LABELS: {_esc(c["label_set_name"])}</span></div>'
                        f'<div style="font-size:13px;color:#4b5563;margin-top:4px">Fontes, em ordem: '
                        f'{" → ".join(_esc(s) for s in c["source_columns"])}'
                        f' · fuzzy {c.get("fuzzy_cutoff", 0.82):.2f} / {c.get("auto_fuzzy_threshold", 0.93):.2f}</div>',
                        unsafe_allow_html=True)
                with b:
                    if st.button("Remover", key=f"ctl_bt_rem_cfg_{i}", use_container_width=True):
                        ss.control_configs = [x for j, x in enumerate(configs) if j != i]
                        for key in _DERIVADOS:
                            ss.pop(key, None)
                        st.rerun()
        if st.button("Limpar todas", key="ctl_bt_limpar_cfg"):
            ss.control_configs = []
            for key in _DERIVADOS:
                ss.pop(key, None)
            st.rerun()
        if "audit_df" in ss:
            st.caption("Mudar as variáveis apaga o processamento atual (é preciso processar de novo).")

    with st.container(border=True):
        st.markdown("**Nova variável de controle**")
        c1, c2, c3 = st.columns(3)
        with c1:
            output_name = st.text_input("Nome da nova variável", placeholder="Ex.: Cidades", key="ctl_nova_nome")
        with c2:
            _padrao("ctl_nova_tipo", "Município")
            tipo_ui = st.selectbox("Tipo de interpretação", list(TIPOS), key="ctl_nova_tipo",
                                   help="Uma fonte de bairro pode ser usada para descobrir o município "
                                        "quando o tipo for Município.")
        with c3:
            _saneia("ctl_nova_labels", label_names, padrao=label_names[0])
            label_set = st.selectbox("VALUE LABELS que define a amostra/códigos", label_names,
                                     key="ctl_nova_labels")
        _saneia("ctl_nova_fontes", columns, multiplo=True)
        source_columns = st.multiselect(
            "Variáveis-fonte (na ordem de prioridade)", columns, key="ctl_nova_fontes",
            format_func=lambda c: f"{c} — {col_labels[c]}" if col_labels.get(c) else c,
            help="Selecione também fontes auxiliares. Ex.: para Cidades, cidade + cidade_outros + bairro.",
        )
        with st.expander("Ajuste fino da comparação aproximada (fuzzy)"):
            c4, c5 = st.columns(2)
            with c4:
                _padrao("ctl_nova_fuzzy", 0.82)
                fuzzy_cutoff = st.slider("Mínimo para sugerir", 0.50, 1.00, step=0.01, key="ctl_nova_fuzzy")
            with c5:
                _padrao("ctl_nova_auto", 0.93)
                auto_fuzzy_threshold = st.slider("Mínimo para ajuste automático", 0.50, 1.00, step=0.01,
                                                 key="ctl_nova_auto")
        geographic_type = TIPOS[tipo_ui]
        if st.button("Adicionar variável de controle", type="primary", key="ctl_bt_add_cfg"):
            if not (output_name or "").strip():
                st.error("Informe o nome da nova variável.")
            elif not source_columns:
                st.error("Selecione ao menos uma variável-fonte.")
            elif geographic_type and "geography_db" not in ss:
                st.error("Para uma variável geográfica, carregue primeiro a Base Brasil (etapa Arquivos).")
            elif any(c["output_name"].casefold() == output_name.strip().casefold() for c in ss.control_configs):
                st.error("Já existe uma variável com esse nome.")
            else:
                ss.control_configs = ss.control_configs + [{
                    "output_name": output_name.strip(),
                    "label_set_name": label_set,
                    "source_columns": tuple(source_columns),
                    "geographic_type": geographic_type,
                    "fuzzy_cutoff": fuzzy_cutoff,
                    "auto_fuzzy_threshold": auto_fuzzy_threshold,
                }]
                for key in _DERIVADOS:
                    ss.pop(key, None)
                st.toast(f"Variável “{output_name.strip()}” adicionada.", icon="✅")
                ss.limpar_form = True
                st.rerun()

    if ss.control_configs and "historical_df" in ss:
        with st.container(border=True):
            st.markdown("**De onde vem o código já pronto no banco anterior**")
            hist_columns = list(ss.historical_df.columns)
            mapping = dict(ss.get("historical_code_columns", {}))
            for c in ss.control_configs:
                name = c["output_name"]
                default_col = name if name in hist_columns else (
                    c["label_set_name"] if c["label_set_name"] in hist_columns else hist_columns[0])
                chave = f"ctl_hist_map_{name}"
                _saneia(chave, hist_columns, padrao=mapping.get(name, default_col))
                mapping[name] = st.selectbox(f"Coluna com o código já pronto de “{name}”", hist_columns, key=chave)
            ss.historical_code_columns = mapping


def _processar() -> None:
    df, meta, label_sets = ss.df, ss.meta, ss.label_sets
    all_audits = []
    output_df = df.copy()
    generated_label_sets = {}
    bank_value_labels = getattr(meta, "variable_value_labels", {}) or {}
    geography_db = ss.get("geography_db")
    ufs_pesquisa = ss.get("ufs_pesquisa") or None
    if geography_db is not None and ufs_pesquisa:
        geography_db = geography_db.filter_by_uf(ufs_pesquisa)

    historical_df = ss.get("historical_df")
    historical_meta = ss.get("historical_meta")
    historical_mapping = ss.get("historical_code_columns", {})
    historical_value_labels = getattr(historical_meta, "variable_value_labels", {}) or {} if historical_meta else {}

    id_column = st.session_state.get("ctl_id_column") or df.columns[0]
    for raw in ss.control_configs:
        config = ControlVariableConfig(**raw)
        labels = label_sets[config.label_set_name]
        audit = recode_dataframe(df, config, labels, bank_value_labels, id_column, geography_db=geography_db)

        code_column = historical_mapping.get(config.output_name)
        if historical_df is not None and code_column:
            hierarchy_for_learning = (get_hierarchy("Campo Grande", geography_db)
                                      if config.geographic_type == "bairro" else None)
            lookup, _conflicts = build_text_to_code_lookup(
                historical_df, config.source_columns, code_column, historical_value_labels,
                hierarchy=hierarchy_for_learning, project_labels=labels,
            )
            audit = apply_learned_lookup(audit, lookup, labels)

        all_audits.append(audit)
        codes = pd.Series(pd.NA, index=df.index, dtype="Float64")
        auto = audit["decisao_automatica"] & audit["codigo_sugerido"].notna()
        codes.loc[auto.to_numpy()] = audit.loc[auto, "codigo_sugerido"].astype(float).to_numpy()
        output_df[config.output_name] = codes
        generated_label_sets[config.output_name] = labels

    audit_df = pd.concat(all_audits, ignore_index=True)
    for k in _DERIVADOS:
        ss.pop(k, None)
    ss.audit_df = audit_df
    ss.output_df = output_df
    ss.generated_label_sets = generated_label_sets
    ss.coherence_df = build_geographic_coherence_report(audit_df, ss.control_configs, label_sets)


COLS_AUDITORIA = ["ID", "variavel_controle", "texto_interpretado", "status", "label_sugerido", "confianca",
                  "metodo", "municipio_base", "fonte_utilizada"]
COLS_REVISAO = ["ID", "variavel_controle", "texto_interpretado", "status", "label_sugerido", "candidatos",
                "localidade_base", "municipio_base", "uf_base", "confianca"]
COLS_COERENCIA = ["ID", "status", "cidade_controle_atual", "localidade_informada", "municipio_base", "recomendacao"]


def _etapa_processar() -> None:
    if _precisa_arquivos():
        return
    if not ss.control_configs:
        st.info("Configure ao menos uma variável de controle na etapa **Variáveis**.")
        st.button("Ir para Variáveis", key="ctl_bt_ir_vars", on_click=_ir("ctl_etapa", 1))
        return
    feito = "audit_df" in ss
    rotulo = "Processar de novo" if feito else "Processar controle"
    if st.button(rotulo, type="secondary" if feito else "primary", key="ctl_bt_processar"):
        with st.spinner("Interpretando respostas e verificando a Base Brasil..."):
            _processar()
        st.toast("Processamento concluído.", icon="✅")
        feito = True
    if not feito:
        st.caption(f"{len(ss.control_configs)} variável(is) de controle × {len(ss.df):,} entrevistas."
                   .replace(",", "."))
        return

    audit_df = ss.audit_df
    _cartoes_status(audit_df["status"].value_counts().to_dict())
    variaveis = list(dict.fromkeys(audit_df["variavel_controle"]))
    if len(variaveis) > 1:
        escolha = _seletor("Variável", ["(todas)"] + variaveis, "ctl_proc_var", padrao="(todas)")
        if escolha != "(todas)":
            audit_df = audit_df[audit_df["variavel_controle"] == escolha]
            _cartoes_status(audit_df["status"].value_counts().to_dict())
    _tabela(audit_df, COLS_AUDITORIA, "ctl_proc_todas")
    coherence = ss.get("coherence_df", pd.DataFrame())
    if coherence is not None and not coherence.empty:
        st.error(f"{len(coherence)} inconsistência(s) de coerência Cidade × Bairro — veja a etapa **Revisar**.")
    pend = _pendentes()
    if pend:
        st.warning(f"{pend} caso(s) pendente(s) de revisão.")


def _etapa_revisar() -> None:
    if "audit_df" not in ss:
        st.info("Execute o processamento na etapa **Processar**.")
        st.button("Ir para Processar", key="ctl_bt_ir_proc", on_click=_ir("ctl_etapa", 2))
        return
    df, meta, label_sets = ss.df, ss.meta, ss.label_sets
    audit_df = ss.audit_df
    review = audit_df[audit_df["status"].isin(PENDENTES)].copy()

    aba_pend, aba_lote, aba_realoc = st.tabs([
        f"Pendências ({len(review)})",
        f"Aceitar sugestões em lote ({int((audit_df['status'] == 'SUGESTÃO').sum())})",
        "Cidade × Bairro",
    ])
    with aba_pend:
        if review.empty:
            st.success("Não há pendências individuais de matching para revisão.")
        else:
            _cartoes_status(review["status"].value_counts().to_dict())
            cols_review = COLS_REVISAO + [c for c in ["texto_normalizado", "fonte_utilizada", "tipo_localidade_base",
                                                      "metodo"] if c not in COLS_REVISAO]
            _tabela(review.reindex(columns=cols_review, fill_value=""), COLS_REVISAO, "ctl_rev_todas")

    with aba_lote:
        st.caption("Aplica de uma vez os casos com status SUGESTÃO acima de uma confiança mínima — você decide o "
                   "corte. Fica registrado na auditoria como “ACEITO EM LOTE” (não é a mesma coisa que automático).")
        sugestoes = audit_df[audit_df["status"] == "SUGESTÃO"].copy()
        if sugestoes.empty:
            st.info("Não há sugestões pendentes no momento.")
        else:
            vars_disponiveis = sorted(sugestoes["variavel_controle"].unique())
            col_a, col_b = st.columns([2, 1])
            with col_a:
                _saneia("ctl_lote_vars", vars_disponiveis, multiplo=True, padrao=vars_disponiveis)
                vars_escolhidas = st.multiselect("Variáveis", vars_disponiveis, key="ctl_lote_vars")
            with col_b:
                _padrao("ctl_lote_corte", 0.85)
                corte_confianca = st.slider("Confiança mínima", 0.50, 1.00, step=0.01, key="ctl_lote_corte")
            elegiveis = sugestoes[
                sugestoes["variavel_controle"].isin(vars_escolhidas)
                & (sugestoes["confianca"] >= corte_confianca)
                & sugestoes["codigo_sugerido"].notna()
            ]
            st.markdown(f"**{len(elegiveis)}** de {len(sugestoes)} sugestões atingem esse corte.")
            if not elegiveis.empty:
                _tabela(elegiveis[["ID", "variavel_controle", "texto_interpretado", "label_sugerido", "metodo",
                                   "confianca"]], None, "ctl_lote_tab")
            if st.button(f"Aplicar aceitação em lote ({len(elegiveis)})", type="primary", disabled=elegiveis.empty,
                         key="ctl_bt_lote"):
                output_df = ss.output_df
                audit_df_atualizado = ss.audit_df.copy()
                for col in ("status", "decisao_automatica", "codigo_sugerido", "label_sugerido"):
                    if col in audit_df_atualizado.columns:
                        audit_df_atualizado[col] = audit_df_atualizado[col].astype(object)
                for var in vars_escolhidas:
                    var_audit = audit_df_atualizado[audit_df_atualizado["variavel_controle"] == var]
                    mask = ((var_audit["status"] == "SUGESTÃO") & (var_audit["confianca"] >= corte_confianca)
                            & var_audit["codigo_sugerido"].notna())
                    if var not in output_df.columns or not mask.any():
                        continue
                    positions = mask.to_numpy()
                    output_df.loc[output_df.index[positions], var] = (
                        var_audit.loc[mask, "codigo_sugerido"].astype(float).to_numpy())
                    idx_global = var_audit.index[mask]
                    audit_df_atualizado.loc[idx_global, "status"] = "ACEITO EM LOTE"
                    audit_df_atualizado.loc[idx_global, "decisao_automatica"] = True
                ss.output_df = output_df
                ss.audit_df = audit_df_atualizado
                ss.pop("sav_bytes", None)
                st.toast(f"{len(elegiveis)} caso(s) aceito(s) em lote.", icon="✅")
                st.rerun()

    with aba_realoc:
        coherence = ss.get("coherence_df", pd.DataFrame())
        if coherence is not None and not coherence.empty:
            st.markdown(f"**Coerência Cidade × Bairro/Distrito** — {len(coherence)} caso(s)")
            _tabela(coherence, COLS_COERENCIA, "ctl_coh_todas")
        st.markdown("**Verificação cruzada Cidade × Bairro (realocação)**")
        st.caption(
            "Confere as duas pontas: (1) o bairro respondido pertence, na Base Brasil, a outro município da "
            "amostra; (2) o que foi escrito no campo de Bairro é, na verdade, o nome de outra cidade. Só sugere "
            "quando o município alternativo está nos VALUE LABELS do projeto — nunca aplica sozinho."
        )
        configs = ss.control_configs
        municipal_cfgs = [c for c in configs if c.get("geographic_type") == "municipio"]
        child_cfgs = [c for c in configs if c.get("geographic_type") in {"bairro", "distrito"}]
        if not municipal_cfgs or not child_cfgs:
            st.info("Configure uma variável de Município e uma de Bairro/Distrito (etapa Variáveis) para usar "
                    "esta verificação.")
            return
        city_cfg = municipal_cfgs[0]
        nomes_filhos = [c["output_name"] for c in child_cfgs]
        if len(child_cfgs) == 1:
            child_cfg_choice = child_cfgs[0]
        else:
            _saneia("ctl_realoc_child", nomes_filhos, padrao=nomes_filhos[0])
            escolhido = st.selectbox("Variável de Bairro/Distrito a usar na verificação", nomes_filhos,
                                     key="ctl_realoc_child")
            child_cfg_choice = child_cfgs[nomes_filhos.index(escolhido)]
        if st.button("Verificar realocações Cidade × Bairro", key="ctl_bt_realoc"):
            with st.spinner("Cruzando as duas respostas..."):
                geography_db_atual = ss.get("geography_db")
                ufs = ss.get("ufs_pesquisa") or None
                if geography_db_atual is not None and ufs:
                    geography_db_atual = geography_db_atual.filter_by_uf(ufs)
                bank_value_labels_atual = getattr(meta, "variable_value_labels", {}) or {}
                ss.realocacao_df = build_cidade_bairro_realocation(
                    df, ss.audit_df, city_cfg, child_cfg_choice, label_sets, bank_value_labels_atual,
                    geography_db_atual, st.session_state.get("ctl_id_column") or df.columns[0],
                )
        realoc_df = ss.get("realocacao_df")
        if realoc_df is None:
            return
        if realoc_df.empty:
            st.success("Nenhuma realocação sugerida com os dados atuais.")
            return
        aplicaveis = realoc_df[realoc_df["status"] == "INCONSISTENTE"]
        ambiguos = realoc_df[realoc_df["status"] == "AMBÍGUO"]
        _cartoes_status({"INCONSISTENTE": len(aplicaveis), "AMBÍGUO": len(ambiguos)},
                        ordem=["INCONSISTENTE", "AMBÍGUO"])
        if not aplicaveis.empty:
            st.markdown(f"**{len(aplicaveis)} entrevista(s) com realocação de Cidade sugerida**")
            _tabela(aplicaveis, ["ID", "cidade_controle_atual", "localidade_informada", "municipio_base",
                                 "motivo_realocacao"], "ctl_realoc_todas")
        if not ambiguos.empty:
            st.info(f"{len(ambiguos)} caso(s) em que o texto do Bairro bate como bairro de uma cidade E como nome de "
                    "outra cidade — as duas leituras discordam, então nada é sugerido. Revise manualmente:")
            _tabela(ambiguos, ["ID", "cidade_controle_atual", "localidade_informada", "municipio_base"],
                    "ctl_realoc_amb_todas")
        if not aplicaveis.empty and st.button(f"Aplicar realocações à variável de Cidade ({len(aplicaveis)})",
                                              type="primary", key="ctl_bt_aplicar_realoc"):
            output_df = ss.output_df
            audit_df_atualizado = ss.audit_df.copy()
            city_var = city_cfg["output_name"]
            if city_var not in output_df.columns:
                st.error(f"A variável '{city_var}' não existe no banco processado.")
                return
            for col in ("status", "decisao_automatica", "codigo_sugerido", "label_sugerido"):
                if col in audit_df_atualizado.columns:
                    audit_df_atualizado[col] = audit_df_atualizado[col].astype(object)
            id_series = df[st.session_state.get("ctl_id_column") or df.columns[0]].astype(str)
            realoc_ids = aplicaveis["ID"].astype(str)
            code_by_id = dict(zip(realoc_ids, pd.to_numeric(aplicaveis["codigo_cidade_sugerido"])))
            label_by_id = dict(zip(realoc_ids, aplicaveis["municipio_base"]))
            mask = id_series.isin(realoc_ids)
            output_df.loc[mask, city_var] = id_series[mask].map(code_by_id).astype(float).to_numpy()

            audit_ids = audit_df_atualizado["ID"].astype(str)
            row_mask = (audit_df_atualizado["variavel_controle"] == city_var) & audit_ids.isin(realoc_ids)
            audit_df_atualizado.loc[row_mask, "status"] = "REALOCADO (Cidade × Bairro)"
            audit_df_atualizado.loc[row_mask, "codigo_sugerido"] = audit_ids[row_mask].map(code_by_id).to_numpy()
            audit_df_atualizado.loc[row_mask, "label_sugerido"] = audit_ids[row_mask].map(label_by_id).to_numpy()
            audit_df_atualizado.loc[row_mask, "decisao_automatica"] = True

            ss.output_df = output_df
            ss.audit_df = audit_df_atualizado
            ss.pop("sav_bytes", None)
            st.toast(f"{len(aplicaveis)} entrevista(s) realocada(s).", icon="✅")
            st.rerun()


def _etapa_exportar() -> None:
    if "audit_df" not in ss:
        st.info("Execute o processamento na etapa **Processar**.")
        st.button("Ir para Processar", key="ctl_bt_ir_proc2", on_click=_ir("ctl_etapa", 2))
        return
    df, meta = ss.df, ss.meta
    audit_df = ss.audit_df
    _cartoes_status(audit_df["status"].value_counts().to_dict())

    with st.expander("Auditoria completa (com filtro por status)"):
        todos = sorted(audit_df["status"].unique())
        _saneia("ctl_exp_status", todos, multiplo=True, padrao=todos)
        status_filter = st.multiselect("Filtrar status", todos, key="ctl_exp_status")
        _tabela(audit_df[audit_df["status"].isin(status_filter)], COLS_AUDITORIA, "ctl_exp_todas")

    # .xlsx do banco: só é refeito quando o banco controlado muda
    assinatura = int(pd.util.hash_pandas_object(ss.output_df, index=True).sum())
    if ss.get("xlsx_cache", (None,))[0] != assinatura:
        buffer = io.BytesIO()
        ss.output_df.to_excel(buffer, index=False)
        ss.xlsx_cache = (assinatura, buffer.getvalue())

    coherence = ss.get("coherence_df", pd.DataFrame())
    original_name = ss.get("sav_original_filename", "banco.sav")
    base_name = original_name[:-4] if original_name.lower().endswith(".sav") else original_name

    def _acao_sav():
        if "sav_bytes" not in ss:
            if st.button("Gerar .sav", key="ctl_bt_gerar_sav", type="primary", use_container_width=True):
                with st.spinner("Gravando .sav (pode levar alguns segundos em bancos grandes)..."):
                    sav_output_df = df.copy()
                    sav_column_labels = dict(getattr(meta, "column_names_to_labels", {}) or {})
                    sav_value_labels = dict(getattr(meta, "variable_value_labels", {}) or {})
                    for config_raw in ss.control_configs:
                        name = config_raw["output_name"]
                        sav_output_df[name] = ss.output_df[name]
                        sav_column_labels[name] = f"{name} — gerado pelo Controle Geral"
                        sav_value_labels[name] = ss.generated_label_sets[name]
                    ss.sav_bytes = write_sav_bytes(
                        sav_output_df, column_labels=sav_column_labels, variable_value_labels=sav_value_labels,
                        file_label="Banco controlado - Gerador de Controle Geral",
                    )
                st.rerun()
        else:
            st.download_button("Baixar SAV", ss.sav_bytes, f"{base_name}_CONTROLADO.sav", "application/octet-stream",
                               key="ctl_dl_sav", type="primary", use_container_width=True)

    _downloads([
        {"titulo": "Banco controlado (.sav)", "icone": "🗂️", "acao": _acao_sav,
         "desc": "Banco original + as variáveis de controle (código e VALUE LABELS), pronto para o SPSS. "
                 "Casos ainda pendentes ficam em branco."},
        {"titulo": "Auditoria", "icone": "🔎", "dados": audit_df.to_csv(index=False).encode("utf-8-sig"),
         "arquivo": "controle_geral_auditoria.csv", "mime": "text/csv",
         "desc": "Uma linha por entrevista e variável: resposta, fonte, método, confiança, status e localidade."},
        {"titulo": "Banco provisório", "icone": "📊", "dados": ss.xlsx_cache[1],
         "arquivo": "banco_controlado_provisorio.xlsx",
         "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
         "desc": "O banco com as novas variáveis em Excel, para conferência rápida."},
        {"titulo": "VALUE LABELS gerados", "icone": "🏷️",
         "dados": make_value_labels_syntax(ss.generated_label_sets).encode("utf-8-sig"),
         "arquivo": "aplicar_value_labels_controle.sps", "mime": "text/plain",
         "desc": "Syntax para aplicar os rótulos das novas variáveis no SPSS."},
        {"titulo": "Inconsistências geográficas", "icone": "⚠️",
         "dados": coherence.to_csv(index=False).encode("utf-8-sig") if coherence is not None and not coherence.empty
         else None,
         "arquivo": "inconsistencias_geograficas.csv", "mime": "text/csv",
         "desc": "Casos em que o bairro informado pertence a outro município.",
         "vazio_msg": "Nenhuma inconsistência neste processamento."},
    ])


# =====================================================================
# Fluxo 2 — Checagens
# =====================================================================
def _consolidado() -> pd.DataFrame | None:
    dup_df = ss.get("duplicates_df")
    ele_df = ss.get("eleitoral_df")
    tem_dup = dup_df is not None and not dup_df.empty
    tem_ele = ele_df is not None and not ele_df.empty
    if not tem_dup and not tem_ele:
        return None
    marc = st.session_state
    pieces = []
    if tem_dup:
        for tipo, chave, nome, padrao in (("certa", "ctl_cons_dup_certa", "Duplicidade certa", True),
                                          ("provavel", "ctl_cons_dup_prov", "Duplicidade provável", False),
                                          ("possivel", "ctl_cons_dup_poss", "Duplicidade possível", False)):
            if marc.get(chave, marc.get(chave + _SUFIXO, padrao)):
                sel = dup_df[(dup_df["tipo_duplicidade"] == tipo) & (dup_df["recomendacao"] == "excluir")]
                pieces += [{"ID": str(r["ID"]), "motivo": f"{nome} — {r['motivo']}"} for _, r in sel.iterrows()]
    if tem_ele:
        if marc.get("ctl_cons_ele_fora", marc.get("ctl_cons_ele_fora" + _SUFIXO, True)):
            fora = ele_df[ele_df["status"] == "VOTA EM OUTRO ESTADO"]
            pieces += [{"ID": str(r["ID"]),
                        "motivo": f"Vota em outro estado ({r['uf_identificada']}) — {r['texto_interpretado']}"}
                       for _, r in fora.iterrows()]
        if marc.get("ctl_cons_ele_naovota", marc.get("ctl_cons_ele_naovota" + _SUFIXO, False)):
            nao = ele_df[ele_df["status"] == "NÃO VOTA"]
            pieces += [{"ID": str(r["ID"]), "motivo": f"Não vota — {r['texto_interpretado']}"}
                       for _, r in nao.iterrows()]
    if not pieces:
        return pd.DataFrame(columns=["ID", "motivo"])
    consolidated = pd.DataFrame(pieces)
    grouped = consolidated.groupby("ID")["motivo"].apply(lambda s: " | ".join(s)).reset_index()
    return grouped.sort_values("ID")


def _resumo_checagens() -> None:
    dup_df = ss.get("duplicates_df")
    ele_df = ss.get("eleitoral_df")
    cons = _consolidado()
    cards = []
    if dup_df is None:
        cards.append('<div class="ctl-res vazio"><div class="h">Duplicidades</div><div class="n">Não analisado</div>'
                     '<div class="d">Nome + telefone</div></div>')
    else:
        excl = dup_df[dup_df["recomendacao"] == "excluir"]["tipo_duplicidade"].value_counts()
        classe = "alerta" if int(excl.get("certa", 0)) else "ok"
        cards.append(f'<div class="ctl-res {classe}"><div class="h">Duplicidades</div>'
                     f'<div class="n">{int(excl.get("certa", 0))} certas</div>'
                     f'<div class="d">{int(excl.get("provavel", 0))} prováveis · {int(excl.get("possivel", 0))} '
                     f'possíveis (a excluir)</div></div>')
    if ele_df is None:
        cards.append('<div class="ctl-res vazio"><div class="h">Validação eleitoral</div>'
                     '<div class="n">Não analisado</div><div class="d">Onde o entrevistado vota</div></div>')
    else:
        cont = ele_df["status"].value_counts()
        fora = int(cont.get("VOTA EM OUTRO ESTADO", 0))
        cards.append(f'<div class="ctl-res {"alerta" if fora else "ok"}"><div class="h">Validação eleitoral</div>'
                     f'<div class="n">{fora} fora</div><div class="d">{int(cont.get("NÃO VOTA", 0))} não votam · '
                     f'{int(cont.get("VOTA NO ESTADO DA PESQUISA", 0))} no estado</div></div>')
    if cons is None:
        cards.append('<div class="ctl-res vazio"><div class="h">Exclusões consolidadas</div>'
                     '<div class="n">—</div><div class="d">Rode as checagens primeiro</div></div>')
    else:
        cards.append(f'<div class="ctl-res"><div class="h">Exclusões consolidadas</div>'
                     f'<div class="n">{len(cons)} IDs</div><div class="d">com o motivo de cada um</div></div>')
    st.markdown(f'<div class="ctl-resumo">{"".join(cards)}</div>', unsafe_allow_html=True)


def _fluxo_checagens() -> None:
    if _precisa_arquivos(syntax=False):
        st.caption("As checagens usam só o banco .sav (a syntax não é necessária).")
        return
    _resumo_checagens()

    def _rot(op):
        return {"dup": "Duplicidades", "ele": "Validação eleitoral", "exc": "Exclusões consolidadas"}[op]
    sub = _seletor("Checagem", ["dup", "ele", "exc"], "ctl_checagem", _rot, padrao="dup")
    {"dup": _checagem_duplicidades, "ele": _checagem_eleitoral, "exc": _checagem_exclusoes}[sub]()


def _checagem_duplicidades() -> None:
    df, meta = ss.df, ss.meta
    st.caption("Entrevistas duplicadas por nome + telefone. As variáveis são achadas pelo LABEL no banco "
               "(ex.: 'Nome', 'Tel.') — confira e ajuste se precisar.")
    column_labels = getattr(meta, "column_names_to_labels", {}) or {}
    auto_name = find_column_by_label(column_labels, NAME_LABEL_PATTERNS)
    auto_phone = find_column_by_label(column_labels, PHONE_LABEL_PATTERNS)
    auto_date = find_column_by_label(column_labels, DATE_LABEL_PATTERNS, also_check_names=True)

    none_option = "(nenhuma)"
    options = [none_option] + list(df.columns)

    def _label_of(col):
        if col == none_option:
            return col
        lbl = column_labels.get(col, "")
        return f"{col} — {lbl}" if lbl else col

    with st.container(border=True):
        col1, col2, col3 = st.columns(3)
        escolhas = {}
        for col, chave, rot, auto in ((col1, "ctl_dup_name_col", "Variável de Nome", auto_name),
                                      (col2, "ctl_dup_phone_col", "Variável de Telefone", auto_phone),
                                      (col3, "ctl_dup_date_col", "Variável de Data (mantém a mais antiga)",
                                       auto_date)):
            with col:
                _saneia(chave, options, padrao=auto if auto in options else none_option)
                escolhas[chave] = st.selectbox(rot, options, format_func=_label_of, key=chave)
        name_col, phone_col, date_col = (escolhas["ctl_dup_name_col"], escolhas["ctl_dup_phone_col"],
                                         escolhas["ctl_dup_date_col"])
        if name_col == none_option and phone_col == none_option:
            st.warning("Selecione ao menos Nome ou Telefone para buscar duplicidades.")
        elif st.button("Analisar duplicidades", type="primary", key="ctl_bt_dup"):
            with st.spinner("Comparando nomes e telefones..."):
                dup_config = DuplicateConfig(
                    id_column=st.session_state.get("ctl_id_column") or df.columns[0],
                    name_column=None if name_col == none_option else name_col,
                    phone_column=None if phone_col == none_option else phone_col,
                    date_column=None if date_col == none_option else date_col,
                )
                ss.duplicates_df = find_duplicates(df, dup_config)
                ss.duplicates_id_column = dup_config.id_column
            st.rerun()

    if "duplicates_df" not in ss:
        return
    dup_df = ss.duplicates_df
    if dup_df.empty:
        st.success("Nenhuma duplicidade encontrada com os critérios atuais.")
        return
    nomes_tipo = {"certa": "Certa (mesmo telefone)", "provavel": "Provável (mesmo nome + telefone parecido)",
                  "possivel": "Possível (mesmo nome, telefones diferentes ou ausentes)"}
    _saneia("ctl_dup_tipos", list(nomes_tipo), multiplo=True, padrao=["certa", "provavel"])
    tipo_filter = st.multiselect("Tipos de duplicidade", list(nomes_tipo), format_func=nomes_tipo.get,
                                 key="ctl_dup_tipos")
    dup_df_filtrado = dup_df[dup_df["tipo_duplicidade"].isin(tipo_filter)]
    _cartoes_status({
        "Grupos": dup_df_filtrado["grupo"].nunique(),
        "Entrevistas envolvidas": len(dup_df_filtrado),
        "excluir": int((dup_df_filtrado["recomendacao"] == "excluir").sum()),
    }, ordem=["Grupos", "Entrevistas envolvidas", "excluir"], total=0)
    _tabela(dup_df_filtrado, ["grupo", "ID", "nome", "telefone", "tipo_duplicidade", "recomendacao", "motivo"],
            "ctl_dup_todas", colorir=("tipo_duplicidade", "recomendacao"))
    st.caption("Nenhuma entrevista é excluída do banco aqui — só se você aplicar a syntax no SPSS. Os arquivos "
               "seguem o filtro de tipos acima.")

    ids_excluir = dup_df_filtrado.loc[dup_df_filtrado["recomendacao"] == "excluir", "ID"].tolist()
    id_col_name = ss.duplicates_id_column
    id_is_string = not pd.api.types.is_numeric_dtype(df[id_col_name])
    syntax = make_exclusion_syntax(id_col_name, ids_excluir, id_is_string,
                                   comment="Mantida a entrevista mais antiga de cada grupo; demais marcadas para "
                                           "exclusão.")
    _downloads([
        {"titulo": "IDs a excluir", "icone": "📝", "dados": "\n".join(ids_excluir).encode("utf-8-sig"),
         "arquivo": "entrevistas_duplicadas_excluir.txt", "mime": "text/plain",
         "desc": f"{len(ids_excluir)} ID(s), um por linha.", "vazio_msg": "Nenhum ID a excluir."},
        {"titulo": "Syntax de exclusão", "icone": "🧾", "dados": syntax.encode("utf-8-sig") if ids_excluir else None,
         "arquivo": "excluir_duplicidades.sps", "mime": "text/plain",
         "desc": "SELECT IF para apagar os duplicados no SPSS.", "vazio_msg": "Nenhum ID a excluir."},
    ], colunas=2)
    with st.expander("Ver a syntax gerada"):
        st.code(syntax, language="sql")


def _checagem_eleitoral() -> None:
    df, meta = ss.df, ss.meta
    st.caption("Verifica se cada entrevistado vota no(s) estado(s) da pesquisa, em outro estado, ou se declarou "
               "que não vota. Usa a Base Brasil nacional (não a restrita ao estado da pesquisa).")
    with st.container(border=True):
        ufs_pesquisa = ss.get("ufs_pesquisa") or []
        if ufs_pesquisa:
            st.markdown(f"Estado(s) da pesquisa (etapa Arquivos): **{', '.join(ufs_pesquisa)}**")
            ufs_eleitoral = ufs_pesquisa
        else:
            _saneia("ctl_ufs_eleitoral_manual", UF_OPTIONS, multiplo=True)
            ufs_escolhidas = st.multiselect(
                "Estado(s) considerado(s) 'dentro da pesquisa' para esta checagem", UF_OPTIONS,
                key="ctl_ufs_eleitoral_manual",
                help="A abrangência da pesquisa está como 'Brasil inteiro', então escolha aqui.")
            ufs_eleitoral = [uf_option_to_sigla(o) for o in ufs_escolhidas]

        columns = list(df.columns)
        col_labels = getattr(meta, "column_names_to_labels", {}) or {}

        def format_col(c):
            return f"{c} — {col_labels.get(c, '')}" if col_labels.get(c) else c

        c1, c2 = st.columns(2)
        with c1:
            _saneia("ctl_eleitoral_var_principal", columns, padrao=columns[0])
            var_principal = st.selectbox("Variável principal (onde vota)", columns, format_func=format_col,
                                         key="ctl_eleitoral_var_principal")
        with c2:
            opcoes_outro = ["(nenhuma)"] + columns
            _saneia("ctl_eleitoral_var_outro", opcoes_outro, padrao="(nenhuma)")
            var_outro = st.selectbox("Variável 'Outros' (texto livre), se houver", opcoes_outro,
                                     format_func=lambda c: c if c == "(nenhuma)" else format_col(c),
                                     key="ctl_eleitoral_var_outro")
        source_columns = tuple([var_principal] + ([var_outro] if var_outro != "(nenhuma)" else []))

        bank_value_labels_local = getattr(meta, "variable_value_labels", {}) or {}
        var_principal_labels = bank_value_labels_local.get(var_principal, {})
        if var_principal_labels:
            chave = "ctl_eleitoral_nao_vota_codes"
            codigos = list(var_principal_labels.keys())
            if st.session_state.get("ctl_eleitoral_codes_var") != var_principal:
                # variável trocada: volta para a sugestão automática
                st.session_state[chave] = detect_nao_vota_codes(var_principal_labels)
                st.session_state["ctl_eleitoral_codes_var"] = var_principal
            _saneia(chave, codigos, multiplo=True)
            nao_vota_escolhidos = st.multiselect(
                "Códigos que significam 'não vota' / 'não vota mais' (sugestão automática já marcada)", codigos,
                format_func=lambda c: f"{c} — {var_principal_labels.get(c, '')}", key=chave)
        else:
            st.caption("Essa variável não tem VALUE LABELS no banco — nenhum código de 'não vota' sugerido.")
            nao_vota_escolhidos = []

        if st.button("Verificar validação eleitoral", type="primary", disabled=not ufs_eleitoral,
                     key="ctl_bt_eleitoral"):
            geo_nacional = ss.get("geography_db")
            if geo_nacional is None:
                st.error("Base Brasil não carregada — veja a etapa Arquivos.")
            else:
                with st.spinner("Verificando local de votação..."):
                    ss.eleitoral_df = validate_electoral(
                        df, st.session_state.get("ctl_id_column") or df.columns[0], source_columns,
                        set(nao_vota_escolhidos), bank_value_labels_local, geo_nacional, set(ufs_eleitoral))
                st.rerun()
        if not ufs_eleitoral:
            st.caption("Escolha ao menos um estado para habilitar a checagem.")

    eleitoral_df = ss.get("eleitoral_df")
    if eleitoral_df is None or eleitoral_df.empty:
        return
    _cartoes_status(eleitoral_df["status"].value_counts().to_dict(),
                    ordem=[s for s in ["VOTA NO ESTADO DA PESQUISA", "VOTA EM OUTRO ESTADO", "NÃO VOTA", "AMBÍGUO",
                                       "NÃO IDENTIFICADO", "VAZIO"] if s in set(eleitoral_df["status"])])
    _tabela(eleitoral_df, ["ID", "status", "texto_interpretado", "uf_identificada", "municipio_identificado"],
            "ctl_ele_todas")
    fora = eleitoral_df[eleitoral_df["status"] == "VOTA EM OUTRO ESTADO"]
    _downloads([
        {"titulo": "IDs que votam fora do estado", "icone": "📝",
         "dados": "\n".join(fora["ID"].astype(str)).encode("utf-8-sig") if not fora.empty else None,
         "arquivo": "vota_fora_do_estado.txt", "mime": "text/plain",
         "desc": f"{len(fora)} ID(s), um por linha.", "vazio_msg": "Ninguém vota fora do estado."},
    ], colunas=2)


def _checagem_exclusoes() -> None:
    df = ss.df
    st.caption("Junta os candidatos a exclusão das Duplicidades e da Validação eleitoral num único arquivo, com o "
               "motivo de cada ID — para não mandar dois arquivos separados para aprovação.")
    dup_df = ss.get("duplicates_df")
    ele_df = ss.get("eleitoral_df")
    tem_dup = dup_df is not None and not dup_df.empty
    tem_ele = ele_df is not None and not ele_df.empty
    if not tem_dup and not tem_ele:
        st.info("Rode as Duplicidades e/ou a Validação eleitoral primeiro.")
        return

    with st.container(border=True):
        st.markdown("**O que entra na exclusão**")
        c1, c2 = st.columns(2)
        with c1:
            if tem_dup:
                exc = dup_df[dup_df["recomendacao"] == "excluir"]["tipo_duplicidade"].value_counts()
                for chave, rot, tipo, padrao in (
                        ("ctl_cons_dup_certa", "Duplicidade certa — mesmo telefone", "certa", True),
                        ("ctl_cons_dup_prov", "Duplicidade provável — mesmo nome + telefone parecido", "provavel",
                         False),
                        ("ctl_cons_dup_poss", "Duplicidade possível — mesmo nome, telefones diferentes ou ausentes",
                         "possivel", False)):
                    _padrao(chave, padrao)
                    st.checkbox(f"{rot} ({int(exc.get(tipo, 0))})", key=chave)
            else:
                st.caption("Duplicidades: não analisado.")
        with c2:
            if tem_ele:
                cont = ele_df["status"].value_counts()
                _padrao("ctl_cons_ele_fora", True)
                st.checkbox(f"Vota em outro estado ({int(cont.get('VOTA EM OUTRO ESTADO', 0))})",
                            key="ctl_cons_ele_fora")
                _padrao("ctl_cons_ele_naovota", False)
                st.checkbox(f"Não vota / não vota mais ({int(cont.get('NÃO VOTA', 0))})", key="ctl_cons_ele_naovota")
            else:
                st.caption("Validação eleitoral: não analisada.")

    grouped = _consolidado()
    if grouped is None or grouped.empty:
        st.info("Nenhuma categoria marcada acima.")
        return
    st.markdown(f"**{len(grouped)}** entrevista(s) única(s) selecionada(s) para exclusão.")
    _tabela(grouped, None, "ctl_cons_tab")
    id_col_name = st.session_state.get("ctl_id_column") or df.columns[0]
    id_is_string = not pd.api.types.is_numeric_dtype(df[id_col_name])
    syntax = make_exclusion_syntax_with_reasons(
        id_col_name, list(zip(grouped["ID"], grouped["motivo"])), id_is_string,
        titulo="Exclusões consolidadas — Duplicidades + Validação Eleitoral",
    )
    _downloads([
        {"titulo": "IDs consolidados", "icone": "📝", "dados": "\n".join(grouped["ID"]).encode("utf-8-sig"),
         "arquivo": "exclusoes_consolidadas.txt", "mime": "text/plain", "desc": f"{len(grouped)} ID(s), um por linha."},
        {"titulo": "Syntax consolidada", "icone": "🧾", "dados": syntax.encode("utf-8-sig"),
         "arquivo": "exclusoes_consolidadas.sps", "mime": "text/plain", "principal": True,
         "desc": "SELECT IF com o motivo de cada ID documentado em comentário."},
    ], colunas=2)
    with st.expander("Ver a syntax gerada"):
        st.code(syntax, language="sql")


# =====================================================================
# Página
# =====================================================================
def modulo_controle():
    st.markdown(CSS, unsafe_allow_html=True)
    st.header("Controle Geral")
    st.caption("Recodificação geográfica com a Base Brasil, duplicidades e validação eleitoral.")

    _restaurar()
    # Versão do formato dos objetos gravados na sessão: uma sessão antiga, de
    # antes de uma mudança na auditoria, perde só os resultados derivados.
    SESSION_SCHEMA_VERSION = 2
    if ss.get("session_schema_version") != SESSION_SCHEMA_VERSION:
        for key in ["audit_df", "output_df", "generated_label_sets", "coherence_df", "realocacao_df"]:
            ss.pop(key, None)
        ss.session_schema_version = SESSION_SCHEMA_VERSION
    if "control_configs" not in ss:
        ss.control_configs = []

    _faixa_status()

    def _rot_area(op):
        if op == "recod":
            return "🗺️ Recodificação geográfica"
        return "✅ Checagens"
    try:
        area = _seletor("Área", ["recod", "checagens"], "ctl_area", _rot_area, padrao="recod")
        if area == "recod":
            _fluxo_recodificacao()
        else:
            _fluxo_checagens()
    finally:
        _guardar()
