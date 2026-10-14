"""Controle Geral dentro do Informa: Base Brasil no pacote e estado com prefixo."""
import sys
import types

if "streamlit" not in sys.modules:
    _st = types.ModuleType("streamlit")
    _st.session_state = {}
    _st.cache_data = _st.cache_resource = lambda *a, **k: (a[0] if a and callable(a[0]) else (lambda f: f))
    sys.modules["streamlit"] = _st

import streamlit as st  # noqa: E402

for _nome in ("cache_data", "cache_resource"):  # outros testes podem ter criado um streamlit simulado sem eles
    if not hasattr(st, _nome):
        setattr(st, _nome, lambda *a, **k: (a[0] if a and callable(a[0]) else (lambda f: f)))
if not hasattr(st, "session_state"):
    st.session_state = {}

from controle import PASTA_DADOS  # noqa: E402


def test_base_brasil_dentro_do_pacote():
    assert (PASTA_DADOS / "Cidades_Bairros_Distritos - Brasil.xlsx").exists()


def test_estado_com_prefixo():
    from analises.controle import ss
    ss.df = "banco"
    assert st.session_state["ctl_df"] == "banco"
    assert "df" in ss and ss.get("df") == "banco"
    assert ss.pop("df") == "banco" and "ctl_df" not in st.session_state
