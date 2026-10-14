"""Testes do app original Controle-Automatico (imports com o prefixo controle.)."""
from io import BytesIO

import pandas as pd

from controle.geography.database import load_geography_excel
from controle.geography.resolver import resolve_to_target
from controle.core.recoding import ControlVariableConfig, recode_dataframe
from controle.core.geographic_validation import build_geographic_coherence_report


def make_geo_bytes():
    frame = pd.DataFrame([
        {"UF": "RJ", "Município": "Rio de Janeiro", "Distrito": "", "Bairro": "Copacabana"},
        {"UF": "RJ", "Município": "Niterói", "Distrito": "", "Bairro": "Icaraí"},
        {"UF": "RJ", "Município": "Duque de Caxias", "Distrito": "", "Bairro": "Olavo Bilac"},
        {"UF": "RJ", "Município": "Teresópolis", "Distrito": "Vale do Bonsucesso", "Bairro": ""},
    ])
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Base Brasil", index=False)
    return buf.getvalue()


def test_load_and_resolve_parent_municipality():
    db = load_geography_excel(make_geo_bytes())
    result = resolve_to_target("Olavo Bilac", "municipio", db)
    assert result.value == "Duque de Caxias"
    assert result.matched_type == "bairro"
    assert result.status == "CONFIRMADO"


def test_same_source_can_recode_city_from_bairro():
    db = load_geography_excel(make_geo_bytes())
    df = pd.DataFrame({"id": [1], "cidade": ["Rio de Janeiro"], "bairro": ["Olavo Bilac"]})
    cfg = ControlVariableConfig(
        output_name="Cidades",
        label_set_name="Cidades",
        source_columns=("bairro",),
        geographic_type="municipio",
    )
    labels = {1: "Rio de Janeiro", 2: "Duque de Caxias"}
    audit = recode_dataframe(df, cfg, labels, {}, "id", geography_db=db)
    assert audit.loc[0, "codigo_sugerido"] == 2
    assert audit.loc[0, "label_sugerido"] == "Duque de Caxias"
    assert audit.loc[0, "decisao_automatica"]


def test_known_geography_outside_project_is_not_auto_recoded():
    db = load_geography_excel(make_geo_bytes())
    df = pd.DataFrame({"id": [1], "bairro": ["Olavo Bilac"]})
    cfg = ControlVariableConfig(
        output_name="Cidades",
        label_set_name="Cidades",
        source_columns=("bairro",),
        geographic_type="municipio",
    )
    labels = {1: "Rio de Janeiro", 3: "Niterói"}
    audit = recode_dataframe(df, cfg, labels, {}, "id", geography_db=db)
    assert audit.loc[0, "status"] == "FORA DA AMOSTRA"
    assert not audit.loc[0, "decisao_automatica"]
    assert pd.isna(audit.loc[0, "codigo_sugerido"])


def test_coherence_flags_bairro_parent_city():
    audit = pd.DataFrame([
        {"ID": 10, "variavel_controle": "Cidades", "label_sugerido": "Rio de Janeiro", "texto_interpretado": "Rio de Janeiro", "municipio_base": "Rio de Janeiro", "localidade_base": "Rio de Janeiro"},
        {"ID": 10, "variavel_controle": "Bairros", "label_sugerido": "Olavo Bilac", "texto_interpretado": "Olavo Bilac", "municipio_base": "Duque de Caxias", "localidade_base": "Olavo Bilac"},
    ])
    configs = [
        {"output_name": "Cidades", "label_set_name": "Cidades", "geographic_type": "municipio"},
        {"output_name": "Bairros", "label_set_name": "Bairros", "geographic_type": "bairro"},
    ]
    label_sets = {"Cidades": {1: "Rio de Janeiro", 2: "Duque de Caxias"}, "Bairros": {1: "Olavo Bilac"}}
    issues = build_geographic_coherence_report(audit, configs, label_sets)
    assert len(issues) == 1
    assert issues.loc[0, "status"] == "INCONSISTENTE"
    assert issues.loc[0, "codigo_cidade_sugerido"] == 2
