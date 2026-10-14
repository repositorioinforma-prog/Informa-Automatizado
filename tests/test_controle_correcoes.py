"""Correções do Controle Geral feitas na lapidação (v31)."""
from io import BytesIO

import pandas as pd

from controle.core.duplicates import DuplicateConfig, find_duplicates, normalize_phone
from controle.core.geographic_validation import build_cidade_bairro_realocation
from controle.core.recoding import ControlVariableConfig, recode_dataframe
from controle.geography.database import load_geography_excel


def _base():
    frame = pd.DataFrame([
        {"UF": "RJ", "Município": "Rio de Janeiro", "Distrito": "", "Bairro": "Copacabana"},
        {"UF": "RJ", "Município": "Niterói", "Distrito": "", "Bairro": "Icaraí"},
        {"UF": "MS", "Município": "Campo Grande", "Distrito": "", "Bairro": "Centro"},
    ])
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        frame.to_excel(w, sheet_name="Base Brasil", index=False)
    return load_geography_excel(buf.getvalue())


def test_realocacao_sem_inconsistencias_nao_quebra():
    db = _base()
    df = pd.DataFrame({"id": [1], "cidade": ["Rio de Janeiro"], "bairro": ["Copacabana"]})
    city = {"output_name": "Cidades", "label_set_name": "Cidades", "source_columns": ("cidade",),
            "geographic_type": "municipio"}
    child = {"output_name": "Bairros", "label_set_name": "Bairros", "source_columns": ("bairro",),
             "geographic_type": "bairro"}
    labels = {"Cidades": {1: "Rio de Janeiro", 2: "Niterói"}, "Bairros": {1: "Copacabana"}}
    audits = [recode_dataframe(df, ControlVariableConfig(**c), labels[c["label_set_name"]], {}, "id", geography_db=db)
              for c in (city, child)]
    audit = pd.concat(audits, ignore_index=True)
    out = build_cidade_bairro_realocation(df, audit, city, child, labels, {}, db, "id")
    assert out.empty


def test_loteamento_de_campo_grande_ms_nao_vale_em_pesquisa_do_rj():
    db = _base()
    df = pd.DataFrame({"id": [1], "bairro": ["Jardim Aclimação"]})
    cfg = ControlVariableConfig("Bairros", "Bairros", ("bairro",), "bairro")
    labels = {1: "Centro", 2: "Copacabana"}
    rj = recode_dataframe(df, cfg, labels, {}, "id", geography_db=db.filter_by_uf(["RJ"]))
    assert not rj.loc[0, "decisao_automatica"] and pd.isna(rj.loc[0, "codigo_sugerido"])
    ms = recode_dataframe(df, cfg, labels, {}, "id", geography_db=db.filter_by_uf(["MS"]))
    assert ms.loc[0, "codigo_sugerido"] == 1 and ms.loc[0, "decisao_automatica"]


def test_telefones():
    assert normalize_phone(11987654321.0) == "11987654321"
    assert normalize_phone("(21) 98765-4321") == "21987654321"
    for vazio in (0, 0.0, "0", "00000000", "99999999", "1234", "", None, float("nan")):
        assert normalize_phone(vazio) == ""


def test_duplicidades_sem_telefone_de_preenchimento_e_mantida_primeiro():
    df = pd.DataFrame({"id": [1, 2, 3, 4], "nome": ["Ana", "Bia", "Caio", "Davi"],
                       "tel": ["0", "0", 11987654321.0, "(11) 98765-4321"]})
    r = find_duplicates(df, DuplicateConfig("id", "nome", "tel"))
    assert list(r["ID"]) == ["3", "4"]
    assert list(r["recomendacao"]) == ["manter", "excluir"]


def test_sav_sem_pyreadstat():
    from core.pyreadstat_seguro import pyreadstat
    from controle.exporters.sav import write_sav_bytes
    if not pyreadstat.usando_leitor_puro:
        return  # com pyreadstat de verdade o caminho é o original
    df = pd.DataFrame({"id": [1.0, 2.0], "Cidades": pd.array([2, None], dtype="Float64")})
    dados = write_sav_bytes(df, {"Cidades": "Cidades"}, {"Cidades": {1: "Rio de Janeiro", 2: "Niterói"}})
    lido, meta = pyreadstat.read_sav(dados)
    assert lido["Cidades"].iloc[0] == 2 and pd.isna(lido["Cidades"].iloc[1])
    assert meta.variable_value_labels["Cidades"][2.0] == "Niterói"
