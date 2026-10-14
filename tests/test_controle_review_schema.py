"""Testes do app original Controle-Automatico (imports com o prefixo controle.)."""
import pandas as pd


def test_review_reindex_tolerates_legacy_audit_without_geographic_columns():
    legacy = pd.DataFrame([
        {
            "ID": 1,
            "variavel_controle": "Cidades",
            "texto_interpretado": "Niteroi",
            "fonte_utilizada": "q0002",
            "status": "SUGESTÃO",
            "label_sugerido": "Niterói",
            "candidatos": "",
            "metodo": "fuzzy",
            "confianca": 0.91,
        }
    ])
    cols_review = [
        "ID", "variavel_controle", "texto_interpretado", "fonte_utilizada", "status",
        "label_sugerido", "localidade_base", "tipo_localidade_base", "municipio_base",
        "uf_base", "candidatos", "metodo", "confianca",
    ]
    result = legacy.reindex(columns=cols_review, fill_value="")
    assert list(result.columns) == cols_review
    assert result.loc[0, "localidade_base"] == ""
    assert result.loc[0, "municipio_base"] == ""
    assert result.loc[0, "uf_base"] == ""
