"""Testes do app original Controle-Automatico (imports com o prefixo controle.)."""
from controle.core.matching import confidence_status, match_text

def test_matching_accent_insensitive():
    result = match_text("sao goncalo", {1: "São Gonçalo", 2: "Rio de Janeiro"})
    assert result.code == 1 and result.method == "exata"
    assert confidence_status(result) == "CONFIRMADO"

def test_prefix_matching():
    result = match_text("Barra", {1: "Barra da Tijuca", 2: "Botafogo"})
    assert result.code == 1 and result.method == "prefixo"
