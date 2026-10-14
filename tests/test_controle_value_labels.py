"""Testes do app original Controle-Automatico (imports com o prefixo controle.)."""
from controle.core.value_labels import parse_all_value_labels_sps

def test_parse_multiple_blocks_and_variables():
    syntax = """VALUE LABELS Cidades\n1 'Rio de Janeiro'\n2 'Niterói'\n.\nVALUE LABELS Bairros Bairros2\n1 'Centro'\n2 'D''Ávila'\n."""
    parsed = parse_all_value_labels_sps(syntax)
    assert parsed["Cidades"][2] == "Niterói"
    assert parsed["Bairros"][2] == "D'Ávila"
    assert parsed["Bairros2"][1] == "Centro"
