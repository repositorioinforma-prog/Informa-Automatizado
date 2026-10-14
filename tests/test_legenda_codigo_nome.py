from core.legendas_math import _extrair_codigo_numerico_inicial as cod


def test_codigo_com_nome_e_hifen():
    assert cod("1 - Mambucaba (27,28%): Bairro A, Bairro B") == "1"
    assert cod("2 – Serra") == "2"


def test_faixas_continuam_fora():
    assert cod("16 - 24") == ""
    assert cod("16 a 29") == ""
