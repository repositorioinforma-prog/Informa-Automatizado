import pandas as pd

from core.revisor_syntax_math import review_syntax


def _codes(result):
    return {f["code"] for f in result["findings"]}


def test_detecta_regra_duplicada_e_value_label_nao_gerado():
    sps = b"""RECODE Q1 (1=1) (1=1) (2=2) INTO Q1R.\nVALUE LABELS Q1R 1 'Sim' 2 'Nao' 98 'NS'.\n"""
    result = review_syntax(sps)
    codes = _codes(result)
    assert "RECODE_REGRA_DUPLICADA" in codes
    assert "VALUE_LABEL_NAO_GERADO" in codes


def test_detecta_valor_real_que_vira_sysmis():
    sps = b"RECODE Q1 (1=1) (2=2) (ELSE=SYSMIS) INTO Q1R.\n"
    df = pd.DataFrame({"Q1": [1, 2, 3, 3]})
    result = review_syntax(sps, dataframe=df)
    hits = [f for f in result["findings"] if f["code"] == "VALOR_REAL_VIRA_SYSMIS"]
    assert len(hits) == 1
    assert "3" in hits[0]["message"]
    assert "2 caso(s)" in hits[0]["message"]


def test_detecta_variavel_origem_inexistente():
    sps = b"RECODE NAO_EXISTE (1=1) INTO X.\n"
    df = pd.DataFrame({"Q1": [1]})
    result = review_syntax(sps, dataframe=df)
    assert "VARIAVEL_ORIGEM_INEXISTENTE" in _codes(result)


def test_detecta_encoding_declarado_divergente():
    text = "* Encoding: UTF-8.\nVAR LABEL P1 'Faixa Etaria'.\n".replace("Etaria", "Et\xe1ria")
    data = text.encode("cp1252")
    result = review_syntax(data)
    assert result["encoding"] == "Windows-1252"
    assert "ENCODING_DECLARADO_DIVERGENTE" in _codes(result)


def test_detecta_secao_sem_programacao():
    sps = b"*****Q6*****\n*****Q7*****\nRECODE Q7 (1=1) INTO Q7R.\n"
    result = review_syntax(sps)
    assert any(f["code"] == "SECAO_SEM_PROGRAMACAO" and f["variable"] == "Q6" for f in result["findings"])


def test_detecta_erro_textual_conhecido():
    sps = "VARIABLE LABELS Q1R 'Deputado Estudual'.\n".encode("cp1252")
    result = review_syntax(sps)
    assert "POSSIVEL_ERRO_TEXTUAL" in _codes(result)


def test_value_labels_com_barra_dentro_do_rotulo_nao_sao_cortados():
    from core.revisor_syntax_math import parse_syntax
    sps = b"VALUE LABELS /Q1R 97 'Branco/Nulo' 98 'Nao irei votar' /Q2R 1 'Espirita/Kardecista' 2 'Outra'.\n"
    parsed, _ = parse_syntax(sps)
    assert parsed.value_labels["Q1R"] == {97: "Branco/Nulo", 98: "Nao irei votar"}
    assert parsed.value_labels["Q2R"] == {1: "Espirita/Kardecista", 2: "Outra"}


def test_ponto_decimal_nao_encerra_o_comando():
    from core.revisor_syntax_math import split_commands
    cmds = split_commands("RECODE X (1.5=2) INTO Y.\nEXECUTE.\n")
    assert [c for _, c in cmds] == ["RECODE X (1.5=2) INTO Y.", "EXECUTE."]
