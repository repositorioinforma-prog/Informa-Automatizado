from collections import OrderedDict

import numpy as np
import pandas as pd

from core.ponderacao_math import (
    aplicar_fatores,
    calcular_razao_simples,
    calcular_raking,
    diagnosticar_componentes,
    gerar_syntax_spss,
)


def _cenario(qtd=16):
    dados = pd.DataFrame({f"V{i}": [1, 2, 1, 2] * 10 for i in range(1, qtd + 1)})
    alvos = OrderedDict(
        (f"V{i}", OrderedDict([(1, 50.0), (2, 50.0)]))
        for i in range(1, qtd + 1)
    )
    return dados, alvos


def _cenario_regional():
    regioes = [1, 2, 3, 4, 5] * 8
    dados = pd.DataFrame(
        {
            "P1": [1, 2] * 20,
            "P2": [1, 2, 3, 4] * 10,
            "P3": [1, 2, 3, 4, 1] * 8,
            "P7": regioes,
        }
    )

    nomes = ["VN_sud", "VN_nord", "VN_sul", "VN_norte", "VN_centro"]
    for codigo_regiao, variavel in enumerate(nomes, start=1):
        valores = []
        contador = 0
        for regiao in regioes:
            if regiao == codigo_regiao:
                valores.append(1 if contador % 2 == 0 else 2)
                contador += 1
            else:
                valores.append(np.nan)
        dados[variavel] = valores

    alvos = OrderedDict(
        [
            ("P1", OrderedDict([(1, 50.0), (2, 50.0)])),
            ("P2", OrderedDict([(1, 25.0), (2, 25.0), (3, 25.0), (4, 25.0)])),
            ("P3", OrderedDict([(1, 40.0), (2, 20.0), (3, 20.0), (4, 20.0)])),
            ("P7", OrderedDict([(1, 20.0), (2, 20.0), (3, 20.0), (4, 20.0), (5, 20.0)])),
        ]
    )
    for variavel in nomes:
        alvos[variavel] = OrderedDict([(1, 50.0), (2, 50.0)])

    componentes = {
        "P1": 1,
        "P2": 2,
        "P3": 3,
        "P7": 4,
        **{variavel: 5 for variavel in nomes},
    }
    return dados, alvos, componentes


def test_syntax_padrao_cria_um_componente_por_variavel_sem_limite_de_cinco():
    dados, alvos = _cenario(16)
    resultado = calcular_razao_simples(dados, alvos)
    syntax = gerar_syntax_spss(resultado.fatores)

    assert "COMPUTE peso1=$SYSMIS." in syntax
    assert "COMPUTE peso16=$SYSMIS." in syntax
    assert "COMPUTE peso=peso1*peso2*peso3*peso4*peso5*peso6*peso7*peso8*peso9*peso10*peso11*peso12*peso13*peso14*peso15*peso16." in syntax


def test_varias_variaveis_podem_compartilhar_o_mesmo_componente():
    dados, alvos = _cenario(6)
    componentes = {f"V{i}": 1 for i in range(1, 7)}
    resultado = calcular_razao_simples(
        dados,
        alvos,
        componentes_por_variavel=componentes,
    )
    syntax = gerar_syntax_spss(
        resultado.fatores,
        componentes_por_variavel=componentes,
    )

    assert resultado.pesos.notna().all()
    assert "COMPUTE peso1=$SYSMIS." in syntax
    assert "COMPUTE #pndok1=0." in syntax
    assert "* V1 -> peso1." in syntax
    assert "* V6 -> peso1." in syntax
    assert "COMPUTE peso=peso1." in syntax
    assert "peso2" not in syntax


def test_componentes_podem_ser_repetidos_e_pular_numeracao():
    dados, alvos = _cenario(6)
    componentes = {"V1": 1, "V2": 1, "V3": 4, "V4": 4, "V5": 16, "V6": 16}
    resultado = calcular_razao_simples(
        dados,
        alvos,
        componentes_por_variavel=componentes,
    )
    syntax = gerar_syntax_spss(resultado.fatores, componentes_por_variavel=componentes)

    assert resultado.pesos.notna().all()
    assert "COMPUTE peso=peso1*peso4*peso16." in syntax
    assert "peso2" not in syntax
    assert "peso3" not in syntax


def test_componente_combina_fatores_disponiveis_e_aceita_variaveis_alternativas():
    dados = pd.DataFrame(
        {
            "A": [1, np.nan, 1, np.nan],
            "B": [np.nan, 2, 2, np.nan],
        }
    )
    fatores = OrderedDict(
        [
            ("A", OrderedDict([(1, 2.0)])),
            ("B", OrderedDict([(2, 3.0)])),
        ]
    )

    pesos = aplicar_fatores(
        dados,
        fatores,
        componentes_por_variavel={"A": 1, "B": 1},
    )

    assert pesos.iloc[0] == 2.0
    assert pesos.iloc[1] == 3.0
    assert pesos.iloc[2] == 6.0
    assert pd.isna(pesos.iloc[3])


def test_cinco_variaveis_regionais_no_peso5_nao_zeram_a_base():
    dados, alvos, componentes = _cenario_regional()

    resultado = calcular_razao_simples(
        dados,
        alvos,
        componentes_por_variavel=componentes,
    )
    cobertura = diagnosticar_componentes(
        dados,
        resultado.fatores,
        componentes_por_variavel=componentes,
    ).set_index("componente")

    assert resultado.pesos.notna().sum() == len(dados)
    assert np.isfinite(resultado.maior_diferenca_pp)
    assert resultado.diagnostico["percentual_com_peso"].notna().all()
    assert cobertura.at["peso5", "casos_validos"] == len(dados)
    assert cobertura.at["peso5", "casos_sem_valor_aplicavel"] == 0
    assert cobertura.at["peso5", "casos_com_multiplas_contribuicoes"] == 0


def test_raking_aceita_margens_regionais_mutuamente_exclusivas_no_mesmo_componente():
    dados, alvos, componentes = _cenario_regional()

    resultado = calcular_raking(
        dados,
        alvos,
        tolerancia_pp=0.01,
        max_iteracoes=20,
        componentes_por_variavel=componentes,
    )

    assert resultado.pesos.notna().sum() == len(dados)
    assert np.isfinite(resultado.maior_diferenca_pp)
    assert resultado.diagnostico["percentual_com_peso"].notna().all()


def test_syntax_de_componente_regional_ignora_ausencias_e_usa_fator_disponivel():
    dados, alvos, componentes = _cenario_regional()
    resultado = calcular_razao_simples(
        dados,
        alvos,
        componentes_por_variavel=componentes,
    )
    syntax = gerar_syntax_spss(
        resultado.fatores,
        componentes_por_variavel=componentes,
    )

    assert "*************** peso5: VN_sud, VN_nord, VN_sul, VN_norte, VN_centro." in syntax
    assert "COMPUTE #pndok5=0." in syntax
    assert "COMPUTE #pnderr5=0." in syntax
    assert "ELSE IF (NOT MISSING(VN_sud))." in syntax
    assert "ELSE IF (NOT MISSING(VN_centro))." in syntax
    assert "IF (#pndok5=0 OR #pnderr5=1) peso5=$SYSMIS." in syntax
    assert "COMPUTE peso=peso1*peso2*peso3*peso4*peso5." in syntax


def test_interface_repassa_componentes_para_todos_os_calculos():
    import ast
    from pathlib import Path

    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    arvore = ast.parse(caminho.read_text(encoding="utf-8"))

    nomes_calculo = {"calcular_razao_simples", "calcular_raking"}
    chamadas_calculo = []
    chamadas_recalculo = []

    for no in ast.walk(arvore):
        if not isinstance(no, ast.Call):
            continue
        if isinstance(no.func, ast.Name) and no.func.id in nomes_calculo:
            chamadas_calculo.append(no)
        if isinstance(no.func, ast.Name) and no.func.id == "_calcular_ponderacao_para_teste_exclusao":
            chamadas_recalculo.append(no)

    assert chamadas_calculo
    assert chamadas_recalculo
    for chamada in [*chamadas_calculo, *chamadas_recalculo]:
        palavras = {kw.arg for kw in chamada.keywords if kw.arg}
        assert "componentes_por_variavel" in palavras
