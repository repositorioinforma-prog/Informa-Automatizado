import ast
from pathlib import Path


def _fonte_ponderacao() -> str:
    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    return caminho.read_text(encoding="utf-8")


def test_interface_nao_exibe_escolha_de_metodo_ou_tolerancia():
    fonte = _fonte_ponderacao()

    assert 'st.subheader("4. Método de ponderação")' not in fonte
    assert 'key="ponderacao_metodo"' not in fonte
    assert 'key="ponderacao_tolerancia"' not in fonte
    assert 'key="ponderacao_max_iter"' not in fonte
    assert "Ajuste iterativo — tentar aproximar todas as margens" not in fonte


def test_parametros_institucionais_sao_fixos():
    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    arvore = ast.parse(caminho.read_text(encoding="utf-8"))

    constantes = {}
    for no in arvore.body:
        if not isinstance(no, ast.Assign) or len(no.targets) != 1:
            continue
        alvo = no.targets[0]
        if isinstance(alvo, ast.Name) and isinstance(no.value, ast.Constant):
            constantes[alvo.id] = no.value.value

    assert constantes["METODO_PONDERACAO_FIXO"] == (
        "Método atual — razão simples e multiplicação"
    )
    assert constantes["TOLERANCIA_PONDERACAO_FIXA_PP"] == 2.0


def test_calculo_visivel_utiliza_apenas_razao_simples():
    fonte = _fonte_ponderacao()

    assert "calcular_raking" not in fonte
    assert "resultado = calcular_razao_simples(" in fonte
    assert 'st.subheader("4. Amostra e universo")' in fonte
    assert "teste o ajuste iterativo" not in fonte
