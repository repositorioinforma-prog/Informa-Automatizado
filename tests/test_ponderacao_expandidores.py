import ast
from pathlib import Path


def test_janelas_de_universo_iniciam_recolhidas():
    """Os editores de cada variável devem ocupar espaço somente quando abertos."""
    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    arvore = ast.parse(caminho.read_text(encoding="utf-8"))

    chamadas_alvo = []
    for no in ast.walk(arvore):
        if not isinstance(no, ast.Call):
            continue
        if not (
            isinstance(no.func, ast.Attribute)
            and isinstance(no.func.value, ast.Name)
            and no.func.value.id == "st"
            and no.func.attr == "expander"
        ):
            continue
        if not no.args:
            continue
        titulo = no.args[0]
        if (
            isinstance(titulo, ast.Call)
            and isinstance(titulo.func, ast.Name)
            and titulo.func.id == "_descricao_variavel"
        ):
            chamadas_alvo.append(no)

    assert len(chamadas_alvo) == 1
    kwargs = {kw.arg: kw.value for kw in chamadas_alvo[0].keywords if kw.arg}
    assert "expanded" in kwargs
    assert isinstance(kwargs["expanded"], ast.Constant)
    assert kwargs["expanded"].value is False
