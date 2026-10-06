import ast
from pathlib import Path


def test_travas_e_protecoes_iniciam_recolhidas():
    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    arvore = ast.parse(caminho.read_text(encoding="utf-8"))

    chamadas = []
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
            isinstance(titulo, ast.Constant)
            and titulo.value == "Travas e proteções de exclusão"
        ):
            chamadas.append(no)

    assert len(chamadas) == 1
    kwargs = {kw.arg: kw.value for kw in chamadas[0].keywords if kw.arg}
    assert "expanded" in kwargs
    assert isinstance(kwargs["expanded"], ast.Constant)
    assert kwargs["expanded"].value is False
