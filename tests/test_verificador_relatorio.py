"""Testes do verificador de divergências (planilha sintética, sem arquivo real).

Rodar da raiz do app:  python -m pytest tests/test_verificador_relatorio.py
"""
import openpyxl

from core.verificador_relatorio_math import montar_inventario, comparar


def _tabela(ws, r, titulo, opcoes, base, colunas=("Total",), pergunta="Pergunta: Teste"):
    ws.cell(r, 1, f"Titulo: {titulo}")
    for i, c in enumerate(colunas):
        ws.cell(r + 1, 2 + i, c)
    r += 2
    for op in opcoes:
        ws.cell(r, 1, op)
        for i in range(len(colunas)):
            ws.cell(r, 2 + i, 10.0)
        r += 1
    ws.cell(r, 1, "Base")
    for i in range(len(colunas)):
        ws.cell(r, 2 + i, base)
    ws.cell(r + 1, 1, pergunta)
    return r + 3


def _wb(totais, segmentos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(1, 1, "RESULTADOS PELO TOTAL")
    r = 3
    for t in totais:
        r = _tabela(ws, r, *t)
    ws.cell(r, 1, "RESULTADOS PELOS SEGMENTOS")
    r += 2
    for s in segmentos:
        r = _tabela(ws, r, *s, colunas=("Total", "Homem", "Mulher"))
    return ws


def _tipos(ws):
    return [a["tipo"] for a in comparar(montar_inventario(ws))]


def test_tudo_igual_nao_gera_achado():
    ws = _wb(
        [("Avaliação do governo", ["Ótima", "Ruim", "Não sabe"], 400)],
        [("Avaliação do governo", ["Ótima", "Ruim", "Não sabe"], 400)],
    )
    assert _tipos(ws) == []


def test_titulo_parecido_mas_diferente():
    ws = _wb(
        [("Intenção de voto para Deputado Federal/2026 (Espontânea e Única)", ["A", "B"], 400)],
        [("Intenção de voto para Deputado Federal/26 (Espontânea e Única)", ["A", "B"], 400)],
    )
    assert _tipos(ws) == ["TITULO_DIVERGENTE"]


def test_acento_e_caixa_nao_contam_como_divergencia():
    ws = _wb(
        [("Avaliação do Governo", ["Ótima", "Ruim"], 400)],
        [("AVALIACAO DO GOVERNO", ["Otima", "Ruim"], 400)],
    )
    assert _tipos(ws) == []


def test_opcao_com_grafia_diferente():
    ws = _wb(
        [("Rejeição", ["Fulano", "Não rejeita nenhum"], 400)],
        [("Rejeição", ["Fulano", "Não rejeito nenhum"], 400)],
    )
    achados = comparar(montar_inventario(ws))
    assert [a["tipo"] for a in achados] == ["OPCOES_DIVERGENTES"]
    assert "nao rejeito nenhum" in achados[0]["detalhe"]


def test_pergunta_so_nos_totais_e_so_nos_segmentos():
    ws = _wb(
        [("Pergunta Alfa", ["A", "B"], 400), ("Tema Totalmente Diferente Beta", ["A", "B"], 400)],
        [("Pergunta Alfa", ["A", "B"], 400), ("Xyz 123", ["A", "B"], 400)],
    )
    assert sorted(_tipos(ws)) == ["SEM_PAR_SEGMENTO", "SEM_PAR_TOTAL"]


def test_arquivo_sem_secoes_nao_gera_achado():
    wb = openpyxl.Workbook()
    ws = wb.active
    _tabela(ws, 1, "Só totais", ["A", "B"], 100)
    inv = montar_inventario(ws)
    assert inv["tem_secoes"] is False
    assert comparar(inv) == []


def test_relatorio_xlsx_tem_resumo_e_achados():
    import io

    from core.verificador_relatorio_math import relatorio_para_xlsx

    ws = _wb(
        [("Rejeição", ["Fulano", "Não rejeita nenhum"], 400)],
        [("Rejeição", ["Fulano", "Não rejeito nenhum"], 400)],
    )
    inv = montar_inventario(ws)
    achados = comparar(inv)
    wb = openpyxl.load_workbook(io.BytesIO(relatorio_para_xlsx("teste.xlsx", inv, achados)))
    assert wb.sheetnames == ["Resumo", "Achados"]
    assert wb["Achados"].max_row == 1 + len(achados)
    assert wb["Achados"]["A2"].value == "Opções divergentes"


def test_observacao_padrao_senador_e_generica():
    from core.verificador_relatorio_math import observacao

    senador = {"tipo": "TITULO_DIVERGENTE", "titulo_total": "Intenção de voto para Senador - Segundo Candidato",
               "titulo_segmento": "Intenção de voto para Senador - Base 100% (1º + 2º Voto)", "detalhe": ""}
    assert "estrutural" in observacao(senador)
    outro = {"tipo": "TITULO_DIVERGENTE", "titulo_total": "Governador", "titulo_segmento": "Governadro", "detalhe": ""}
    assert observacao(outro).startswith("Conferir")


def test_escolher_aba_ignora_sumario():
    from core.verificador_relatorio_math import escolher_aba

    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.create_sheet("Sumário")
    assert escolher_aba(wb) == ["Sheet1"]
