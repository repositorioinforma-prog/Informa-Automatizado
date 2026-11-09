"""Leitura da capa do Relatório Descritivo (caixas de texto + imagem do mapa) e uso no esqueleto."""
import io
import zipfile
from datetime import date
from xml.sax.saxutils import escape

from PIL import Image

from core.capa_descritivo_math import ler_capa, parse_campo, parse_regioes, regioes_para_texto

_XDR = ('xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def _caixa(linha, paragrafos, col=0, partes=False):
    ps = ""
    for p in paragrafos:
        if partes:  # texto quebrado em vários runs (acontece no Excel)
            meio = len(p) // 2
            ps += f"<a:p><a:r><a:t>{escape(p[:meio])}</a:t></a:r><a:r><a:t>{escape(p[meio:])}</a:t></a:r></a:p>"
        else:
            ps += f"<a:p><a:r><a:t>{escape(p)}</a:t></a:r></a:p>"
    return (f'<xdr:twoCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:row>{linha}</xdr:row></xdr:from>'
            f'<xdr:sp><xdr:txBody>{ps}</xdr:txBody></xdr:sp></xdr:twoCellAnchor>')


def _pic(linha, rid):
    return (f'<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>{linha}</xdr:row></xdr:from>'
            f'<xdr:pic><xdr:blipFill><a:blip r:embed="{rid}"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>')


def _png(cor):
    b = io.BytesIO()
    Image.new("RGB", (200, 100), cor).save(b, "PNG")
    return b.getvalue()


def _xlsx(ancoras: str, imagens: dict) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("xl/drawings/drawing1.xml", f'<?xml version="1.0"?><xdr:wsDr {_XDR}>{ancoras}</xdr:wsDr>')
        rels = "".join(f'<Relationship Id="{rid}" Type="x" Target="../media/{n}"/>' for rid, n in
                       zip(imagens, [f"image{i + 1}.png" for i in range(len(imagens))]))
        z.writestr("xl/drawings/_rels/drawing1.xml.rels",
                   f'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        for i, dados in enumerate(imagens.values()):
            z.writestr(f"xl/media/image{i + 1}.png", dados)
    return b.getvalue()


NOTA = ["NOTA METODOLÓGICA",
        "As entrevistas foram realizadas entre os dias 19 e 24 de setembro de 2026. O tempo médio de aplicação do "
        "questionário foi de cerca de 10 minutos. Pessoas com 16 anos ou mais.",
        "Foram realizadas 1000 entrevistas."]


def _capa_rua(partes=False):
    reg = ("Região 1 Mambucaba (27,28%) – Bairro A, Bairro B; Região 2 Centro (72,72%) – Bairro C e Bairro D.")
    anc = (_caixa(2, ["ANGRA DOS REIS - RIO DE JANEIRO", "Diagnóstico Político e Eleitoral - Pesquisa Quantitativa",
                      "Setembro de 2026"])
           + _caixa(14, NOTA, partes=partes)
           + _caixa(35, ["ABRANGÊNCIA GEOGRÁFICA", "A cidade foi dividida em 02 regiões."])
           + _caixa(39, [reg], col=4, partes=partes) + _pic(39, "rId1")
           + _caixa(62, ["TECNOLOGIA AVANÇADA"]) + _pic(69, "rId2"))
    return _xlsx(anc, {"rId1": _png("red"), "rId2": _png("blue")})


def test_capa_rua_completa():
    c = ler_capa(_capa_rua())
    assert c["achou"] and c["local"] == "Angra dos Reis" and c["uf"] == "RJ"
    assert c["mes_ano"] == "Setembro de 2026"
    assert c["campo"] == (date(2026, 9, 19), date(2026, 9, 24))
    assert c["tempo_medio"] == 10 and c["idade_minima"] == 16 and c["n_regioes"] == 2
    assert c["regioes_texto"].splitlines() == [
        "Região 1 Mambucaba (27,28%): Bairro A, Bairro B",
        "Região 2 Centro (72,72%): Bairro C e Bairro D"]
    # só o mapa da seção de abrangência, não o ícone da seção de tecnologia
    assert len(c["mapas"]) == 1 and c["mapas"][0]["linha"] == 39
    assert Image.open(io.BytesIO(c["mapas"][0]["bytes"])).getpixel((0, 0)) == (255, 0, 0)


def test_capa_com_runs_quebrados():
    c = ler_capa(_capa_rua(partes=True))
    assert len(c["regioes"]) == 2 and c["campo"] == (date(2026, 9, 19), date(2026, 9, 24))


def test_capa_sem_mapa_nem_regioes_nao_inventa():
    anc = _caixa(2, ["OSASCO - SÃO PAULO", "x", "Agosto de 2026"]) + _caixa(14, NOTA)
    c = ler_capa(_xlsx(anc, {}))
    assert c["regioes"] == [] and c["mapas"] == [] and c["regioes_texto"] == ""
    assert any("regiões" in a.lower() for a in c["avisos"]) and any("mapa" in a.lower() for a in c["avisos"])


def test_capa_estadual_online_com_segunda_lista():
    reg = "1 - Baixada (22,09%): A, B; 2 - Leste (11,41%): C; 6 - Capital (66,50%): Rio de Janeiro."
    sec = "1 - Sul: Flamengo, Glória;2 - Central: Centro, Caju"
    anc = (_caixa(2, ["ESTADUAL RIO DE JANEIRO", "Diagnóstico do Senado - Pesquisa Quantitativa", "Setembro de 2026"])
           + _caixa(14, ["NOTA METODOLÓGICA", "Entrevistas em ambiente digital. Foram realizadas entre os dias 17 e 24 de "
                         "Setembro de 2026. Tempo médio de cerca de 5 minutos. Maiores de 18 anos ou mais."])
           + _caixa(32, ["ABRANGÊNCIA GEOGRÁFICA", "moradores das 3 regiões, contemplando respondentes de 60 cidades."])
           + _caixa(38, [reg], col=4) + _pic(40, "rId1")
           + _caixa(58, ["Regiões da Capital"]) + _pic(61, "rId2") + _caixa(70, [sec]))
    c = ler_capa(_xlsx(anc, {"rId1": _png("red"), "rId2": _png("green")}))
    assert c["local"] is None and c["uf"] == "RJ" and c["modo"] == "online" and c["n_cidades"] == 60
    assert len(c["regioes"]) == 3 and c["n_regioes"] == 3
    assert len(c["regioes_secundarias"]) == 2 and c["regioes_secundarias_titulo"] == "Regiões da Capital"
    assert len(c["mapas"]) == 2


def test_aviso_quando_n_regioes_diverge():
    anc = (_caixa(2, ["X - RIO DE JANEIRO"]) + _caixa(35, ["ABRANGÊNCIA GEOGRÁFICA", "dividida em 05 regiões"])
           + _caixa(39, ["Região 1 (50,00%): A; Região 2 (50,00%): B"], col=4))
    c = ler_capa(_xlsx(anc, {}))
    assert any("5 regiões" in a for a in c["avisos"])


def test_parsers_formatos_de_regiao():
    assert [r["rotulo"] for r in parse_regioes("Região 1 [20,60%]: A, B; Região 2 [7,30%]: C;")] == [
        "Região 1 (20,60%)", "Região 2 (7,30%)"]
    txt = regioes_para_texto(parse_regioes("Região 4 (14,50%): Olinda, Da Mina. *Área sem dados"))
    assert txt == "Região 4 (14,50%): Olinda, Da Mina. *Área sem dados"


def test_parse_campo_variantes():
    assert parse_campo("realizadas no dia 26 de setembro de 2026.") == (date(2026, 9, 26), date(2026, 9, 26))
    assert parse_campo("entre os dias 28 de agosto e 2 de setembro de 2026") == (date(2026, 8, 28), date(2026, 9, 2))
    assert parse_campo("sem data") is None


def test_arquivo_invalido():
    c = ler_capa(b"nao e zip")
    assert not c["achou"] and c["avisos"]


def test_legenda_na_pagina_seguinte_sem_pesos():
    anc = (_caixa(2, ["X - RIO DE JANEIRO"]) + _caixa(35, ["ABRANGÊNCIA GEOGRÁFICA", "dividida em 02 regiões"])
           + _pic(39, "rId1") + _caixa(80, ["Região 1: Bairro A, Bairro B; Região 2: Bairro C"], col=4))
    c = ler_capa(_xlsx(anc, {"rId1": _png("red")}))
    assert c["regioes_texto"].splitlines() == ["Região 1: Bairro A, Bairro B", "Região 2: Bairro C"]
    assert len(c["mapas"]) == 1 and c["n_regioes"] == 2
