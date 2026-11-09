"""Perfil dos respondentes: leitura do arquivo, classificação dos envios e preenchimento do slide."""
import io
import os
import sys
import zipfile
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from pptx import Presentation

from core.esqueleto_apresentacao_math import ConfigEstudo, montar_esqueleto
from core.perfil_math import ler_perfil, achar_slide_perfil
from core.preenchimento_math import classificar, preencher, reunir_arquivos


def _xlsx(linhas):
    wb = openpyxl.Workbook()
    ws = wb.active
    for l in linhas:
        ws.append(l)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def _bloco(titulo, itens, base, media=None):
    out = [[titulo], [" ", "Total"]] + [[k, v] for k, v in itens] + [["Base", base]]
    if media is not None:
        out.append(["Média", media])
    out.append([])
    return out


SEXO = _bloco("Sexo", [("Masculino", 48.0), ("Feminino", 52.0)], 400)
IDADE = _bloco("Idade", [("16 a 19", 10.0), ("20 a 29", 20.0), ("30 a 39", 20.0), ("40 a 49", 20.0), ("50 ou mais", 30.0)], 400, 41.5)
RENDA = _bloco("Renda Familiar", [("Até 02", 35.0), ("02 e 05", 40.0), ("05 e 10", 18.0), ("Mais de 10", 7.0)], 400)
ESCOL = _bloco("Escolaridade", [("Ensino Fundamental", 20.0), ("Ensino Médio", 60.0), ("Ensino Superior", 20.0)], 396)
RELIG = _bloco("Religiões", [("Católica Total", 60.0), ("Evangélica Total", 40.0)], 400)


def _perfil(*blocos):
    return _xlsx([l for b in blocos for l in b])


def _esqueleto():
    cfg = ConfigEstudo(local="Nilópolis", uf="RJ", modo="rua", campo_inicio=date(2026, 9, 10), campo_fim=date(2026, 9, 14),
                       tempo_medio=12, base=400, margem="5%", regioes="Centro: A", n_regioes=1)
    return montar_esqueleto(None, cfg)[0]


def test_le_blocos_pelo_titulo_e_ignora_a_ordem():
    p = ler_perfil(_perfil(ESCOL, SEXO, RELIG, RENDA, IDADE))
    assert set(p["blocos"]) == {"sexo", "idade", "renda", "escolaridade", "religiao"}
    assert p["blocos"]["idade"]["media"] == 41.5
    assert p["blocos"]["renda"]["categorias"][0] == "Até 02"
    assert p["base"] == 400 and not p["avisos"] and not p["pendentes"]


def test_bloco_ausente_e_soma_errada_viram_pendencia_e_aviso():
    errado = _bloco("Renda Familiar", [("Até 02", 35.0), ("02 e 05", 40.0)], 400)   # soma 75
    p = ler_perfil(_perfil(SEXO, IDADE, errado))
    assert any("Escolaridade" in x for x in p["pendentes"])
    assert any("somam 75" in a for a in p["avisos"])


def test_zip_dentro_de_zip_e_classificacao_por_conteudo():
    interno = io.BytesIO()
    with zipfile.ZipFile(interno, "w") as z:
        z.writestr("Relatórios/Perfil_qualquer_nome.xlsx", _perfil(SEXO, IDADE, RENDA, ESCOL))
        z.writestr("__MACOSX/lixo.xlsx", b"x")
        z.writestr("Relatórios/leiame.txt", b"oi")
    externo = io.BytesIO()
    with zipfile.ZipFile(externo, "w") as z:
        z.writestr("estudo.zip", interno.getvalue())
    nomes = [n for n, _ in reunir_arquivos([("tudo.zip", externo.getvalue())])]
    assert "Perfil_qualquer_nome.xlsx" in nomes and "lixo.xlsx" not in nomes
    c = classificar([("tudo.zip", externo.getvalue()), ("esq.pptx", _esqueleto())])
    assert set(c["por_tipo"]) == {"perfil", "esqueleto"}
    assert "leiame.txt" in c["ignorados"]


def test_preenche_os_quatro_graficos_e_a_base():
    esq = _esqueleto()
    c = classificar([("Perfil.xlsx", _perfil(SEXO, IDADE, RENDA, ESCOL))])
    out, conf = preencher(esq, c["por_tipo"], base_esperada=400)
    prs = Presentation(io.BytesIO(out))
    pos, slide, graficos = achar_slide_perfil(prs)
    valores = {}
    for g, _ in graficos:
        cats = list(g.chart.plots[0].categories)
        valores[cats[0]] = list(g.chart.plots[0].series[0].values)
    assert valores["Masculino"] == [48.0, 52.0]
    assert valores["Ensino Fundamental"] == [20.0, 60.0, 20.0]            # escolaridade não pode receber idade
    assert valores["16 a 19"] == [10.0, 20.0, 20.0, 20.0, 30.0]
    assert valores["Até 02 s.m."] == [35.0, 40.0, 18.0, 7.0]
    textos = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
    assert "Base: 400 entrevistas (escolaridade: 396)" in textos
    assert "INFORMA_PENDENTE" not in [sh.name for sh in slide.shapes]
    assert not conf.divergencias
    assert any("salário mínimo" in p[1] for p in conf.pendencias)          # valor do s.m. não é inventado
    assert all(l[3] is not None for l in conf.linhas)


def test_sem_perfil_mantem_slide_com_aviso():
    out, conf = preencher(_esqueleto(), {})
    prs = Presentation(io.BytesIO(out))
    _, slide, graficos = achar_slide_perfil(prs)
    assert any("Arquivo de Perfil não enviado – conteúdo pendente" in sh.text_frame.text
               for sh in slide.shapes if sh.has_text_frame)
    assert all(list(g.chart.plots[0].series[0].values) == [None] * len(list(g.chart.plots[0].categories)) for g, _ in graficos)
    assert conf.pendencias and not conf.linhas


def test_base_diferente_do_descritivo_vira_divergencia():
    c = classificar([("Perfil.xlsx", _perfil(SEXO, IDADE, RENDA, ESCOL))])
    _, conf = preencher(_esqueleto(), c["por_tipo"], base_esperada=604)
    assert conf.divergencias and "400" in conf.divergencias[0][1]
    wb = openpyxl.load_workbook(io.BytesIO(conf.para_xlsx({"local": "x"})))
    assert wb.sheetnames == ["Resumo", "Conferência", "Divergências", "Pendências e observações"]
