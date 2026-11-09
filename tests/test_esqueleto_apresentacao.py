"""Testes do esqueleto da apresentação (usa o molde real em assets/)."""
import io
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation

from core.esqueleto_apresentacao_math import (
    ConfigEstudo, montar_esqueleto, sugerir_margem_erro, texto_periodo_campo, verificar_residuos,
)


def _cfg(**kw):
    base = dict(
        local="Nilópolis", uf="RJ", modo="rua", campo_inicio=date(2026, 9, 10), campo_fim=date(2026, 9, 14),
        tempo_medio=12, base=400, margem="5%", regioes="Centro: Bairro A\nNorte: Bairro C", n_regioes=2,
        governador="Fulano Gov", prefeito="Sicrano Pref",
    )
    base.update(kw)
    return ConfigEstudo(**base)


def _textos(pptx_bytes):
    from core.esqueleto_apresentacao_math import _paragrafos
    prs = Presentation(io.BytesIO(pptx_bytes))
    return prs, [[p.text for p in _paragrafos(s)] for s in prs.slides]


def test_margem_e_periodo():
    assert sugerir_margem_erro(400) == "5%"
    assert sugerir_margem_erro(1000) == "3%"
    assert sugerir_margem_erro(2000) == "2,19%"
    assert sugerir_margem_erro(None) == ""
    assert texto_periodo_campo(date(2026, 9, 10), date(2026, 9, 14)) == "entre os dias 10 e 14 de setembro de 2026"
    assert texto_periodo_campo(date(2026, 9, 10), date(2026, 9, 10)) == "no dia 10 de setembro de 2026"


def test_local_obrigatorio():
    try:
        montar_esqueleto(None, ConfigEstudo())
    except ValueError:
        return
    raise AssertionError("deveria exigir o local")


def test_esqueleto_basico_sem_residuos_e_manifesto_json():
    dados, man = montar_esqueleto(None, _cfg())
    assert man["residuos_de_outros_estudos"] == []
    ids = [s["id"] for s in man["slides"]]
    assert len(ids) == len(set(ids))
    assert len(man["slides"]) == len(Presentation(io.BytesIO(dados)).slides)
    json.dumps(man, ensure_ascii=False)  # serializável
    prs, textos = _textos(dados)
    plano = " ".join(" ".join(t) for t in textos)
    assert "{{" not in plano and "Paty do Alferes" not in plano and "Tarcísio" not in plano
    assert "Nilópolis" in plano and "Fulano Gov" in plano
    assert man["slides"][-1]["modulo"] == "encerramento"


def test_aviso_pendente_e_imagem_de_regioes_removida():
    dados, man = montar_esqueleto(None, _cfg())
    prs = Presentation(io.BytesIO(dados))
    pend = [s for s, m in zip(prs.slides, man["slides"]) if m["status"] == "pendente"]
    assert pend
    for s in pend:
        assert any(sh.name.startswith("INFORMA_PENDENTE") for sh in s.shapes)
    for s in prs.slides:
        for sh in s.shapes:
            assert not (sh.shape_type == 13 and sh.top == 5772139 and sh.height == 577473)


def test_online_remove_slide_de_auditoria_e_analises_desligadas():
    _, man_rua = montar_esqueleto(None, _cfg())
    _, man_on = montar_esqueleto(None, _cfg(modo="online"))
    assert len(man_on["slides"]) == len(man_rua["slides"]) - 1
    sem = _cfg(analises={"cruzamentos": False, "residuais": False, "correlacao": False,
                         "mapas_calor": False, "correspondencia": False})
    _, man_sem = montar_esqueleto(None, sem)
    assert len(man_sem["slides"]) < len(man_rua["slides"])
    assert not any(s["modulo"] == "analises_estrategicas" for s in man_sem["slides"])


def test_cargos_selecionados_definem_modulos():
    _, man = montar_esqueleto(None, _cfg(cargos=["Presidente"], governos=["federal"]))
    mods = {s["modulo"] for s in man["slides"]}
    assert "presidente" in mods and "senador" not in mods and "governador" not in mods
    assert "comparativo_governos" not in mods


def test_campos_pendentes_quando_faltam_dados():
    _, man = montar_esqueleto(None, ConfigEstudo(local="X", uf="SP"))
    for item in ("Base (nº de entrevistas)", "Margem de erro", "Tempo médio de aplicação", "Período de campo"):
        assert item in man["campos_pendentes"]


def test_cav_usa_candidato_e_cargo():
    dados, man = montar_esqueleto(None, _cfg(modulos_extras=["cav"], candidato_foco="Maria Teste",
                                              cargo_foco="Deputado Federal"))
    _, textos = _textos(dados)
    plano = " ".join(" ".join(t) for t in textos)
    assert "Maria Teste" in plano and "{{CANDIDATO}}" not in plano
    assert verificar_residuos(dados) == []


def test_nome_do_proprio_estudo_nao_e_residuo():
    _, man = montar_esqueleto(None, _cfg(governador="Desembargador Ricardo Couto"))
    assert man["residuos_de_outros_estudos"] == []


def test_serie_de_outro_estudo_limpa_e_regioes_longas_cabem():
    longa = "\n".join(f"Região {i} (10,00%) – " + ", ".join(f"Bairro {j}" for j in range(60)) for i in range(6))
    dados, man = montar_esqueleto(None, _cfg(regioes=longa, n_regioes=6))
    prs = Presentation(io.BytesIO(dados))
    import re as _re
    for s in prs.slides:
        for sh in s.shapes:
            if getattr(sh, "has_chart", False) and sh.has_chart:
                for x in sh.chart.plots[0].series:
                    assert not _re.match(r"^[a-zç]{3}/\d{2}$", x.name or "", _re.I)
    caixa = next(sh for sh in prs.slides[7].shapes if sh.name == "CaixaDeTexto 6")
    tamanhos = {r.font.size.pt for p in caixa.text_frame.paragraphs for r in p.runs if r.font.size}
    assert tamanhos and max(tamanhos) < 12


def test_sugerir_nomes():
    from core.esqueleto_apresentacao_math import sugerir_nomes
    inv = {"totais": [{"titulo": "Avaliação do Governo do Prefeito Ferreti (Estimulada e Única)"},
                      {"titulo": "Avaliação do Governador Interino - Desembargador Ricardo Couto (Estimulada e Única)"}]}
    assert sugerir_nomes(inv) == {"prefeito": "Ferreti", "governador": "Desembargador Ricardo Couto"}


# ------------------------------------------------------------------ análise das perguntas
def test_analise_das_perguntas_gera_divisoria_e_slide_modelo_pendente():
    dados, man = montar_esqueleto(None, _cfg())
    itens = [s for s in man["slides"] if s["modulo"] == "analise_perguntas"]
    assert [s["papel"] for s in itens] == ["divisoria_perguntas", "analise_perguntas"]
    assert itens[1]["status"] == "pendente" and itens[1]["fonte_esperada"] == "Análise das perguntas"
    assert itens[0]["status"] == "pronto"
    _, textos = _textos(dados)
    plano = " ".join(" ".join(t) for t in textos)
    assert "ANÁLISE DAS PERGUNTAS" in plano and "Pergunta × segmentos" in plano
    assert verificar_residuos(dados) == []
    # vem antes das análises estratégicas e depois dos cargos
    mods = [s["modulo"] for s in man["slides"]]
    assert mods.index("analise_perguntas") > mods.index("deputados")
    assert mods.index("analise_perguntas") < mods.index("analises_estrategicas")


def test_analise_das_perguntas_desmarcada_remove_slides():
    from core.esqueleto_apresentacao_math import ANALISES
    assert "analise_perguntas" in ANALISES
    base = {k: True for k in ANALISES}
    base["analise_perguntas"] = False
    _, man = montar_esqueleto(None, _cfg(analises=base))
    assert not any(s["modulo"] == "analise_perguntas" for s in man["slides"])
    _, man_on = montar_esqueleto(None, _cfg())
    assert len(man_on["slides"]) == len(man["slides"]) + 2


def test_config_antiga_sem_a_nova_analise_assume_marcada():
    cfg = ConfigEstudo.de_dict({"local": "X", "uf": "SP", "analises": {"cruzamentos": False}})
    assert cfg.analises["analise_perguntas"] is True and cfg.analises["cruzamentos"] is False


# ------------------------------------------------------------------ objetivos pré-definidos
def test_objetivos_sugeridos_a_partir_do_descritivo_e_da_configuracao():
    from core.esqueleto_apresentacao_math import compor_objetivo, sugerir_objetivos
    inv = {"totais": [
        {"titulo": "Intenção de voto para Senador (Estimulada e Única)", "pergunta": "Qual o primeiro e o segundo voto para senador?"},
        {"titulo": "Rejeição aos candidatos a Deputado Estadual", "pergunta": ""},
        {"titulo": "Transferência de voto sem Fulano (Deputado Estadual)", "pergunta": ""},
        {"titulo": "Melhor ação da gestão do Prefeito", "pergunta": ""},
    ], "segmentos": []}
    ids = {o["id"] for o in sugerir_objetivos(inv, _cfg())}
    assert {"transferencia_votos", "rejeicao", "gestao_acoes", "gestao_municipal", "analise_perguntas", "associacoes"} <= ids
    assert "senado_dois_votos" in ids
    assert "candidato_foco" not in ids and "prefeito_2028" not in ids
    sug = sugerir_objetivos(inv, _cfg(prefeito="Ferreti"))
    assert any("Ferreti" in o["texto"] for o in sug if o["id"] == "gestao_municipal")
    assert all(o["origem"] and o["texto"].endswith(".") for o in sug)
    txt = compor_objetivo([sug[0]["texto"], "", sug[1]["texto"]])
    assert txt.count(". ") == 1 and txt.startswith(sug[0]["texto"])


def test_objetivos_nao_inventam_sem_base_nos_arquivos():
    from core.esqueleto_apresentacao_math import sugerir_objetivos
    cfg = _cfg(analises={k: False for k in ("cruzamentos", "residuais", "correlacao", "mapas_calor", "correspondencia", "analise_perguntas")},
               governos=["federal"], cargos=["Presidente"])
    assert sugerir_objetivos(None, cfg) == []


def test_objetivo_candidato_foco_e_prefeito_2028():
    from core.esqueleto_apresentacao_math import sugerir_objetivos
    cfg = _cfg(modulos_extras=["cav"], candidato_foco="Maria Teste", cargo_foco="Deputado Federal", cargos=["Prefeito"])
    sug = {o["id"]: o for o in sugerir_objetivos(None, cfg)}
    assert "Maria Teste" in sug["candidato_foco"]["texto"] and "Deputado Federal" in sug["candidato_foco"]["texto"]
    assert "2028" in sug["prefeito_2028"]["texto"]
    sem_nome = {o["id"]: o for o in sugerir_objetivos(None, _cfg(modulos_extras=["cav"]))}
    assert "(pendente)" in sem_nome["candidato_foco"]["texto"]


def test_objetivo_escolhido_vai_para_o_slide_de_objetivos_com_fonte_ajustada():
    longo = " ".join(["Identificar diferenças de percepção entre os segmentos em cada pergunta do questionário."] * 8)
    dados, _ = montar_esqueleto(None, _cfg(objetivo_especifico=longo))
    prs = Presentation(io.BytesIO(dados))
    caixa = next(sh for sh in prs.slides[5].shapes if sh.name == "CaixaDeTexto 7")
    assert "Identificar diferenças de percepção" in caixa.text_frame.text
    tam = {r.font.size.pt for p in caixa.text_frame.paragraphs for r in p.runs if r.font.size}
    assert tam and max(tam) < 14
    curto, _ = montar_esqueleto(None, _cfg(objetivo_especifico="Avaliar X."))
    caixa = next(sh for sh in Presentation(io.BytesIO(curto)).slides[5].shapes if sh.name == "CaixaDeTexto 7")
    assert max(r.font.size.pt for p in caixa.text_frame.paragraphs for r in p.runs if r.font.size) == 14


def _png_bytes(w=300, h=200):
    from PIL import Image
    b = io.BytesIO()
    Image.new("RGB", (w, h), "red").save(b, "PNG")
    return b.getvalue()


def _slide8(dados, manifesto):
    prs = Presentation(io.BytesIO(dados))
    pos = next((i for i, m in enumerate(manifesto["slides"]) if m["slide_no_molde"] == 8), None)
    assert pos is not None
    return prs.slides[pos]


def test_mapa_da_capa_entra_no_slide_8():
    dados, man = montar_esqueleto(None, _cfg(), None, None, mapa_regioes=_png_bytes())
    sl8 = _slide8(dados, man)
    nomes = [s.name for s in sl8.shapes]
    assert "Mapa das regiões" in nomes
    assert not any(n.startswith("INFORMA_PENDENTE") and "mapa" in (s.text_frame.text.lower() if s.has_text_frame else "")
                   for n, s in zip(nomes, sl8.shapes))
    pic = next(s for s in sl8.shapes if s.name == "Mapa das regiões")
    assert abs(pic.width / pic.height - 1.5) < 0.02          # proporção preservada
    assert "Imagem do mapa das regiões" not in man["campos_pendentes"]


def test_sem_mapa_mantem_aviso_pendente():
    dados, man = montar_esqueleto(None, _cfg())
    sl8 = _slide8(dados, man)
    assert "Mapa das regiões" not in [s.name for s in sl8.shapes]
    assert any("IMAGEM PENDENTE" in s.text_frame.text for s in sl8.shapes if s.has_text_frame)
    assert "Imagem do mapa das regiões" in man["campos_pendentes"]


def test_mapa_invalido_vira_pendente():
    dados, man = montar_esqueleto(None, _cfg(), None, None, mapa_regioes=b"lixo")
    assert "Imagem do mapa das regiões" in man["campos_pendentes"]


def _slides_mapa_calor(dados):
    prs = Presentation(io.BytesIO(dados))
    return [s for s in prs.slides if any(sh.name.startswith("INFORMA_PENDENTE_REGIOES") or sh.name == "Legenda das regiões"
                                         for sh in s.shapes)]


def test_legenda_das_regioes_nos_mapas_de_calor():
    regioes = "Região 1 Centro (60,00%): Bairro A, Bairro B\nRegião 2 Norte (40,00%): Bairro C"
    dados, man = montar_esqueleto(None, _cfg(regioes=regioes))
    sl = _slides_mapa_calor(dados)
    assert len(sl) >= 3
    for s in sl:
        nomes = [sh.name for sh in s.shapes]
        assert "Legenda das regiões" in nomes and not any(n.startswith("INFORMA_PENDENTE_REGIOES") for n in nomes)
        caixa = next(sh for sh in s.shapes if sh.name == "Legenda das regiões")
        assert "Bairro C" in caixa.text_frame.text and "Região 1 Centro (60,00%)" in caixa.text_frame.text
        assert not any(sh.shape_type == 13 and sh.top == 5772139 for sh in s.shapes)   # imagem do molde removida
    assert "Legenda das regiões (slides de mapa de calor)" not in man["campos_pendentes"]


def test_legenda_pendente_sem_regioes():
    dados, man = montar_esqueleto(None, _cfg(regioes=""))
    assert all(any(sh.name.startswith("INFORMA_PENDENTE_REGIOES") for sh in s.shapes) for s in _slides_mapa_calor(dados))
    assert "Legenda das regiões (slides de mapa de calor)" in man["campos_pendentes"]


def test_legenda_longa_reduz_fonte_e_avisa():
    from core.esqueleto_apresentacao_math import _fonte_legenda
    curta, coube = _fonte_legenda(["Região 1 (50%): A, B", "Região 2 (50%): C"], 11100000, 590000)
    assert coube and curta == 8
    longa, coube2 = _fonte_legenda(["Região 1 (50%): " + "Bairro Grande, " * 400], 11100000, 590000)
    assert not coube2 and longa == 5.5
