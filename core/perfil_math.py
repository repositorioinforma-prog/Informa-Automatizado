"""
Perfil dos respondentes (lógica pura, sem Streamlit).

Lê o arquivo «Perfil_<local>_<UF>_<mês>.xls(x)» (saída padronizada do SPSS) e preenche o
slide «PERFIL DOS RESPONDENTES» (4 gráficos: Gênero, Idade, Renda, Escolaridade).

Formato do arquivo (primeira aba), repetido para cada variável:

    Sexo                      <- título do bloco
    (vazio) | Total
    Masculino | 48,74         <- categorias com o % do total (ponderado)
    Feminino  | 51,26
    Base      | 604
    (Média    | 42,4)         <- só em Idade

Blocos conhecidos: Sexo, Idade, Renda Familiar, Escolaridade, Raça, Religiões. Os blocos são
identificados pelo TÍTULO (nunca pela posição da linha), então a ordem pode mudar.

Regras do projeto: nunca inventar valor; sem zero como preenchimento; o que não vier no
arquivo fica pendente e entra na conferência.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from pptx.chart.data import CategoryChartData

from core.planilha_leitura import linhas_da_planilha
from core.verificador_relatorio_math import chave

# título normalizado -> id do bloco
_BLOCOS = (
    ("sexo", ("sexo", "genero")),
    ("idade", ("idade", "faixa etaria")),
    ("renda", ("renda",)),
    ("escolaridade", ("escolaridade", "instrucao")),
    ("raca", ("raca", "cor")),
    ("religiao", ("religiao", "religioes")),
)
ROTULO_BLOCO = {"sexo": "Gênero", "idade": "Idade", "renda": "Renda", "escolaridade": "Escolaridade",
                "raca": "Raça", "religiao": "Religiões"}
OBRIGATORIOS = ("sexo", "idade", "renda", "escolaridade")   # os 4 gráficos do slide
TOLERANCIA_SOMA = 1.0   # pontos percentuais: a soma de cada bloco deve ficar em 100 ± 1 (arredondamentos)


def _id_bloco(titulo: str) -> Optional[str]:
    k = chave(titulo)
    for bid, termos in _BLOCOS:
        if any(k == t or k.startswith(t + " ") for t in termos):
            return bid
    return None


def _num(v: Any) -> Optional[float]:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace("%", "").replace(".", "").replace(",", ".")) if "," in str(v) \
            else float(str(v).strip().replace("%", ""))
    except ValueError:
        return None


def _texto(v: Any) -> str:
    return " ".join(str(v).split()) if isinstance(v, str) and v.strip() else ""


def eh_perfil(linhas: List[List[Any]]) -> bool:
    """O arquivo tem cara de Perfil? (bloco «Sexo» + linha «Total» + «Base»)."""
    return bool(_ler_blocos(linhas).get("sexo"))


def _ler_blocos(linhas: List[List[Any]]) -> Dict[str, dict]:
    blocos: Dict[str, dict] = {}
    n = len(linhas)
    i = 0
    while i < n:
        linha = linhas[i] or []
        titulo = _texto(linha[0]) if linha else ""
        resto_vazio = all(v in (None, "") or (isinstance(v, str) and not v.strip()) for v in linha[1:])
        bid = _id_bloco(titulo) if titulo and resto_vazio else None
        # o título é seguido pela linha de cabeçalho «Total»
        j = i + 1
        while j < n and all(v in (None, "") or (isinstance(v, str) and not v.strip()) for v in (linhas[j] or [])):
            j += 1
        cab = [chave(v) for v in (linhas[j] or [])] if j < n else []
        if bid and "total" in cab:
            col_total = cab.index("total")
            bloco = {"id": bid, "titulo": titulo, "categorias": [], "valores": [], "base": None, "media": None,
                     "linha": i + 1}
            k = j + 1
            while k < n:
                lk = linhas[k] or []
                rot = _texto(lk[0]) if lk else ""
                val = lk[col_total] if len(lk) > col_total else None
                if not rot:
                    if bloco["base"] is not None or bloco["categorias"]:
                        break          # linha em branco depois dos dados: fim do bloco
                    k += 1
                    continue
                rk = chave(rot)
                if rk == "base":
                    bloco["base"] = _num(val)
                elif rk in ("media", "mediana"):
                    bloco[rk] = _num(val)
                elif _num(val) is not None:
                    bloco["categorias"].append(rot)
                    bloco["valores"].append(_num(val))
                else:
                    break              # começou outro bloco
                k += 1
            if bloco["categorias"]:
                blocos.setdefault(bid, bloco)
            i = k
            continue
        i += 1
    return blocos


def ler_perfil(conteudo: bytes) -> dict:
    """
    Lê o arquivo de Perfil. Retorna {"blocos": {id: {...}}, "base": N, "avisos": [...], "pendentes": [...]}.
    Levanta ErroLeitura se o arquivo não abrir; avisos/pendentes descrevem o que veio incompleto.
    """
    blocos = _ler_blocos(linhas_da_planilha(conteudo))
    avisos: List[str] = []
    pendentes: List[str] = []
    for bid in OBRIGATORIOS:
        if bid not in blocos:
            pendentes.append(f"Perfil: bloco «{ROTULO_BLOCO[bid]}» não encontrado no arquivo")
    for bid, b in blocos.items():
        soma = sum(b["valores"])
        if abs(soma - 100) > TOLERANCIA_SOMA:
            avisos.append(f"Perfil · {ROTULO_BLOCO.get(bid, b['titulo'])}: as categorias somam {soma:.1f}% (esperado ≈ 100%).")
        if b["base"] is None:
            avisos.append(f"Perfil · {ROTULO_BLOCO.get(bid, b['titulo'])}: linha «Base» não encontrada.")
    base = (blocos.get("sexo") or {}).get("base") or next((b["base"] for b in blocos.values() if b["base"]), None)
    return {"blocos": blocos, "base": int(base) if base else None, "avisos": avisos, "pendentes": pendentes}


# ---------------------------------------------------------------- preenchimento do slide
def _walk_pai(shapes, pai=None):
    for sh in shapes:
        yield sh, pai
        if sh.shape_type == 6:
            yield from _walk_pai(sh.shapes, sh)


def _textos(grupo) -> str:
    if grupo is None:
        return ""
    from_ = []
    for sh in grupo.shapes:
        if sh.has_text_frame:
            from_.append(sh.text_frame.text)
        if sh.shape_type == 6:
            from_.append(_textos(sh))
    return " ".join(from_)


def _bloco_do_grafico(grafico_shape, pai) -> Optional[str]:
    """Qual variável o gráfico mostra? Pelo rótulo do grupo («IDADE», «Gênero»...); senão, pelas categorias."""
    rot = chave(_textos(pai))
    palavras = rot.split()
    for bid, termos in _BLOCOS:   # palavra inteira: «idade» não pode casar dentro de «escolaridade»
        if bid in OBRIGATORIOS and any((t in rot) if " " in t else (t in palavras) for t in termos):
            return bid
    cats = [chave(c) for c in grafico_shape.chart.plots[0].categories]
    if any(c.startswith(("masc", "fem")) for c in cats):
        return "sexo"
    if any(c.startswith("ensino") for c in cats):
        return "escolaridade"
    if any("s m" in c or c.startswith("ate ") or re.match(r"^\d+ e \d+", c) for c in cats):
        return "renda"
    if any(re.match(r"^\d+ a \d+", c) or "ou mais" in c for c in cats):
        return "idade"
    return None


def _destacar_maior(chart, valores: List[float]) -> None:
    """Barra de maior valor em azul-marinho (002060); as demais ficam com a cor da série. Remove os
    destaques que vinham do molde (pontos fixos de outro estudo)."""
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    serie = chart.plots[0].series[0]
    for dpt in serie._element.findall(qn("c:dPt")):
        serie._element.remove(dpt)
    if not valores:
        return
    imax = valores.index(max(valores))
    for i in range(len(valores)):      # todos os pontos explícitos: independe do preenchimento da série no molde
        fmt = serie.points[i].format.fill
        fmt.solid()
        if i == imax:
            fmt.fore_color.rgb = RGBColor(0x00, 0x20, 0x60)
        else:
            fmt.fore_color.rgb = RGBColor(0xAD, 0xB9, 0xCA)         # cinza-azulado do molde


def _rotulos(bid: str, categorias: List[str]) -> List[str]:
    if bid == "renda":   # no slide a renda é em salários mínimos
        return [c if "s.m" in c.lower() else f"{c} s.m." for c in categorias]
    return list(categorias)


def achar_slide_perfil(prs):
    """Slide «PERFIL DOS RESPONDENTES» com os gráficos (None se o molde não tiver)."""
    for pos, slide in enumerate(prs.slides, 1):
        graficos = [(sh, pai) for sh, pai in _walk_pai(slide.shapes)
                    if getattr(sh, "has_chart", False) and sh.has_chart]
        if len(graficos) >= 3 and any("perfil" in chave(sh.text_frame.text)
                                      for sh, _ in _walk_pai(slide.shapes) if sh.has_text_frame):
            return pos, slide, graficos
    return None


def preencher_slide_perfil(prs, perfil: Optional[dict], conf, nome_arquivo: str = "Perfil") -> bool:
    """
    Preenche o slide de perfil com os dados lidos (dict de `ler_perfil`). `conf` é a Conferencia
    que recebe linhas/pendências/divergências. Sem `perfil` (arquivo não enviado): mantém o slide
    e troca o aviso por «Arquivo de Perfil não enviado – conteúdo pendente».
    """
    achado = achar_slide_perfil(prs)
    if not achado:
        conf.aviso("Slide de perfil não encontrado no esqueleto (módulo desligado?).")
        return False
    pos, slide, graficos = achado
    from core.esqueleto_apresentacao_math import ID_PENDENTE, _forma, _marcar_pendente, _remover_forma

    if not perfil:
        aviso = _forma(slide, ID_PENDENTE)
        if aviso is not None:
            _remover_forma(aviso)
        _marcar_pendente(slide, "Arquivo de Perfil não enviado – conteúdo pendente")
        conf.pendencia(pos, "Perfil dos respondentes", "Arquivo de Perfil não enviado")
        return False

    blocos = perfil["blocos"]
    for g, pai in graficos:
        bid = _bloco_do_grafico(g, pai)
        if bid is None:
            conf.aviso(f"Slide {pos}: não identifiquei qual variável o gráfico «{g.name}» mostra.")
            continue
        bloco = blocos.get(bid)
        if not bloco:
            conf.pendencia(pos, f"Perfil · {ROTULO_BLOCO[bid]}", f"Bloco não encontrado em {nome_arquivo}")
            continue
        dados = CategoryChartData(number_format="0.0")
        rotulos = _rotulos(bid, bloco["categorias"])
        dados.categories = rotulos
        dados.add_series("%", bloco["valores"])
        g.chart.replace_data(dados)
        _destacar_maior(g.chart, bloco["valores"])
        for rot, val in zip(rotulos, bloco["valores"]):
            conf.linha(pos, f"Perfil · {ROTULO_BLOCO[bid]} · {rot}", nome_arquivo, round(val, 2))
        conf.linha(pos, f"Perfil · {ROTULO_BLOCO[bid]} · Base", nome_arquivo, bloco["base"])
        if bloco.get("media") is not None and bid == "idade":
            conf.linha(pos, "Perfil · Idade · Média (não aparece no slide)", nome_arquivo, round(bloco["media"], 2))

    # base: a do bloco Sexo; se a escolaridade tiver base menor, avisa no rodapé do slide
    base = perfil.get("base")
    if base:
        esc = (blocos.get("escolaridade") or {}).get("base")
        _escrever_base(slide, base, int(esc) if esc and int(esc) != base else None)
        conf.linha(pos, "Perfil · Base do slide", nome_arquivo, base)

    # nota do salário mínimo: o valor não vem do Perfil, então não é inventado
    nota = _forma(slide, "Retângulo 12")
    if nota is not None and nota.has_text_frame and "s.m." in nota.text_frame.text:
        from core.esqueleto_apresentacao_math import _reescrever_paragrafo
        _reescrever_paragrafo(nota.text_frame.paragraphs[0], [("s.m.: salário mínimo (valor a confirmar)", False)])
        conf.pendencia(pos, "Perfil · nota do salário mínimo", "Valor do salário mínimo vigente não vem do Perfil — confirmar")

    aviso = _forma(slide, ID_PENDENTE)
    if aviso is not None:
        _remover_forma(aviso)
    for a in perfil.get("avisos", []):
        conf.aviso(a)
    for p in perfil.get("pendentes", []):
        conf.pendencia(pos, "Perfil dos respondentes", p)
    return True


def _escrever_base(slide, base: int, base_escolaridade: Optional[int]) -> None:
    """Rodapé «Base: N entrevistas» (com «(escolaridade: M)» quando a base de escolaridade for menor)."""
    from core.esqueleto_apresentacao_math import _paragrafos, _reescrever_paragrafo
    texto = f"Base: {base} entrevistas" + (f" (escolaridade: {base_escolaridade})" if base_escolaridade else "")
    for par in _paragrafos(slide):
        t = par.text.strip()
        if t.startswith("Base:") and "entrevistas" in t and par.runs:
            _reescrever_paragrafo(par, [(texto, False)])
