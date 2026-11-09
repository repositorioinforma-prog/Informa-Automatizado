"""Lê a capa do Relatório Descritivo (xlsx): título, nota metodológica, regiões/bairros e imagem do mapa.

A capa do Descritivo é feita de caixas de texto e imagens ancoradas na planilha (não de células); por isso a leitura
é direta no XML do xlsx (zip), sem depender do openpyxl. Nada é inventado: o que não for encontrado volta como None/vazio.
"""
from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from datetime import date
from posixpath import normpath, dirname, join

from lxml import etree

_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"

MESES = {m: i + 1 for i, m in enumerate(
    ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"])}
_NOME_MES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

_ROTULO_REGIAO = re.compile(r"(regi[aã]o\s+)?\d+\s*[-–:.]?", re.I)
_PCT = re.compile(r"[\(\[]\s*(\d+(?:[.,]\d+)?)\s*%\s*[\)\]]")


def _sem_acento(t: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn").lower()


def _compacta(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "").replace(" ", " ")).strip()


# ------------------------------------------------------------------ leitura do xlsx
def _ancoras(zf: zipfile.ZipFile):
    """Itera (linha, coluna, parágrafos, imagens) de cada âncora do primeiro desenho que tenha a capa."""
    desenhos = sorted(n for n in zf.namelist() if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n))
    for nome in desenhos:
        raiz = etree.fromstring(zf.read(nome))
        rels = {}
        rel_nome = join(dirname(nome), "_rels", nome.rsplit("/", 1)[1] + ".rels")
        if rel_nome in zf.namelist():
            for r in etree.fromstring(zf.read(rel_nome)).iter(_REL):
                rels[r.get("Id")] = normpath(join(dirname(nome), r.get("Target")))
        itens = []
        for anc in raiz:
            fr = anc.find("xdr:from", _NS)
            linha = int(fr.find("xdr:row", _NS).text) if fr is not None else 0
            col = int(fr.find("xdr:col", _NS).text) if fr is not None else 0
            paras = []
            for p in anc.iter("{%s}p" % _NS["a"]):
                t = "".join(x.text or "" for x in p.iter("{%s}t" % _NS["a"]))
                if t.strip():
                    paras.append(t)
            imgs = []
            for pic in anc.iter("{%s}pic" % _NS["xdr"]):
                blip = pic.find(".//a:blip", _NS)
                rid = blip.get("{%s}embed" % _NS["r"]) if blip is not None else None
                alvo = rels.get(rid)
                if alvo and alvo in zf.namelist():
                    imgs.append((alvo, zf.read(alvo)))
            itens.append({"linha": linha, "col": col, "paragrafos": paras, "imagens": imgs})
        txt = " ".join(" ".join(i["paragrafos"]) for i in itens)
        if "NOTA METODOL" in txt.upper() or "ABRANG" in txt.upper():
            return itens
    return []


# ------------------------------------------------------------------ interpretação dos textos
def parse_regioes(texto: str) -> list[dict]:
    """Quebra o texto das regiões em [{'rotulo': 'Região 1 Mambucaba (27,28%)', 'peso': '27,28', 'bairros': '...'}].

    Aceita «Região 1 (23,20%): a, b;», «Região 1 [20,60%]: …», «Região 1 Mambucaba (27,28%) – …» e «1 - Nome (22,09%): …».
    """
    texto = _compacta(texto)
    saida = []
    for item in re.split(r";\s*", texto):
        item = item.strip()
        if not item:
            continue
        m = _PCT.search(item)
        if m:
            cabeca = item[:m.start()].strip()
            resto = item[m.end():]
            peso = m.group(1)
            rotulo = f"{cabeca} ({peso}%)"
        else:
            sep = re.search(r"\s[–-]\s|:", item)
            if not sep:
                continue
            rotulo, resto, peso = item[:sep.start()].strip(), item[sep.end():], None
        bairros = re.sub(r"^[\s:–-]+", "", resto).strip().rstrip(".; ").strip()
        if not rotulo or not bairros:
            continue
        saida.append({"rotulo": rotulo.strip(), "peso": peso, "bairros": bairros})
    return saida


def regioes_para_texto(regioes: list[dict]) -> str:
    """Formato do campo do app: uma região por linha — «Região 1 (28,49%): Bairro A, Bairro B»."""
    return "\n".join(f"{r['rotulo']}: {r['bairros']}" for r in regioes)


def parse_campo(texto: str) -> tuple[date, date] | None:
    """Período de campo da nota metodológica ('entre os dias 19 e 24 de setembro de 2026', 'no dia 26 de …')."""
    t = _sem_acento(_compacta(texto))
    mes = "(" + "|".join(MESES) + ")"
    m = re.search(rf"entre os dias (\d{{1,2}}) de {mes} e (\d{{1,2}}) de {mes} de (\d{{4}})", t)
    if m:
        a, ma, b, mb, ano = int(m.group(1)), MESES[m.group(2)], int(m.group(3)), MESES[m.group(4)], int(m.group(5))
        return date(ano, ma, a), date(ano, mb, b)
    m = re.search(rf"entre os dias (\d{{1,2}}) e (\d{{1,2}}) de {mes} de (\d{{4}})", t)
    if m:
        a, b, mm, ano = int(m.group(1)), int(m.group(2)), MESES[m.group(3)], int(m.group(4))
        return date(ano, mm, a), date(ano, mm, b)
    m = re.search(rf"no dia (\d{{1,2}}) de {mes} de (\d{{4}})", t)
    if m:
        d = date(int(m.group(3)), MESES[m.group(2)], int(m.group(1)))
        return d, d
    return None


def parse_nota(texto: str) -> dict:
    t = _compacta(texto)
    s = _sem_acento(t)
    out = {"campo": parse_campo(t), "tempo_medio": None, "idade_minima": None, "modo": None, "n_regioes": None, "n_cidades": None}
    m = re.search(r"cerca de (\d+) minuto", s)
    if m:
        out["tempo_medio"] = int(m.group(1))
    m = re.search(r"(\d+) anos ou mais", s)
    if m:
        out["idade_minima"] = int(m.group(1))
    if "ambiente digital" in s or "redes sociais" in s or "autopreenchimento" in s:
        out["modo"] = "online"
    elif "presencia" in s or "pessoais" in s:
        out["modo"] = "rua"
    return out


def _estado_para_uf(nome: str, estados: dict | None) -> str | None:
    if not estados:
        return None
    alvo = _sem_acento(nome)
    for uf, n in estados.items():
        if _sem_acento(n) == alvo or _sem_acento(uf) == alvo:
            return uf
    return None


def parse_titulo(paragrafos: list[str], estados: dict | None = None) -> dict:
    """Primeiro parágrafo: 'ANGRA DOS REIS - RIO DE JANEIRO'. Retorna local/estado/UF, mês-ano e o subtítulo."""
    out = {"local": None, "estado": None, "uf": None, "mes_ano": None, "subtitulo": None}
    if not paragrafos:
        return out
    linhas = [_compacta(p) for p in paragrafos]
    if estados is None:
        from core.esqueleto_apresentacao_math import ESTADOS as estados  # noqa: N811
    me = re.fullmatch(r"ESTADUAL\s+(.+)", linhas[0], re.I)
    if me:  # estudo estadual: não há município, só o estado
        out["estado"] = me.group(1).strip().title().replace(" De ", " de ")
        out["uf"] = _estado_para_uf(me.group(1).strip(), estados)
        m = None
    else:
        m = re.fullmatch(r"(.+?)\s+[-–]\s+(.+)", linhas[0])
    if m:
        local, estado = m.group(1).strip(), m.group(2).strip()
        uf = _estado_para_uf(estado, estados)
        if uf:
            out["local"], out["estado"], out["uf"] = local.title().replace(" De ", " de ").replace(" Dos ", " dos ").replace(" Da ", " da ").replace(" Do ", " do "), estado, uf
    resto = " - ".join(linhas[1:])
    mm = re.search(r"\s*-?\s*(" + "|".join(_NOME_MES) + r")\s+de\s+(\d{4})\s*$", resto, re.I)
    if mm:
        out["mes_ano"] = f"{mm.group(1).capitalize()} de {mm.group(2)}"
        resto = resto[:mm.start()]
    out["subtitulo"] = resto.strip(" -") or None
    return out


# ------------------------------------------------------------------ função principal
def ler_capa(conteudo: bytes, estados: dict | None = None) -> dict:
    """Extrai da capa do Descritivo (bytes do xlsx): local/UF, período, tempo médio, modo, regiões/bairros e mapas.

    Retorna {'achou': bool, 'local', 'uf', 'mes_ano', 'subtitulo', 'campo': (ini, fim)|None, 'tempo_medio', 'idade_minima',
    'modo', 'n_regioes', 'regioes': [..], 'regioes_texto': str, 'mapas': [{'bytes','ext','linha'}], 'avisos': [..]}.
    """
    base = {"achou": False, "local": None, "estado": None, "uf": None, "mes_ano": None, "subtitulo": None, "campo": None,
            "tempo_medio": None, "idade_minima": None, "modo": None, "n_regioes": None, "regioes": [], "regioes_texto": "",
            "mapas": [], "avisos": []}
    try:
        zf = zipfile.ZipFile(io.BytesIO(conteudo))
    except zipfile.BadZipFile:
        base["avisos"].append("Arquivo não é um xlsx válido.")
        return base
    itens = _ancoras(zf)
    if not itens:
        base["avisos"].append("Não encontrei a capa (caixas de texto) no arquivo.")
        return base
    base["achou"] = True
    textos = [i for i in itens if i["paragrafos"]]
    textos.sort(key=lambda i: i["linha"])
    # título: primeira caixa de texto (a mais acima)
    base.update({k: v for k, v in parse_titulo(textos[0]["paragrafos"], estados).items() if v})
    nota = next((i for i in textos if "NOTA METODOL" in " ".join(i["paragrafos"]).upper()), None)
    if nota:
        info = parse_nota(" ".join(nota["paragrafos"]))
        base.update({k: v for k, v in info.items() if v})
    else:
        base["avisos"].append("Nota metodológica não encontrada: datas e tempo médio ficam para preencher.")
    abr = next((i for i in textos if "ABRANG" in " ".join(i["paragrafos"]).upper() and len(" ".join(i["paragrafos"])) < 700), None)
    linha_abr = abr["linha"] if abr else None
    if abr:
        t = _sem_acento(_compacta(" ".join(abr["paragrafos"])))
        m = re.search(r"dividida em (\d+) regioes|(\d+) regioes", t)
        if m:
            base["n_regioes"] = int(m.group(1) or m.group(2))
        m = re.search(r"(\d+) cidades", t)
        if m:
            base["n_cidades"] = int(m.group(1))
    # caixa das regiões: a primeira, depois de ABRANGÊNCIA, com 2+ rótulos com peso (%) ou "Região n"
    cand = [i for i in textos if (linha_abr is None or i["linha"] >= linha_abr)]
    for i in cand:
        txt = _compacta(" ".join(i["paragrafos"]))
        regs = parse_regioes(txt)
        if len(regs) >= 2 and sum(1 for r in regs if r["peso"]) >= 2:
            base["regioes"] = regs
            base["regioes_texto"] = regioes_para_texto(regs)
            if not base.get("n_regioes"):
                base["n_regioes"] = len(regs)
            elif base["n_regioes"] != len(regs):
                base["avisos"].append(f"A capa diz {base['n_regioes']} regiões, mas listei {len(regs)}: confira o texto.")
            # segunda lista (ex.: "Regiões da Capital"): rótulos "N - Nome: bairros" sem peso
            for j in cand:
                if j is i or j["linha"] <= i["linha"]:
                    continue
                regs2 = parse_regioes(" ".join(j["paragrafos"]))
                if len(regs2) >= 2 and all(_ROTULO_REGIAO.match(r["rotulo"]) for r in regs2):
                    base["regioes_secundarias"] = regs2
                    base["regioes_secundarias_texto"] = regioes_para_texto(regs2)
                    tit = [t for t in textos if t["linha"] < j["linha"] and t["linha"] > i["linha"] and t["paragrafos"] and len(" ".join(t["paragrafos"])) < 60]
                    base["regioes_secundarias_titulo"] = _compacta(" ".join(tit[-1]["paragrafos"])) if tit else None
                    break
            break
    else:
        # a legenda pode estar na página seguinte e sem o peso (%) de cada região: aceita rótulos «Região N» / «N - Nome»
        for i in cand:
            regs = parse_regioes(" ".join(i["paragrafos"]))
            if len(regs) >= 2 and all(_ROTULO_REGIAO.match(r["rotulo"]) for r in regs):
                base["regioes"] = regs
                base["regioes_texto"] = regioes_para_texto(regs)
                base["n_regioes"] = base.get("n_regioes") or len(regs)
                break
        else:
            base["avisos"].append("Lista de regiões e bairros não encontrada na capa.")
    # mapas: imagens depois de ABRANGÊNCIA e antes de TECNOLOGIA
    linha_tec = next((i["linha"] for i in textos if "TECNOLOGIA" in " ".join(i["paragrafos"]).upper()), None)
    lim_ini = linha_abr if linha_abr is not None else 0
    for i in sorted(itens, key=lambda x: x["linha"]):
        if not i["imagens"] or i["linha"] < lim_ini:
            continue
        if linha_tec is not None and i["linha"] >= linha_tec and not (linha_abr is not None and i["linha"] < linha_tec):
            continue
        for nome, dados in i["imagens"]:
            base["mapas"].append({"bytes": dados, "ext": nome.rsplit(".", 1)[-1].lower(), "linha": i["linha"]})
    if not base["mapas"]:
        base["avisos"].append("Imagem do mapa não encontrada na capa.")
    return base
