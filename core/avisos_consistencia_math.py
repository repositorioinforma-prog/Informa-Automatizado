"""
Avisos de consistência do relatório (etapa final do Relatório Automatizado).

Três checagens que não mexem no arquivo — só apontam coisas para a pessoa
conferir antes de entregar:

1. Base de segmento MAIOR que a base Total da mesma tabela (quase sempre
   erro de filtro/ponderação no SPSS).
2. O MESMO título aparecendo com bases Totais diferentes (ex.: a tabela do
   Total e a dos segmentos rodadas com filtros diferentes).
3. Candidato/opção que aparece em um cenário e some de outro
   ("... - Cenário 1" x "... - Cenário 2"). Pode ser intencional (o
   cenário existe justamente para tirar alguém), por isso é só um aviso.

Tudo é lido da planilha já formatada: uma tabela = título (linha mesclada
na coluna A ou texto começando com 'Titulo'), cabeçalho, linhas de dado e
a linha 'Base' / 'Base reduzida' (coluna B = Total, C em diante =
segmentos).
"""
from __future__ import annotations

import re
import unicodedata
from collections import OrderedDict
from difflib import SequenceMatcher

# Rótulos que não são "candidatos" — ficam de fora da checagem de cenários.
ROTULOS_GENERICOS = {
    "outros", "outro", "outras", "outra", "branco/nulo", "branco / nulo",
    "branco", "nulo", "brancos/nulos", "nenhum", "nenhuma", "ninguem",
    "nao sabe", "nao sabe/nao respondeu", "ns/nr", "nao respondeu",
    "nao votaria", "nao vai votar", "nao vota", "indeciso", "indecisos",
    "total", "base", "base reduzida",
}
# Começos de rótulo que também não são candidatos ("Não irá votar", "Nenhum
# desses", "Branco", "Outros nomes"...).
_PREFIXOS_GENERICOS = ("nao ", "nenhum", "ninguem", "branco", "nulo", "outro", "outra", "indecis")


def _eh_generico(rotulo: str) -> bool:
    k = _chave(rotulo)
    return (
        k in ROTULOS_GENERICOS
        or k.startswith(_PREFIXOS_GENERICOS)
        or bool(_RE_CONTINUACAO.fullmatch(k))
        or k.startswith("continuacao")
    )


def _compacto(rotulo: str) -> str:
    return re.sub(r"[^a-z0-9]", "", _chave(rotulo))


def _parecidos(a: str, b: str) -> bool:
    ca, cb = _compacto(a), _compacto(b)
    if not ca or not cb:
        return False
    return ca == cb or SequenceMatcher(None, ca, cb).ratio() >= 0.9

_RE_CENARIO = re.compile(r"\s*[-–—/]?\s*cen[aá]rio\s*(\d+)\s*", re.IGNORECASE)
_RE_CONTINUACAO = re.compile(r"\s*[-–—]?\s*\(?\s*continua[cç][aã]o\s*\)?\s*", re.IGNORECASE)
_RE_PREFIXO_TITULO = re.compile(r"^\s*titulo\s*:\s*", re.IGNORECASE)


def _sem_acento(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )


def _chave(texto: str) -> str:
    return re.sub(r"\s+", " ", _sem_acento(str(texto)).strip().lower())


def _numero(valor):
    """Número de uma célula de base: aceita 604, 604.0, '25*', '25 *'."""
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        m = re.fullmatch(r"\s*(\d+(?:[.,]\d+)?)\s*\**\s*", valor)
        if m:
            return float(m.group(1).replace(",", "."))
    return None


def _fmt(n: float) -> str:
    return str(int(n)) if float(n).is_integer() else str(n).replace(".", ",")


def _titulo_limpo(texto: str) -> str:
    texto = _RE_PREFIXO_TITULO.sub("", str(texto))
    texto = _RE_CONTINUACAO.sub(" ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def extrair_tabelas(ws) -> list[dict]:
    """
    Lista as tabelas da planilha, na ordem: {titulo, linha_titulo,
    linha_base, cabecalhos (col -> texto), rotulos (lista col. A),
    base_total, bases_segmentos [(col, cabeçalho, valor)]}.
    Só entram tabelas que têm linha 'Base'/'Base reduzida' (pedaços de
    'Continuação' sem base são ignorados).
    """
    linhas_titulo = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row == rng.max_row and rng.min_col == 1 and rng.max_col > 1:
            v = ws.cell(row=rng.min_row, column=1).value
            if v is not None and str(v).strip():
                linhas_titulo.add(rng.min_row)

    # Cabeçalho de segmento: o texto pode estar numa mesclagem vertical/horizontal;
    # guardamos o valor do canto superior esquerdo para cada célula coberta.
    valor_mesclado = {}
    for rng in ws.merged_cells.ranges:
        v = ws.cell(row=rng.min_row, column=rng.min_col).value
        if v is None:
            continue
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                valor_mesclado[(r, c)] = v

    def celula(r, c):
        v = ws.cell(row=r, column=c).value
        return valor_mesclado.get((r, c)) if v is None else v

    tabelas = []
    titulo = None
    linha_titulo = None
    inicio_dados = None
    rotulos = []
    max_col = ws.max_column

    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        texto_a = str(a).strip() if a is not None else ""

        if r in linhas_titulo or _chave(texto_a).startswith("titulo"):
            titulo = _titulo_limpo(texto_a)
            linha_titulo = r
            inicio_dados = None
            rotulos = []
            continue
        if titulo is None:
            continue

        # Primeira linha com rótulo na coluna A depois do título = início dos dados
        # (o cabeçalho tem a coluna A vazia/mesclada sem texto próprio).
        if texto_a and inicio_dados is None and not _chave(texto_a).startswith("pergunta"):
            inicio_dados = r

        chave_a = _chave(texto_a)
        if chave_a in ("base", "base reduzida"):
            base_total = _numero(ws.cell(row=r, column=2).value)
            bases_seg = []
            for c in range(3, max_col + 1):
                n = _numero(ws.cell(row=r, column=c).value)
                if n is None:
                    continue
                cab = ""
                for rr in range((inicio_dados or r) - 1, linha_titulo, -1):
                    v = celula(rr, c)
                    if v is not None and str(v).strip():
                        cab = re.sub(r"\s+", " ", str(v).replace("-\n", "")).strip()
                        break
                bases_seg.append((c, cab, n))
            tabelas.append({
                "titulo": titulo,
                "linha_titulo": linha_titulo,
                "linha_base": r,
                "base_reduzida": chave_a == "base reduzida",
                "base_total": base_total,
                "bases_segmentos": bases_seg,
                "rotulos": list(rotulos),
            })
            titulo = None
            continue

        if inicio_dados is not None and texto_a and not chave_a.startswith("pergunta"):
            rotulos.append(texto_a)

    return tabelas


def checar_segmento_maior_que_total(tabelas) -> list[dict]:
    achados = []
    for t in tabelas:
        total = t["base_total"]
        if not total:
            continue
        for _col, cab, n in t["bases_segmentos"]:
            if n > total + 0.5:
                achados.append({
                    "titulo": t["titulo"], "linha": t["linha_base"],
                    "segmento": cab or "(sem nome)", "base_segmento": n, "base_total": total,
                })
    return achados


def checar_titulos_com_bases_diferentes(tabelas) -> list[dict]:
    grupos: "OrderedDict[str, list]" = OrderedDict()
    for t in tabelas:
        if t["base_total"] is None:
            continue
        grupos.setdefault(_chave(t["titulo"]), []).append(t)
    achados = []
    for lista in grupos.values():
        valores = sorted({t["base_total"] for t in lista})
        if len(valores) > 1:
            achados.append({
                "titulo": lista[0]["titulo"],
                "ocorrencias": [(t["linha_titulo"], t["base_total"]) for t in lista],
            })
    return achados


def checar_candidatos_entre_cenarios(tabelas) -> list[dict]:
    """
    Agrupa títulos que só diferem pelo 'Cenário N' e aponta os rótulos
    (candidatos) que não estão em todos os cenários do grupo.
    """
    grupos: "OrderedDict[str, OrderedDict]" = OrderedDict()
    for t in tabelas:
        m = _RE_CENARIO.search(t["titulo"])
        if not m:
            continue
        base_titulo = _RE_CENARIO.sub(" ", t["titulo"], count=1)
        base_titulo = re.sub(r"\s+", " ", base_titulo).strip(" -/")
        cenarios = grupos.setdefault(_chave(base_titulo), OrderedDict())
        nome = f"Cenário {int(m.group(1))}"
        info = cenarios.setdefault(nome, {"titulo_base": base_titulo, "rotulos": OrderedDict()})
        for rot in t["rotulos"]:
            if not _eh_generico(rot):
                info["rotulos"].setdefault(_chave(rot), rot)

    achados = []
    for cenarios in grupos.values():
        if len(cenarios) < 2:
            continue
        todos: "OrderedDict[str, str]" = OrderedDict()
        for info in cenarios.values():
            for k, rot in info["rotulos"].items():
                todos.setdefault(k, rot)
        faltas = []
        for k, rot in todos.items():
            presentes = [n for n, info in cenarios.items() if k in info["rotulos"]]
            ausentes = [n for n in cenarios if n not in presentes]
            if ausentes:
                faltas.append({"candidato": rot, "presente_em": presentes, "ausente_em": ausentes})

        # Mesma pessoa escrita diferente entre cenários ("Alexda Confecção" x
        # "Alex da Confecção") — sai da lista de "some" e vira aviso de grafia.
        grafias = []
        usados = set()
        for i, f1 in enumerate(faltas):
            for f2 in faltas[i + 1:]:
                if id(f1) in usados or id(f2) in usados:
                    continue
                if set(f1["presente_em"]) & set(f2["presente_em"]):
                    continue
                if _parecidos(f1["candidato"], f2["candidato"]):
                    grafias.append({
                        "grafias": [(f1["candidato"], f1["presente_em"]), (f2["candidato"], f2["presente_em"])],
                    })
                    usados.update({id(f1), id(f2)})
        faltas = [f for f in faltas if id(f) not in usados]

        if faltas or grafias:
            primeiro = next(iter(cenarios.values()))
            achados.append({
                "titulo": primeiro["titulo_base"], "cenarios": list(cenarios),
                "faltas": faltas, "grafias": grafias,
            })
    return achados


def gerar_avisos_consistencia(ws) -> dict:
    tabelas = extrair_tabelas(ws)
    return {
        "segmento_maior_que_total": checar_segmento_maior_que_total(tabelas),
        "titulos_bases_diferentes": checar_titulos_com_bases_diferentes(tabelas),
        "candidatos_cenarios": checar_candidatos_entre_cenarios(tabelas),
        "total_tabelas": len(tabelas),
    }


def textos_avisos(avisos: dict) -> list[str]:
    """Mensagens em markdown, uma por tipo de aviso encontrado."""
    msgs = []
    seg = avisos["segmento_maior_que_total"]
    if seg:
        # O mesmo segmento costuma repetir em todas as tabelas — agrupa.
        grupos: "OrderedDict[tuple, list]" = OrderedDict()
        for a in seg:
            grupos.setdefault((a["segmento"], a["base_segmento"], a["base_total"]), []).append(a)
        itens = []
        for (segmento, bs, bt), lista in grupos.items():
            exemplos = "; ".join(f"{a['titulo']} (linha {a['linha']})" for a in lista[:3])
            resto = f" e mais {len(lista) - 3}" if len(lista) > 3 else ""
            n = f"{len(lista)} tabela" + ("s" if len(lista) > 1 else "")
            itens.append(
                f"- **{segmento}**: base {_fmt(bs)} > Total {_fmt(bt)} em {n} — {exemplos}{resto}"
            )
        msgs.append(
            "**Base de segmento maior que a base Total.** Um segmento não pode ter "
            "mais entrevistas que o Total da própria tabela — confira o filtro/"
            "ponderação no SPSS:\n\n" + "\n".join(itens)
        )
    tit = avisos["titulos_bases_diferentes"]
    if tit:
        itens = "\n".join(
            f"- {a['titulo']}: "
            + ", ".join(f"base {_fmt(v)} (linha {l})" for l, v in a["ocorrencias"])
            for a in tit
        )
        msgs.append(
            "**Mesmo título com bases Totais diferentes.** Normalmente é a mesma "
            "pergunta rodada com filtros diferentes (ex.: Total x segmentos):\n\n" + itens
        )
    cen = avisos["candidatos_cenarios"]
    grafia = [a for a in cen if a["grafias"]]
    if grafia:
        itens = []
        for a in grafia:
            for g in a["grafias"]:
                partes = " x ".join(f"\"{rot}\" ({', '.join(c)})" for rot, c in g["grafias"])
                itens.append(f"- {a['titulo']}: {partes}")
        msgs.append(
            "**Mesmo candidato escrito diferente entre cenários.** Padronize o "
            "nome (e o partido) para ficar igual em todos:\n\n" + "\n".join(itens)
        )
    some = [a for a in cen if a["faltas"]]
    if some:
        blocos = []
        for a in some:
            por_cenario: "OrderedDict[str, list]" = OrderedDict((c, []) for c in a["cenarios"])
            for f in a["faltas"]:
                chave = " e ".join(f["presente_em"])
                por_cenario.setdefault(chave, []).append(f["candidato"])
            linhas = "\n".join(
                f"  - Só no {c}: {', '.join(nomes)}" for c, nomes in por_cenario.items() if nomes
            )
            blocos.append(f"- {a['titulo']}\n{linhas}")
        msgs.append(
            "**Candidato que aparece num cenário e some de outro.** Pode ser "
            "intencional (o cenário tira alguém), mas vale conferir:\n\n" + "\n".join(blocos)
        )
    return msgs
