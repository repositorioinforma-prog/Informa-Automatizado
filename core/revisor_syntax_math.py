"""Auditoria de sintaxes SPSS (.sps) com cruzamento opcional de banco .sav.

O módulo é deliberadamente conservador: interpreta apenas construções que
consegue reconhecer com segurança e emite alertas em vez de tentar executar
uma linguagem SPSS completa.

Além das checagens estruturais (RECODE sem regra, value label não gerado,
variável inexistente etc.), o módulo faz uma checagem SEMÂNTICA de
inversão: segue cada código do banco original através de VALUE LABELS ->
RECODE -> VALUE LABELS de destino e confere se a categoria de chegada
corresponde objetivamente a outra categoria que já existia no banco
original (ex.: código que deveria virar "Aprovação" chegou em "Reprovação").
Isso é diferente de comparar o texto das labels literalmente: reformulação
de redação (ex. "40 a 49 anos" -> "40 a 49") não gera alerta, só troca real
de categoria gera. Para variáveis de perfil/segmento (P1-P8, sexo, idade,
renda, escolaridade, raça, religião, região...) o módulo é ainda mais
conservador — variáveis de perfil legitimamente têm rótulos abreviados na
Syntax, então só aparece alerta quando é uma inversão real comprovada.
"""
from __future__ import annotations

from dataclasses import dataclass, asdict
from collections import Counter, defaultdict
from functools import lru_cache
import re
import unicodedata
from typing import Any, Iterable, Optional, Tuple


SEVERITY_ORDER = {"critico": 0, "possivel_erro": 1, "revisar": 2, "informativo": 3}
SEVERITY_LABELS = {
    "critico": "Erro crítico",
    "possivel_erro": "Possível erro",
    "revisar": "Revisar",
    "informativo": "Informativo",
}


@dataclass(frozen=True)
class Finding:
    severity: str
    code: str
    title: str
    message: str
    line: int | None = None
    variable: str | None = None
    snippet: str | None = None
    suggestion: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["severity_label"] = SEVERITY_LABELS.get(self.severity, self.severity)
        return data


@dataclass(frozen=True)
class RecodeRule:
    raw_source: str
    raw_target: str


@dataclass
class RecodeCommand:
    source_variables: list[str]
    target_variables: list[str]
    rules: list[RecodeRule]
    line: int
    raw: str


@dataclass
class ParsedSyntax:
    text: str
    encoding: str
    commands: list[tuple[int, str]]
    recodes: list[RecodeCommand]
    variable_labels: dict[str, tuple[str, int]]
    value_labels: dict[str, dict[Any, str]]
    value_label_lines: dict[str, int]
    mrsets_variables: list[tuple[str, int]]
    mrset_groups: dict[str, list[str]]
    section_markers: list[tuple[str, int]]
    referenced_variables: set[str]
    produced_variables: set[str]


def decode_sps(data: bytes) -> tuple[str, str, list[Finding]]:
    """Decodifica uma syntax preservando textos Windows-1252 comuns no legado."""
    findings: list[Finding] = []
    try:
        text = data.decode("utf-8-sig")
        encoding = "UTF-8"
    except UnicodeDecodeError:
        try:
            text = data.decode("cp1252")
            encoding = "Windows-1252"
        except UnicodeDecodeError:
            text = data.decode("latin-1", errors="replace")
            encoding = "ISO-8859-1"

    declared_utf8 = bool(re.search(r"^\s*\*\s*Encoding\s*:\s*UTF-?8\s*\.", text, flags=re.I | re.M))
    if declared_utf8 and encoding != "UTF-8":
        findings.append(Finding(
            "revisar", "ENCODING_DECLARADO_DIVERGENTE",
            "Encoding declarado não corresponde ao arquivo",
            f'A syntax declara UTF-8, mas foi decodificada como {encoding}. Isso pode causar problemas de acentuação em outros ambientes.',
            line=1,
            suggestion="Salvar o arquivo efetivamente como UTF-8 ou ajustar/remover a declaração de encoding.",
        ))
    return text.replace("\r\n", "\n").replace("\r", "\n"), encoding, findings


def _strip_inline_comments(text: str) -> str:
    # Comentários SPSS usuais começam por * no início lógico e terminam em ponto.
    # Aqui apenas preservamos comprimento de linhas; remoção de bloco é feita no parser de comandos.
    return text


def split_commands(text: str) -> list[tuple[int, str]]:
    """Separa comandos pelo ponto final fora de aspas.

    Linhas cujo primeiro caractere não branco é ``*`` são tratadas como
    comentários de linha. Isso é necessário porque as syntaxes do projeto
    usam separadores visuais ``*****Q1*****`` sem ponto final.
    """
    filtered_lines = []
    for line_text in text.splitlines(keepends=True):
        if line_text.lstrip().startswith("*"):
            # Mantém a quebra para preservar a numeração original.
            filtered_lines.append("\n" if line_text.endswith(("\n", "\r")) else "")
        else:
            filtered_lines.append(line_text)
    clean = "".join(filtered_lines)

    commands: list[tuple[int, str]] = []
    buf: list[str] = []
    line = 1
    start_line = 1
    quote: str | None = None
    at_statement_start = True

    i = 0
    while i < len(clean):
        ch = clean[i]
        if at_statement_start:
            if ch in " \t\n":
                if ch == "\n":
                    line += 1
                i += 1
                continue
            start_line = line
            at_statement_start = False

        if quote:
            buf.append(ch)
            if ch == quote:
                if i + 1 < len(clean) and clean[i + 1] == quote:
                    buf.append(clean[i + 1])
                    i += 1
                else:
                    quote = None
        else:
            if ch in "'\"":
                quote = ch
                buf.append(ch)
            elif ch == ".":
                cmd = "".join(buf).strip()
                if cmd:
                    commands.append((start_line, cmd + "."))
                buf = []
                at_statement_start = True
            else:
                buf.append(ch)

        if ch == "\n":
            line += 1
        i += 1

    trailing = "".join(buf).strip()
    if trailing:
        commands.append((start_line, trailing))
    return commands


def _normalize_var(name: str) -> str:
    return name.strip().upper()


# ---------------------------------------------------------------------------
# Identidade semântica de categorias — usada pra detectar inversão real
# distinguindo de simples reformulação de redação. Ported/adaptado do
# Auditor de Projeto (app desktop), reaproveitando o motor de RECODE já
# existente neste módulo (`_recode_value`) em vez de duplicá-lo.
# ---------------------------------------------------------------------------

def _num_key(value: Any) -> Any:
    try:
        f = float(value)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return str(value).strip()


@lru_cache(maxsize=4096)
def _canon(value: Any) -> str:
    """Normaliza texto pra comparação (sem acento, minúsculo, só
    alfanumérico) — cacheado porque a mesma label costuma se repetir
    dezenas de vezes num banco/syntax grande."""
    s = "" if value is None else str(value)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.casefold().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _strip_parenthetical(text: str) -> str:
    """Remove observações entre parênteses usadas só pra detalhar a label."""
    return re.sub(r"\([^)]*\)", " ", text or " ")


@lru_cache(maxsize=4096)
def _semantic_label(text: str) -> str:
    """Normaliza labels segundo o padrão usado nos relatórios (remove
    palavras de apoio que não mudam o significado da categoria, zera
    zeros à esquerda: '02 e 05' == '2 e 5')."""
    s = _canon(_strip_parenthetical(text))
    replacements = (
        "anos", "ano", "salarios minimos", "salario minimo", "salarios",
        "salario", "minimos", "minimo", "entre", "renda familiar",
    )
    for old in replacements:
        s = re.sub(rf"\b{re.escape(old)}\b", "", s)
    for old in ("o", "a", "os", "as", "de"):
        s = re.sub(rf"\b{re.escape(old)}\b", "", s)
    s = re.sub(r"\b0+(\d+)\b", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()


def _is_standard_segment(name: str, label: str = "") -> bool:
    """Identifica variáveis de perfil/segmentação do padrão do projeto
    (P1-P8 e subitens, sexo, idade, renda, escolaridade, raça, religião,
    região...). Essas variáveis podem ter labels propositalmente
    resumidas na Syntax, então não geram alerta por comparação literal."""
    n = _normalize_var(name).lstrip("$")
    l = _canon(label)

    if re.fullmatch(r"P\d+(?:_[A-Z0-9]+)?[A-Z]*", n):
        return True

    exact = {
        "SEXO", "SEX", "IDADE", "IDADE_A", "FAIXA", "FAIXAETARIA",
        "RENDA", "ESC", "ESCOLARIDADE", "COR", "RACA", "RAÇA",
        "RELIG", "REL", "RELIGIAO", "RELIGIÃO", "JUNCAORELIGIAO",
        "V7", "REGIAO", "REGIÃO", "RPA", "MESO", "MESORREGIAO",
        "MESORREGIÃO", "MUNICIPIO", "MUNICÍPIO", "BAIRRO",
        "CAT", "CAT1", "EVA", "EVA1", "AGN", "ATEU", "JEO", "JUD",
        "ESP", "AFRO", "ORI", "OU", "NS", "PEA", "NAOPEA", "NÃOPEA",
    }
    if n in exact:
        return True

    segment_terms = (
        "sexo", "idade", "faixa etaria", "renda", "escolaridade",
        "raca", "cor", "relig", "regiao", "mesorregiao", "rpa",
        "municipio", "bairro", "rede social", "redes sociais",
        "perfil", "segmento",
    )
    return any(term in l for term in segment_terms)


def _profile_dimension(name: str, label: str = "") -> Optional[str]:
    """Descobre qual perfil está sendo analisado a partir do NOME/LABEL DA
    VARIÁVEL — não da redação da categoria em si (uma label curta como
    'Até 02' só é entendida como renda porque a variável é RENDA/P3)."""
    n = _normalize_var(name).lstrip("$")
    l = _canon(label)

    if n in {"SEXO", "SEX", "P1"} or "sexo" in l:
        return "sexo"
    if n in {"IDADE", "IDADE_A", "FAIXA", "FAIXAETARIA", "P2", "MEP2"} or "idade" in l or "faixa etaria" in l:
        return "idade"
    if n in {"RENDA", "P3"} or "renda" in l or "salario minimo" in l:
        return "renda"
    if n in {"ESC", "ESCOLARIDADE"} or "escolaridade" in l:
        return "escolaridade"
    if n in {"COR", "RACA", "RAÇA"} or "raca" in l or l == "cor":
        return "raca"
    if n in {"RELIG", "REL", "RELIGIAO", "RELIGIÃO", "JUNCAORELIGIAO", "V7", "CAT", "CAT1",
             "EVA", "EVA1", "AGN", "ATEU", "JEO", "JUD", "ESP", "AFRO", "ORI", "OU", "NS"} or "relig" in l:
        return "religiao"
    if n in {"REGIAO", "REGIÃO", "RPA", "MESO", "MESORREGIAO", "MESORREGIÃO"} or "regiao" in l or "mesorregiao" in l:
        return "regiao"
    if n in {"MUNICIPIO", "MUNICÍPIO"} or "municipio" in l:
        return "municipio"
    if n == "BAIRRO" or "bairro" in l:
        return "bairro"
    return None


def _strict_label_key(text: str) -> str:
    """Chave conservadora pra detectar troca entre categorias textuais
    (região/município/bairro) sem tentar análise gramatical."""
    s = _canon(_strip_parenthetical(text))
    s = re.sub(r"\b(anos?|salarios?|minimos?|salario minimo|salarios minimos)\b", " ", s)
    s = re.sub(r"\b0+(\d+)\b", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()


def _profile_value_key(text: str, dimension: Optional[str] = None) -> Optional[Tuple[Any, ...]]:
    """Extrai a identidade objetiva de uma categoria de perfil (sexo,
    faixa etária, faixa de renda...) a partir do texto da label."""
    raw = text or ""
    c = _canon(raw)
    base = _canon(_strip_parenthetical(raw))
    if not base:
        return None

    dim = dimension
    if dim is None:
        if re.search(r"\b(masculino|homem|masc|feminino|mulher|fem)\b", base):
            dim = "sexo"
        elif "salario" in c or "renda" in c or "r$" in c:
            dim = "renda"
        elif "ano" in c or "idade" in c:
            dim = "idade"
        elif any(x in base for x in ("fundamental", "medio", "superior", "graduacao", "instrucao")):
            dim = "escolaridade"
        elif re.search(r"catolic|evangelic|testemunh.*jeova|judai|espirit|kardec|afro|umbanda|"
                        r"candomble|orient|budis|islam|agnost|ateu", base):
            dim = "religiao"

    if dim == "sexo":
        if re.search(r"\b(masculino|homem|masc)\b", base):
            return ("sexo", "masculino")
        if re.search(r"\b(feminino|mulher|fem)\b", base):
            return ("sexo", "feminino")
        return None

    if dim in {"idade", "renda"}:
        nums = [int(x) for x in re.findall(r"\b0*(\d+)\b", base)]
        if not nums:
            return None
        if "mais" in base:
            return (dim, "mais", nums[0])
        if "ate" in base and len(nums) == 1:
            return (dim, "ate", nums[0])
        if len(nums) >= 2:
            return (dim, "faixa", nums[0], nums[1])
        return (dim, "valor", nums[0])

    if dim == "escolaridade":
        if "sem instrucao" in base or "sem escolaridade" in base:
            return ("esc", "sem instrucao")
        nivel = None
        if "fundamental" in base:
            nivel = "fundamental"
        elif re.search(r"\bmedio\b", base):
            nivel = "medio"
        elif "superior" in base or "graduacao" in base:
            nivel = "superior"
        if nivel:
            if "incomplet" in base:
                return ("esc", nivel, "incompleto")
            if "complet" in base:
                return ("esc", nivel, "completo")
            return ("esc", nivel)
        return None

    if dim == "raca":
        race_map = {
            "branca": "branca", "branco": "branca",
            "preta": "preta", "preto": "preta", "negra": "preta", "negro": "preta",
            "parda": "parda", "pardo": "parda",
            "amarela": "amarela", "amarelo": "amarela",
            "indigena": "indigena",
        }
        for token, val in race_map.items():
            if re.search(rf"\b{re.escape(token)}\b", base):
                return ("raca", val)
        return None

    if dim == "religiao":
        religion_patterns = [
            (r"catolic", "catolica"),
            (r"evangelic|protestant", "evangelica"),
            (r"testemunh.*jeova", "testemunha de jeova"),
            (r"judai", "judaica"),
            (r"espirit|kardec", "espirita"),
            (r"afro|umbanda|candomble", "afro"),
            (r"orient|budis|islam", "oriental"),
            (r"agnost", "agnostico"),
            (r"ateu", "ateu"),
        ]
        for pat, val in religion_patterns:
            if re.search(pat, base):
                return ("religiao", val)
        if "outra" in base:
            return ("religiao", "outras")
        if "nao sabe" in base or base == "ns":
            return ("religiao", "nao sabe")
        return None

    if dim in {"regiao", "municipio", "bairro"}:
        key = _strict_label_key(raw)
        return (dim, key) if key else None

    if re.search(r"\bnao sabe\b|\bns\b", base):
        return ("especial", "nao sabe")
    if "branco nulo" in base or ("branco" in base and "nulo" in base):
        return ("especial", "branco nulo")
    if "nao respondeu" in base or re.fullmatch(r"nr", base):
        return ("especial", "nao respondeu")

    return None


def _category_key(text: str, dimension: Optional[str] = None) -> Optional[Tuple[Any, ...]]:
    """
    Identidade da categoria pra comparar BANCO ORIGINAL x VALUE LABELS da
    Syntax, seguindo o valor pelo RECODE (Homem==Masculino, 40 a 49 anos
    == 40 a 49, Não sei/Indeciso == Não sabe...). Quando não há
    equivalência conhecida, usa o texto normalizado — permite detectar,
    por exemplo, quando um candidato termina na label de outro.
    """
    raw = text or ""
    c = _canon(_strip_parenthetical(raw))
    if not c:
        return None

    if re.search(r"\bnao (sei|sabe)\b|\bindecis", c) or c in {"ns", "nao sabe"}:
        return ("especial", "nao_sabe")
    if ("branco" in c and "nulo" in c) or re.search(r"votarei.*nulo|em branco", c):
        return ("especial", "branco_nulo")
    if re.search(r"\bnao (irei|ira|vai|pretendo|pretende) votar\b|\bnao votarei\b", c):
        return ("especial", "nao_vai_votar")
    if re.fullmatch(r"outros?|outras?", c):
        return ("especial", "outros")
    if "nao conheco" in c and "suficiente" in c and "votar" in c:
        return ("especial", "nao_conhece_suficiente")

    if dimension:
        pk = _profile_value_key(raw, dimension)
        if pk is not None:
            return pk
        # Religião pode agrupar denominações detalhadas do banco (Batista,
        # Presbiteriana...) em categorias amplas ("Evangélica Total") — se
        # não der pra classificar com segurança, não cai na comparação
        # textual genérica (evitaria falso positivo em agrupamento legítimo).
        if dimension == "religiao":
            return None

    alias_groups = [
        ({"homem", "masculino", "masc"}, ("sexo", "masculino")),
        ({"mulher", "feminino", "fem"}, ("sexo", "feminino")),
        ({"sim", "yes"}, ("binario", "sim")),
        ({"nao", "no"}, ("binario", "nao")),
        ({"aprovacao", "aprova", "aprovado"}, ("avaliacao", "aprovacao")),
        ({"reprovacao", "reprova", "reprovado"}, ("avaliacao", "reprovacao")),
        ({"regular"}, ("avaliacao", "regular")),
    ]
    for aliases, key in alias_groups:
        if c in aliases:
            return key

    return ("texto", c)


def _parse_rule_group(group: str) -> RecodeRule | None:
    if "=" not in group:
        return None
    left, right = group.split("=", 1)
    left = " ".join(left.strip().split())
    right = " ".join(right.strip().split())
    if not left or not right:
        return None
    return RecodeRule(left, right)


def parse_recode(line: int, raw: str) -> RecodeCommand | None:
    m = re.match(r"\s*RECODE\s+(.+)$", raw, flags=re.I | re.S)
    if not m:
        return None
    body = m.group(1).rstrip(".").strip()
    into_match = re.search(r"\bINTO\b", body, flags=re.I)
    if into_match:
        before = body[:into_match.start()].strip()
        after = body[into_match.end():].strip()
        targets = [_normalize_var(v) for v in re.findall(r"[A-Za-z_@$#][\w@$#\.]*", after)]
    else:
        before = body
        targets = []

    first_paren = before.find("(")
    source_part = before if first_paren < 0 else before[:first_paren]
    sources = [_normalize_var(v) for v in re.findall(r"[A-Za-z_@$#][\w@$#\.]*", source_part)]
    groups = re.findall(r"\(([^()]*)\)", before[first_paren:] if first_paren >= 0 else "")
    rules = [r for g in groups if (r := _parse_rule_group(g)) is not None]
    return RecodeCommand(sources, targets, rules, line, raw.strip())


def _parse_variable_labels(commands: Iterable[tuple[int, str]]) -> dict[str, tuple[str, int]]:
    result: dict[str, tuple[str, int]] = {}
    for line, raw in commands:
        if not re.match(r"\s*(VARIABLE\s+LABELS|VAR\s+LABEL)\b", raw, flags=re.I):
            continue
        body = re.sub(r"^\s*(VARIABLE\s+LABELS|VAR\s+LABEL)\b", "", raw, flags=re.I).rstrip(".")
        for m in re.finditer(r"([A-Za-z_@$#][\w@$#\.]*)\s+(['\"])(.*?)\2", body, flags=re.S):
            result[_normalize_var(m.group(1))] = (m.group(3), line)
    return result


def _parse_value_labels(commands: Iterable[tuple[int, str]]) -> tuple[dict[str, dict[Any, str]], dict[str, int]]:
    values: dict[str, dict[Any, str]] = defaultdict(dict)
    lines: dict[str, int] = {}
    pair_re = re.compile(r"(?<!\S)([-+]?\d+(?:\.\d+)?|['\"][^'\"]*['\"])\s+(['\"])(.*?)\2(?=\s|$)", re.S)

    for line, raw in commands:
        if not re.match(r"\s*(VALUE\s+LABELS|VALUE\s+LABEL)\b", raw, flags=re.I):
            continue
        body = re.sub(r"^\s*(VALUE\s+LABELS|VALUE\s+LABEL)\b", "", raw, flags=re.I).rstrip(".").strip()
        # Seções podem vir como /var1 var2 1 'x' ... ou var 1 'x'.
        slash_sections = re.split(r"(?=/\s*[A-Za-z_@$#])", body)
        if len(slash_sections) == 1:
            slash_sections = [body]
        for section in slash_sections:
            section = section.strip()
            if not section:
                continue
            if section.startswith("/"):
                section = section[1:].lstrip()
            first_pair = pair_re.search(section)
            if not first_pair:
                continue
            header = section[:first_pair.start()].strip()
            vars_ = [_normalize_var(v) for v in re.findall(r"[A-Za-z_@$#][\w@$#\.]*", header)]
            if not vars_:
                continue
            for var in vars_:
                lines.setdefault(var, line)
            for pm in pair_re.finditer(section[first_pair.start():]):
                key_raw = pm.group(1)
                if key_raw[0:1] in {"'", '"'}:
                    key: Any = key_raw[1:-1]
                else:
                    num = float(key_raw)
                    key = int(num) if num.is_integer() else num
                label = pm.group(3)
                for var in vars_:
                    values[var][key] = label
    return dict(values), lines


def _extract_section_markers(text: str) -> list[tuple[str, int]]:
    markers: list[tuple[str, int]] = []
    for idx, line in enumerate(text.splitlines(), 1):
        if "*" not in line:
            continue
        cleaned = line.replace("*", " ").strip()
        # Aceita Q6, P1, Q12A etc.; evita pegar textos longos aleatórios.
        m = re.fullmatch(r"([PQ]\d+[A-Z]?)", cleaned, flags=re.I)
        if m:
            markers.append((_normalize_var(m.group(1)), idx))
    return markers


def _parse_mcgroups(text: str) -> dict[str, list[str]]:
    """
    Mapeia NAME -> lista de variáveis-membro de cada `/MCGROUP` (sintaxe
    de MRSETS usada via subcomando, comum quando a múltipla é declarada
    dentro de CTABLES em vez do comando `MRSETS ... VARIABLES=` isolado
    já coberto em `parse_syntax`). Cobre o caso adicional sem duplicar
    a lista `mrsets_variables` (que já é alimentada pelo laço principal).
    """
    out: dict[str, list[str]] = {}
    pat = re.compile(
        r"/MCGROUP\s+NAME\s*=\s*(\$[A-Za-z][\w$#@.]*)\b(.*?)(?=/MCGROUP\b|/DISPLAY\b|\.)",
        re.I | re.S,
    )
    for m in pat.finditer(text):
        name = _normalize_var(m.group(1))
        body = m.group(2)
        vm = re.search(r"\bVARIABLES\s*=\s*(.*?)(?=/|$)", body, re.I | re.S)
        if vm:
            members = re.findall(r"\b[A-Za-z][\w$#@.]*\b", vm.group(1))
            out[name] = [_normalize_var(x) for x in members]
    return out


def parse_syntax(data: bytes) -> tuple[ParsedSyntax, list[Finding]]:
    text, encoding, findings = decode_sps(data)
    commands = split_commands(_strip_inline_comments(text))
    recodes = [r for line, raw in commands if (r := parse_recode(line, raw)) is not None]
    variable_labels = _parse_variable_labels(commands)
    value_labels, value_label_lines = _parse_value_labels(commands)

    produced: set[str] = set()
    referenced: set[str] = set()
    mrsets_vars: list[tuple[str, int]] = []

    for rec in recodes:
        referenced.update(rec.source_variables)
        if rec.target_variables:
            produced.update(rec.target_variables)
        else:
            produced.update(rec.source_variables)

    for line, raw in commands:
        if re.match(r"\s*COMPUTE\b", raw, flags=re.I):
            m = re.match(r"\s*COMPUTE\s+([A-Za-z_@$#][\w@$#\.]*)\s*=", raw, flags=re.I)
            if m:
                produced.add(_normalize_var(m.group(1)))
        if re.match(r"\s*(IF|COUNT)\b", raw, flags=re.I):
            # Extrai referências de forma conservadora, filtrando palavras-chave.
            toks = {_normalize_var(t) for t in re.findall(r"\b[A-Za-z_][\w\.]*\b", raw)}
            referenced.update(toks - {"IF", "AND", "OR", "NOT", "COUNT", "THRU", "SYSMIS", "MISSING", "ANY"})
        if re.match(r"\s*MRSETS\b", raw, flags=re.I):
            for vm in re.finditer(r"\bVARIABLES\s*=\s*([^/\.]+)", raw, flags=re.I | re.S):
                for v in re.findall(r"[A-Za-z_@$#][\w@$#\.]*", vm.group(1)):
                    nv = _normalize_var(v)
                    mrsets_vars.append((nv, line))
                    referenced.add(nv)

    mrset_groups = _parse_mcgroups(text)
    for group_name in mrset_groups:
        referenced.add(group_name)

    parsed = ParsedSyntax(
        text=text, encoding=encoding, commands=commands, recodes=recodes,
        variable_labels=variable_labels, value_labels=value_labels,
        value_label_lines=value_label_lines, mrsets_variables=mrsets_vars,
        mrset_groups=mrset_groups,
        section_markers=_extract_section_markers(text), referenced_variables=referenced,
        produced_variables=produced,
    )
    return parsed, findings


def _to_number(value: Any) -> float | None:
    try:
        if isinstance(value, bool):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _source_matches(rule: RecodeRule, value: Any) -> bool:
    s = rule.raw_source.strip()
    su = s.upper()
    if su in {"ELSE", "MISSING", "SYSMIS"}:
        return False
    # lista por vírgula ou espaços: 1,2,3 / 1 2 3
    mt = re.fullmatch(r"(.+?)\s+THRU\s+(.+)", s, flags=re.I)
    if mt:
        v = _to_number(value); a = _to_number(mt.group(1)); b = _to_number(mt.group(2))
        return v is not None and a is not None and b is not None and a <= v <= b
    tokens = [t for t in re.split(r"[\s,]+", s) if t]
    for tok in tokens:
        if tok[:1] in {"'", '"'} and tok[-1:] == tok[:1]:
            if str(value) == tok[1:-1]:
                return True
        else:
            tv = _to_number(tok); vv = _to_number(value)
            if tv is not None and vv is not None and tv == vv:
                return True
            if str(value) == tok:
                return True
    return False


def _target_value(rule: RecodeRule, source_value: Any) -> Any:
    t = rule.raw_target.strip()
    tu = t.upper()
    if tu == "COPY":
        return source_value
    if tu in {"SYSMIS", "$SYSMIS"}:
        return None
    if t[:1] in {"'", '"'} and t[-1:] == t[:1]:
        return t[1:-1]
    n = _to_number(t)
    if n is not None:
        return int(n) if n.is_integer() else n
    return t


def _recode_value(rec: RecodeCommand, value: Any) -> tuple[Any, str]:
    for rule in rec.rules:
        if _source_matches(rule, value):
            return _target_value(rule, value), "matched"
    for rule in rec.rules:
        if rule.raw_source.upper() == "ELSE":
            return _target_value(rule, value), "else"
    return value, "unchanged"


def _normalize_domain(values: Iterable[Any]) -> list[Any]:
    clean = []
    for v in values:
        try:
            if v != v:  # NaN
                continue
        except Exception:
            pass
        clean.append(v)
    return clean


# ---------------------------------------------------------------------------
# Identificação de variáveis de perfil/segmento e checagem semântica de
# inversão — a parte mais valiosa portada do Auditor de Projeto.
# ---------------------------------------------------------------------------

def _extract_bank_labels(metadata) -> tuple[dict[str, str], dict[str, dict[Any, str]]]:
    """Extrai (labels de variável, value labels) de um objeto `meta` do
    pyreadstat, cobrindo tanto a API recente (`variable_value_labels`
    pronto) quanto o fallback mais antigo (`variable_to_label` +
    `value_labels`)."""
    if metadata is None:
        return {}, {}

    var_labels_raw = getattr(metadata, "column_names_to_labels", None) or {}
    var_labels = {_normalize_var(k): (v or "") for k, v in var_labels_raw.items()}

    value_labels: dict[str, dict[Any, str]] = {}
    vvl = getattr(metadata, "variable_value_labels", None) or {}
    if vvl:
        for var, mp in vvl.items():
            value_labels[_normalize_var(var)] = {_num_key(k): str(v) for k, v in (mp or {}).items()}
    else:
        var_to_label = getattr(metadata, "variable_to_label", None) or {}
        all_value_labels = getattr(metadata, "value_labels", None) or {}
        for var, labelset in var_to_label.items():
            mp = all_value_labels.get(labelset, {})
            value_labels[_normalize_var(var)] = {_num_key(k): str(v) for k, v in mp.items()}

    return var_labels, value_labels


def _identify_profile_variables(parsed: ParsedSyntax, bank_variable_labels: dict[str, str]) -> set[str]:
    """Monta a lista de variáveis de perfil/segmento do projeto —
    dinâmica, não uma lista fixa: propaga a classificação origem<->
    destino de cada RECODE (se P3 é perfil e vira RENDA_REC, RENDA_REC
    também é), e trata todo `$P...`/conjunto de perfil e seus membros
    como perfil (útil pro caso de redes sociais em múltipla, cujas
    labels individuais são só 'Sim'/'Não')."""
    candidates = (
        set(bank_variable_labels)
        | parsed.produced_variables
        | parsed.referenced_variables
        | set(parsed.variable_labels)
        | set(parsed.value_labels)
        | {v for v, _ in parsed.mrsets_variables}
        | set(parsed.mrset_groups)
    )

    profiles: set[str] = set()
    for var in candidates:
        label = parsed.variable_labels.get(var, ("", 0))[0] or bank_variable_labels.get(var, "")
        if _is_standard_segment(var, label):
            profiles.add(var)

    for set_name, members in parsed.mrset_groups.items():
        if _is_standard_segment(set_name) or set_name.lstrip("$").startswith("P"):
            profiles.add(set_name)
            profiles.update(members)

    changed = True
    while changed:
        changed = False
        for rec in parsed.recodes:
            targets = rec.target_variables or rec.source_variables
            for i, source in enumerate(rec.source_variables):
                target = targets[i] if i < len(targets) else targets[-1]
                source_label = bank_variable_labels.get(source, "") or parsed.variable_labels.get(source, ("", 0))[0]
                target_label = parsed.variable_labels.get(target, ("", 0))[0]
                is_profile = (
                    source in profiles or target in profiles
                    or _is_standard_segment(source, source_label)
                    or _is_standard_segment(target, target_label)
                )
                if is_profile:
                    before = len(profiles)
                    profiles.add(source)
                    profiles.add(target)
                    if len(profiles) != before:
                        changed = True

    return profiles


def _recode_pairs(rec: RecodeCommand) -> list[tuple[str, str]]:
    """Lista (origem, destino) de um RECODE — cobre RECODE de várias
    variáveis de uma vez (`RECODE V1 V2 (..) INTO T1 T2.`), aplicando o
    mesmo conjunto de regras a cada par posicionalmente, como o SPSS
    faz; sem INTO, cada origem recodifica sobre si mesma."""
    if not rec.target_variables:
        return [(s, s) for s in rec.source_variables]
    pairs = []
    for i, source in enumerate(rec.source_variables):
        target = rec.target_variables[i] if i < len(rec.target_variables) else rec.target_variables[-1]
        pairs.append((source, target))
    return pairs


def _mrset_members(parsed: ParsedSyntax) -> set[str]:
    return {m for members in parsed.mrset_groups.values() for m in members}


def _audit_semantic_inversions(parsed: ParsedSyntax, bank_value_labels: dict[str, dict[Any, str]],
                                bank_variable_labels: dict[str, str], profile_vars: set[str]) -> list[Finding]:
    """
    Segue cada código do banco original por VALUE LABEL -> RECODE ->
    VALUE LABEL de destino e confere se a categoria de chegada
    corresponde objetivamente a OUTRA categoria que já existia no banco
    original (inversão real) — reformulação de redação não conta.
    Reaproveita o motor de regras já existente (`_recode_value`) em vez
    de duplicar o parser de faixa/lista do RECODE.
    """
    findings: list[Finding] = []
    mrset_members = _mrset_members(parsed)

    for rec in parsed.recodes:
        for source, target in _recode_pairs(rec):
            src_labels = bank_value_labels.get(source) or parsed.value_labels.get(source) or {}
            dst_labels = parsed.value_labels.get(target) or {}
            if not src_labels or not dst_labels:
                continue

            # Indicador de múltipla tipo 1='Sim' -> label da categoria
            # (ex.: Instagram) é intencional, não é inversão.
            if target in mrset_members:
                source_semantics = {_canon(v) for v in src_labels.values()}
                if source_semantics & {"sim", "nao", "yes", "no"}:
                    continue

            is_profile = source in profile_vars or target in profile_vars
            dimension = None
            if is_profile:
                src_var_label = bank_variable_labels.get(source, "") or parsed.variable_labels.get(source, ("", 0))[0]
                dst_var_label = parsed.variable_labels.get(target, ("", 0))[0] or bank_variable_labels.get(target, "")
                dimension = _profile_dimension(source, src_var_label) or _profile_dimension(target, dst_var_label)

            original_by_key: dict[Tuple[Any, ...], list[tuple[Any, str]]] = defaultdict(list)
            for original_code, original_label in src_labels.items():
                key = _category_key(str(original_label), dimension if is_profile else None)
                if key is not None:
                    original_by_key[key].append((_num_key(original_code), str(original_label)))

            seen_errors = set()
            for original_code, original_label in src_labels.items():
                original_code = _num_key(original_code)
                final_code, mode = _recode_value(rec, original_code)
                if mode != "matched":
                    # Só compara mapeamento explícito (faixa/lista literal do
                    # RECODE) — ELSE tipicamente agrupa várias categorias
                    # num destino comum e não representa 1:1 comparável.
                    continue
                final_code = _num_key(final_code)
                final_label = dst_labels.get(final_code)
                if final_label is None:
                    continue

                src_key = _category_key(str(original_label), dimension if is_profile else None)
                dst_key = _category_key(str(final_label), dimension if is_profile else None)
                if src_key is None or dst_key is None or src_key == dst_key:
                    continue

                other_candidates = [
                    (code, label) for code, label in original_by_key.get(dst_key, [])
                    if _num_key(code) != original_code
                ]
                if not other_candidates:
                    continue

                other_code, other_label = other_candidates[0]
                dedupe = (target, original_code, final_code, other_code)
                if dedupe in seen_errors:
                    continue
                seen_errors.add(dedupe)

                titulo = "Inversão de segmento" if is_profile else "Inversão de valores"
                codigo_finding = "INVERSAO_SEGMENTO" if is_profile else "INVERSAO_VALORES"
                findings.append(Finding(
                    "critico", codigo_finding, titulo,
                    f"{source}: código {original_code} = '{original_label}' -> "
                    f"{target}: código {final_code} = '{final_label}'. No banco original, "
                    f"'{final_label}' corresponde à categoria {other_code} = '{other_label}'.",
                    rec.line, variable=target, snippet=rec.raw[:500],
                    suggestion="Conferir se os códigos do RECODE não foram trocados entre categorias.",
                ))

    return findings


def _audit_gov_ranges(parsed: ParsedSyntax) -> list[Finding]:
    """Validação especial pro padrão do projeto: escala 0-10 recodificada
    em Aprovação/Regular/Reprovação deve seguir 0-3=Reprovação,
    4-6=Regular, 7-10=Aprovação."""
    findings: list[Finding] = []
    for rec in parsed.recodes:
        for _source, target in _recode_pairs(rec):
            dst = parsed.value_labels.get(target) or {}
            if not dst:
                continue
            code_by_label = {_canon(label): code for code, label in dst.items()}
            approval = next((c for lab, c in code_by_label.items() if lab in {"aprovacao", "aprova"}), None)
            regular = next((c for lab, c in code_by_label.items() if lab == "regular"), None)
            rejection = next((c for lab, c in code_by_label.items() if lab in {"reprovacao", "reprova"}), None)
            if approval is None or regular is None or rejection is None:
                continue

            expected = {
                0: rejection, 1: rejection, 2: rejection, 3: rejection,
                4: regular, 5: regular, 6: regular,
                7: approval, 8: approval, 9: approval, 10: approval,
            }
            wrong = []
            for src_code, expected_out in expected.items():
                got, mode = _recode_value(rec, src_code)
                if mode != "matched":
                    continue
                if _num_key(got) != _num_key(expected_out):
                    wrong.append(f"{src_code}→{got} (esperado {expected_out})")
            if wrong:
                findings.append(Finding(
                    "critico", "ESCALA_GOV_INVERTIDA", "Escala GOV invertida",
                    "A recodificação 0–10 não segue 0–3=Reprovação, 4–6=Regular e 7–10=Aprovação: "
                    + "; ".join(wrong),
                    rec.line, variable=target, snippet=rec.raw[:500],
                    suggestion="Conferir os pontos de corte da recodificação da escala 0-10.",
                ))
    return findings


# ---------------------------------------------------------------------------
# Cruzamento opcional com o programa Total + Completo
# ---------------------------------------------------------------------------

def split_program(text: str) -> tuple[str, str]:
    """Separa o texto do programa em (Total, Completo) a partir de um
    cabeçalho `*****COMPLETO*****` (ou variação com menos asteriscos)."""
    m = re.search(r"(?im)^\s*\*+\s*COMPLETO\s*\*+\s*$", text)
    if m:
        return text[:m.start()], text[m.end():]
    m = re.search(r"(?i)\*{5,}[^\n]*COMPLETO[^\n]*\*{5,}", text)
    if m:
        return text[:m.start()], text[m.end():]
    return text, ""


def _is_report_question(name: str) -> bool:
    n = _normalize_var(name).lstrip("$")
    if re.match(r"^Q\d+[A-Z]*(?:R|A|X|_\d+R)?$", n):
        return True
    return n.startswith(("GOV", "MEGOV", "IPV", "MEIPV"))


def _program_variables(section: str) -> set[str]:
    """Variáveis referenciadas nas tabelas de um trecho do programa
    (`/TABLE=` legado ou `/TABLE $Q [...]` do CTABLES)."""
    vars_: set[str] = set()

    for m in re.finditer(r"/TABLE\s*=\s*(.*?)\s+BY\s+", section, re.I | re.S):
        for token in re.findall(r"\$?[A-Za-z][\w$#@.]*", m.group(1)):
            k = _normalize_var(token)
            if k.startswith("$T"):
                continue
            vars_.add(k)

    for m in re.finditer(r"/TABLE\s+(\$?[A-Za-z][\w$#@.]*)\s*\[", section, re.I):
        vars_.add(_normalize_var(m.group(1)))

    for m in re.finditer(r"\bVARIABLE\s+LABELS\s+([#$A-Za-z][\w$#@.]*)", section, re.I):
        k = _normalize_var(m.group(1))
        if _is_report_question(k):
            vars_.add(k)

    return vars_


def _audit_program_coverage(parsed: ParsedSyntax, program_text: str, bank_vars: set[str] | None = None) -> list[Finding]:
    """Confere se as perguntas usadas no Total/Completo existem
    (banco+Syntax), se o Completo cobre tudo que está no Total (e
    vice-versa) e se perguntas criadas na Syntax ficaram de fora dos
    dois — mesma lógica do Auditor de Projeto, adaptada pro parser
    deste módulo.

    `bank_vars` é essencial pra não gerar falso positivo em variáveis
    como Q1A (abertas codificadas à mão direto no banco, sem RECODE/
    COMPUTE nenhum na Syntax) — sem o banco, essas ficariam sempre
    marcadas como "não existe em lugar nenhum"."""
    findings: list[Finding] = []
    bank_vars = bank_vars or set()

    total_text, completo_text = split_program(program_text)
    total_vars = _program_variables(total_text)
    completo_vars = _program_variables(completo_text)
    has_completo = bool(completo_text.strip())

    available = parsed.produced_variables | parsed.referenced_variables | bank_vars
    for location, vars_ in (("Total", total_vars), ("Completo", completo_vars)):
        for var in sorted(vars_):
            if var not in available:
                findings.append(Finding(
                    "critico", "PROGRAMACAO_SEM_VARIAVEL", "Programação sem variável",
                    f"{var} é utilizada no {location}, mas não existe no banco nem é criada na Syntax.",
                    variable=var, snippet=location,
                ))

    if not has_completo:
        findings.append(Finding(
            "revisar", "ESTRUTURA_SEM_COMPLETO", "Sem cabeçalho de COMPLETO",
            "Não foi encontrado um cabeçalho de COMPLETO no arquivo de programação Total + Completo.",
        ))

    for var in sorted(total_vars - completo_vars):
        if _is_report_question(var):
            findings.append(Finding(
                "critico", "FALTA_NO_COMPLETO", "Falta no Completo",
                f"{var} aparece no Total, mas não foi encontrada no Completo.", variable=var,
            ))
    for var in sorted(completo_vars - total_vars):
        if _is_report_question(var):
            findings.append(Finding(
                "possivel_erro", "FALTA_NO_TOTAL", "Falta no Total",
                f"{var} aparece no Completo, mas não foi encontrada no Total. Confira se é intencional.",
                variable=var,
            ))

    mr_members = _mrset_members(parsed)
    # Candidatos a "pergunta criada na Syntax mas não programada": inclui o
    # que a Syntax produz E o que já existe pronto no banco SEM passar por
    # RECODE nenhum (Q1A e afins) — bank_vars que são só ORIGEM de RECODE
    # ficam de fora daqui de propósito, já que aí quem deve aparecer no
    # relatório é o destino recodificado, não a variável crua.
    recoded_sources = {s for rec in parsed.recodes for s in rec.source_variables}
    bank_vars_finais = bank_vars - recoded_sources
    produced_and_bank = parsed.produced_variables | bank_vars_finais
    programmed = total_vars | completo_vars
    candidates = [v for v in sorted(produced_and_bank) if v not in mr_members and _is_report_question(v)]
    candidates.extend(sorted(v for v in parsed.mrset_groups if _is_report_question(v)))
    for var in candidates:
        if var not in programmed:
            findings.append(Finding(
                "possivel_erro", "PERGUNTA_NAO_PROGRAMADA", "Pergunta não programada",
                f"{var} é criada na Syntax, mas não foi localizada nem no Total nem no Completo.",
                variable=var,
            ))

    return findings


def audit_syntax(parsed: ParsedSyntax, dataframe=None, metadata=None, processed_dataframe=None,
                  program_text: str | None = None) -> list[Finding]:
    findings: list[Finding] = []

    # Comando sem ponto final.
    if parsed.commands and not parsed.commands[-1][1].strip().endswith("."):
        line, raw = parsed.commands[-1]
        findings.append(Finding("critico", "COMANDO_SEM_PONTO", "Comando possivelmente incompleto",
                                "O último comando da syntax não termina em ponto final.", line, snippet=raw[:300],
                                suggestion="Finalizar o comando com '.' e revisar se o trecho foi truncado."))

    produced_at: dict[str, list[int]] = defaultdict(list)
    for rec in parsed.recodes:
        if not rec.source_variables:
            findings.append(Finding("critico", "RECODE_SEM_ORIGEM", "RECODE sem variável de origem",
                                    "Não foi possível identificar a variável de origem do RECODE.", rec.line, snippet=rec.raw[:500]))
            continue
        if not rec.rules:
            findings.append(Finding("critico", "RECODE_SEM_REGRAS", "RECODE sem regras reconhecíveis",
                                    "Nenhum par origem=destino foi reconhecido neste RECODE.", rec.line,
                                    variable=rec.source_variables[0], snippet=rec.raw[:500]))

        for target in (rec.target_variables or rec.source_variables):
            produced_at[target].append(rec.line)

        # Duplicidade textual de origem e mapeamentos conflitantes.
        seen: dict[str, str] = {}
        for rule in rec.rules:
            key = re.sub(r"\s+", " ", rule.raw_source.strip().upper())
            val = re.sub(r"\s+", " ", rule.raw_target.strip().upper())
            if key in seen:
                if seen[key] == val:
                    findings.append(Finding("revisar", "RECODE_REGRA_DUPLICADA", "Regra de RECODE repetida",
                                            f"A origem '{rule.raw_source}' aparece mais de uma vez com o mesmo destino '{rule.raw_target}'.",
                                            rec.line, variable=rec.source_variables[0], snippet=rec.raw[:500],
                                            suggestion="Remover a regra duplicada para reduzir risco de manutenção."))
                else:
                    findings.append(Finding("critico", "RECODE_REGRA_CONFLITANTE", "Regras de RECODE conflitantes",
                                            f"A origem '{rule.raw_source}' aparece com destinos diferentes ({seen[key]} e {val}). A primeira regra tende a prevalecer.",
                                            rec.line, variable=rec.source_variables[0], snippet=rec.raw[:500],
                                            suggestion="Definir um único destino para esta origem."))
            else:
                seen[key] = val

        # Value labels de destino que nenhuma regra literal pode gerar.
        for target in rec.target_variables:
            labels = parsed.value_labels.get(target, {})
            if labels:
                generated = set()
                has_copy = False
                for rule in rec.rules:
                    if rule.raw_target.upper() == "COPY":
                        has_copy = True
                    tv = _target_value(rule, None)
                    if tv is not None and rule.raw_target.upper() != "COPY":
                        generated.add(tv)
                if not has_copy:
                    for code, label in labels.items():
                        if code not in generated:
                            findings.append(Finding("possivel_erro", "VALUE_LABEL_NAO_GERADO",
                                                    "Value label não é produzido pelo RECODE",
                                                    f"{target} possui o código {code!r} ('{label}'), mas nenhuma regra deste RECODE gera esse valor.",
                                                    rec.line, variable=target, snippet=rec.raw[:500],
                                                    suggestion="Verificar se faltou uma regra no RECODE ou se sobrou um value label copiado de outra variável."))

    # Variável destino reescrita por múltiplos recodes.
    for var, lines in produced_at.items():
        if len(lines) > 1:
            findings.append(Finding("revisar", "VARIAVEL_RECODIFICADA_MULTIPLAS_VEZES",
                                    "Variável recebe RECODE mais de uma vez",
                                    f"{var} é produzida por RECODE nas linhas {', '.join(map(str, lines))}. Isso pode ser intencional, mas merece conferência.",
                                    lines[1], variable=var))

    # Marcadores de seção aparentemente vazios.
    command_lines = [line for line, _ in parsed.commands]
    for idx, (marker, line) in enumerate(parsed.section_markers):
        end = parsed.section_markers[idx + 1][1] if idx + 1 < len(parsed.section_markers) else 10**9
        has_command = any(line < c_line < end for c_line in command_lines)
        if not has_command:
            findings.append(Finding("revisar", "SECAO_SEM_PROGRAMACAO", "Seção identificada sem programação",
                                    f"O marcador {marker} aparece na syntax, mas não há comando SPSS entre ele e a próxima seção.",
                                    line, variable=marker,
                                    suggestion="Confirmar se a ausência de programação é proposital."))

    # Erros textuais recorrentes e de alta confiança em labels/perguntas.
    typos = {
        "estudual": "estadual",
        "canditado": "candidato",
        "prefeitoa": "prefeita",
        "deputadoa": "deputada",
    }
    for idx, line_text in enumerate(parsed.text.splitlines(), 1):
        lower = line_text.lower()
        for wrong, right in typos.items():
            if re.search(rf"\b{re.escape(wrong)}\b", lower):
                findings.append(Finding(
                    "revisar", "POSSIVEL_ERRO_TEXTUAL", "Possível erro de escrita",
                    f"Foi encontrado '{wrong}'. O termo esperado provavelmente é '{right}'.",
                    idx, snippet=line_text.strip()[:500],
                    suggestion=f"Revisar o texto e, se for erro de digitação, substituir '{wrong}' por '{right}'.",
                ))

    # Banco original opcional.
    bank_vars: set[str] = set()
    if dataframe is not None:
        cols = {str(c).upper(): c for c in dataframe.columns}
        bank_vars = set(cols)
        known_created = set(parsed.produced_variables)
        for rec in parsed.recodes:
            for src in rec.source_variables:
                if src not in cols and src not in known_created:
                    findings.append(Finding("critico", "VARIAVEL_ORIGEM_INEXISTENTE", "Variável de origem não existe no banco",
                                            f"{src} é usada em RECODE, mas não existe entre as colunas do .sav original nem foi criada antes pela syntax.",
                                            rec.line, variable=src, snippet=rec.raw[:500]))

            # Auditoria de domínio somente para RECODE 1->1, caso comum do projeto.
            if len(rec.source_variables) == 1 and len(rec.target_variables) <= 1:
                src = rec.source_variables[0]
                if src in cols:
                    series = dataframe[cols[src]]
                    counts = Counter(_normalize_domain(series.tolist()))
                    for value, count in counts.items():
                        result, mode = _recode_value(rec, value)
                        if mode == "else" and result is None:
                            findings.append(Finding("critico", "VALOR_REAL_VIRA_SYSMIS",
                                                    "Valor existente no banco será perdido pelo RECODE",
                                                    f"{src} contém o valor {value!r} em {count} caso(s), e esse valor cai em ELSE=SYSMIS.",
                                                    rec.line, variable=src, snippet=rec.raw[:500],
                                                    suggestion="Confirmar se esses casos devem realmente virar missing ou incluir a categoria na recodificação."))
                        elif mode == "unchanged" and rec.target_variables:
                            # INTO sem ELSE mantém SYSMIS no destino para casos não casados no SPSS.
                            findings.append(Finding("possivel_erro", "VALOR_REAL_NAO_CONTEMPLADO",
                                                    "Valor existente no banco não é contemplado pelo RECODE",
                                                    f"{src} contém o valor {value!r} em {count} caso(s), mas nenhuma regra o contempla. Como há INTO, o destino pode ficar missing nesses casos.",
                                                    rec.line, variable=src, snippet=rec.raw[:500],
                                                    suggestion="Revisar se faltou mapear esta categoria."))

        # MRSETS e demais referências diretas.
        for var, line in parsed.mrsets_variables:
            if var not in cols and var not in known_created:
                findings.append(Finding("critico", "MRSETS_VARIAVEL_INEXISTENTE", "Variável de MRSETS não existe",
                                        f"{var} foi incluída em um MRSETS, mas não existe no banco original nem é criada pela syntax.",
                                        line, variable=var))

    # Banco processado opcional: valida existência de destinos e labels de domínio.
    if processed_dataframe is not None:
        pcols = {str(c).upper(): c for c in processed_dataframe.columns}
        for var in sorted(parsed.produced_variables):
            if var not in pcols:
                findings.append(Finding("critico", "DESTINO_AUSENTE_BANCO_PROCESSADO",
                                        "Variável criada não aparece no banco processado",
                                        f"A syntax cria {var}, mas essa variável não foi encontrada no .sav processado enviado.",
                                        variable=var,
                                        suggestion="Confirmar se a syntax completa foi executada e se o banco processado corresponde a esta versão."))

    # Identificação de perfis/segmentos + checagem semântica de inversão —
    # roda sempre (não depende do banco: cai pra usar os VALUE LABELS já
    # declarados na própria Syntax quando a origem é uma variável criada
    # em cadeia, e usa o banco original quando disponível).
    bank_variable_labels, bank_value_labels = _extract_bank_labels(metadata)
    profile_vars = _identify_profile_variables(parsed, bank_variable_labels)
    findings.extend(_audit_semantic_inversions(parsed, bank_value_labels, bank_variable_labels, profile_vars))
    findings.extend(_audit_gov_ranges(parsed))

    # Cruzamento opcional com Total + Completo.
    if program_text:
        findings.extend(_audit_program_coverage(parsed, program_text, bank_vars))

    # REGRA DOS PERFIS/SEGMENTOS: dentro de variáveis de perfil, só
    # mostramos inversão real — o resto (padronização de redação, label
    # ausente, duplicidade etc.) é ruído esperado nesse tipo de variável
    # e fica oculto pra não afogar os alertas que realmente importam.
    findings = [
        f for f in findings
        if f.variable not in profile_vars or f.code == "INVERSAO_SEGMENTO"
    ]

    # Ordenação e deduplicação exata.
    unique: dict[tuple, Finding] = {}
    for f in findings:
        key = (f.severity, f.code, f.line, f.variable, f.message)
        unique[key] = f
    return sorted(unique.values(), key=lambda f: (SEVERITY_ORDER.get(f.severity, 99), f.line or 10**9, f.code))


def review_syntax(sps_bytes: bytes, dataframe=None, metadata=None, processed_dataframe=None,
                   program_bytes: bytes | None = None) -> dict[str, Any]:
    parsed, decode_findings = parse_syntax(sps_bytes)
    program_text = None
    if program_bytes:
        program_text, _prog_encoding, _prog_decode_findings = decode_sps(program_bytes)
    findings = decode_findings + audit_syntax(parsed, dataframe, metadata, processed_dataframe, program_text)
    findings = sorted(findings, key=lambda f: (SEVERITY_ORDER.get(f.severity, 99), f.line or 10**9, f.code))
    counts = Counter(f.severity for f in findings)
    return {
        "encoding": parsed.encoding,
        "commands": len(parsed.commands),
        "recodes": len(parsed.recodes),
        "variable_labels": len(parsed.variable_labels),
        "value_label_variables": len(parsed.value_labels),
        "mrsets_variables": len(parsed.mrsets_variables),
        "produced_variables": sorted(parsed.produced_variables),
        "referenced_variables": sorted(parsed.referenced_variables),
        "counts": {k: counts.get(k, 0) for k in SEVERITY_ORDER},
        "findings": [f.to_dict() for f in findings],
    }
