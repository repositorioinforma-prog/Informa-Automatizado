import re
from collections import OrderedDict


def decode_sps_bytes(raw: bytes) -> str:
    """Decodifica o conteúdo de um arquivo .sps de forma resiliente.

    Muitos exports de sintaxe SPSS (principalmente de sistemas mais antigos,
    tipo Data Entry / CATI) são salvos em CP1252/Latin-1, mesmo quando o
    próprio arquivo tem um comentário dizendo "* Encoding: UTF-8." — esse
    comentário não garante nada sobre a codificação real dos bytes. Decodificar
    esse tipo de arquivo como UTF-8 com errors='replace' não dá erro, mas
    corrompe silenciosamente toda letra acentuada (ex.: 'ã' vira '�'), o que
    estraga os VALUE LABELS sem avisar ninguém. Por isso tentamos primeiro
    UTF-8 estrito e só caímos para CP1252 (e por último Latin-1) se ele falhar.
    """
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")

_QUOTED = r"'(?:[^']|'')*'|\"(?:[^\"]|\"\")*\""
_CODE = rf"([-+]?\d+(?:\.\d+)?|{_QUOTED})"


def _unquote(label: str) -> str:
    if not label or label[0] not in "'\"":
        return label
    quote = label[0]
    body = label[1:-1]
    return body.replace(quote * 2, quote)


def _decode_code(raw_code: str):
    """Um código de VALUE LABELS pode ser numérico ('1') ou uma string entre
    aspas ('RJ') — variáveis string também podem ter VALUE LABELS no SPSS."""
    if raw_code and raw_code[0] in "'\"":
        return _unquote(raw_code)
    number = float(raw_code)
    return int(number) if number.is_integer() else number


def parse_all_value_labels_sps(content: str) -> dict[str, dict[float | int | str, str]]:
    """Parse all VALUE LABELS blocks from SPSS syntax."""
    blocks = re.finditer(
        r"\bVALUE\s+LABELS\b(.*?)(?=(?:\bVALUE\s+LABELS\b)|\Z)",
        content,
        flags=re.IGNORECASE | re.DOTALL,
    )
    parsed: dict[str, dict[float | int | str, str]] = OrderedDict()
    for match in blocks:
        body = match.group(1)
        first_pair = re.search(rf"(?<!\S){_CODE}\s+({_QUOTED})", body)
        if not first_pair:
            continue
        variables_part = body[: first_pair.start()].strip()
        variables = [v for v in re.split(r"\s+", variables_part) if v and v != "/"]
        if not variables:
            continue
        entries = re.findall(rf"(?<!\S){_CODE}\s+({_QUOTED})", body)
        labels: dict[float | int | str, str] = OrderedDict()
        for raw_code, raw_label in entries:
            labels[_decode_code(raw_code)] = _unquote(raw_label)
        for variable in variables:
            parsed.setdefault(variable, OrderedDict()).update(labels)
    return parsed
