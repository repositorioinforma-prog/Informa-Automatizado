import re
import unicodedata

# "Textos estilizados" (small caps) usados em geradores de "fancy text" — não têm
# decomposição Unicode padrão, então mapeamos manualmente para a letra comum.
_SMALL_CAPS = {
    "ᴀ": "a", "ʙ": "b", "ᴄ": "c", "ᴅ": "d", "ᴇ": "e", "ꜰ": "f", "ɢ": "g",
    "ʜ": "h", "ɪ": "i", "ᴊ": "j", "ᴋ": "k", "ʟ": "l", "ᴍ": "m", "ɴ": "n",
    "ᴏ": "o", "ᴘ": "p", "ǫ": "q", "ʀ": "r", "ꜱ": "s", "ᴛ": "t", "ᴜ": "u",
    "ᴠ": "v", "ᴡ": "w", "x": "x", "ʏ": "y", "ᴢ": "z",
}
_SMALL_CAPS_TABLE = str.maketrans(_SMALL_CAPS)


def strip_accents(value: str) -> str:
    value = value.translate(_SMALL_CAPS_TABLE)
    return "".join(
        char
        # NFKD (não só NFD) também desmonta variantes de fonte "estilizadas" comuns
        # em respostas abertas — negrito/itálico matemático (𝐀𝐁𝐂), largura total
        # (ＡＢＣ), círculos (Ⓐ) etc. — de volta para a letra ASCII simples.
        for char in unicodedata.normalize("NFKD", value)
        if unicodedata.category(char) != "Mn"
    )


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = strip_accents(str(value))
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def is_generic_other_label(value) -> bool:
    normalized = normalize_text(value)
    generic = {
        "outro", "outros", "outra", "outras", "outro especifique",
        "outros especifique", "outra especifique", "outro qual",
        "outros qual", "especifique",
    }
    return normalized in generic
