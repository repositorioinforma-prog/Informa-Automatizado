"""
Escritor mínimo, em Python puro, de arquivos .sav (SPSS), sem depender de
pyreadstat (mesmo motivo do minisav.py: ambiente sem acesso à internet).

Gera arquivos NÃO comprimidos (mais simples e robusto) com:
- nomes curtos automáticos + nomes longos reais via subtype 13;
- labels de variável;
- VALUE LABELS numéricos e de string;
- encoding UTF-8 (subtype 20);
- variáveis string de até 255 bytes (suficiente para este projeto — strings
  maiores são truncadas com aviso, ver `truncated` no retorno).

Não escreve: compressão bytecode, "strings muito longas" (>255), pesos,
formatos de data especiais (datas viram string ou numérico simples).
"""
from __future__ import annotations

import datetime as _dt
import struct

SYSMISS = -1.7976931348623157e+308  # valor system-missing padrão do SPSS


def _pad4(n):
    return (n + 3) // 4 * 4


def _short_names(long_names):
    """Gera nomes curtos (<=8, A-Z0-9_, únicos) preservando a ordem."""
    used = set()
    mapping = {}
    for name in long_names:
        base = "".join(ch if (ch.isalnum() or ch == "_") else "_" for ch in name.upper())
        base = base[:8] or "VAR"
        candidate = base
        i = 1
        while candidate in used:
            suffix = str(i)
            candidate = (base[: 8 - len(suffix)] + suffix)
            i += 1
        used.add(candidate)
        mapping[name] = candidate
    return mapping


def write_sav_bytes(
    df,
    var_types: dict,
    var_labels: dict | None = None,
    value_labels: dict | None = None,
    file_label: str = "",
) -> tuple[bytes, list[str]]:
    """
    df: pandas.DataFrame — cada coluna é uma variável, na ordem de saída.
    var_types: nome -> 0 (numérica) ou largura em bytes (string). Colunas não
        listadas são inferidas do dtype do pandas.
    var_labels: nome -> label de variável (opcional).
    value_labels: nome -> {codigo: label} (opcional).
    Retorna (bytes_do_arquivo, avisos).
    """
    import pandas as pd

    var_labels = var_labels or {}
    value_labels = value_labels or {}
    warnings: list[str] = []

    columns = list(df.columns)
    short = _short_names(columns)

    # -------- determina tipo/largura final de cada coluna --------
    col_info = []  # dicts: name, short, kind('num'/'str'), width
    for name in columns:
        declared = var_types.get(name)
        series = df[name]
        is_numeric_dtype = pd.api.types.is_numeric_dtype(series) and declared in (None, 0)
        if declared == 0 or (declared is None and is_numeric_dtype):
            col_info.append({"name": name, "short": short[name], "kind": "num", "width": 0})
        else:
            # largura: usa o conteúdo real (mais robusto que confiar na largura original)
            as_text = series.dropna().astype(str)
            real_max = int(as_text.str.encode("utf-8", errors="replace").str.len().max()) if len(as_text) else 0
            width = max(real_max, 1)
            width = max(width, 8)  # largura mínima prática para reamostragem/edição futura
            width = min(width, 255)
            if real_max > 255:
                warnings.append(f"Coluna '{name}': texto truncado para 255 bytes (máximo encontrado: {real_max}).")
            col_info.append({"name": name, "short": short[name], "kind": "str", "width": width})

    # -------- cabeçalho --------
    now = _dt.datetime.now()
    creation_date = now.strftime("%d %b %y")[:9].ljust(9)
    creation_time = now.strftime("%H:%M:%S")[:8].ljust(8)
    prod_name = "@(#) Gerador de Controle Geral - Python".ljust(60)[:60]

    ncases = len(df)
    header = b"$FL2"
    header += prod_name.encode("latin-1", errors="replace")
    header += struct.pack("<i", 2)          # layout_code
    nominal_case_size_placeholder = 0
    header += struct.pack("<i", nominal_case_size_placeholder)
    header += struct.pack("<i", 0)          # compression_switch = 0 (sem compressão)
    header += struct.pack("<i", 0)          # case_weight_index
    header += struct.pack("<i", ncases)
    header += struct.pack("<d", 100.0)      # compression_bias
    header += creation_date.encode("ascii", errors="replace")
    header += creation_time.encode("ascii", errors="replace")
    header += file_label.ljust(64)[:64].encode("latin-1", errors="replace")
    header += b"\x00\x00\x00"

    # -------- registros de variável (rec type 2) --------
    body = bytearray()
    slot_start_1based = {}  # nome -> índice (1-based) do primeiro slot no dicionário
    slot_counter = 0

    def write_var_record(vtype, name8, label, n_missing=0):
        nonlocal body
        rec = struct.pack("<i", 2)
        rec += struct.pack("<i", vtype)
        rec += struct.pack("<i", 1 if label else 0)
        rec += struct.pack("<i", n_missing)
        if vtype == 0:
            fmt = (5 << 16) | (8 << 8) | 2  # F8.2
        elif vtype == -1:
            fmt = 0
        else:
            fmt = (1 << 16) | (min(vtype, 255) << 8) | 0  # A<width>
        rec += struct.pack("<i", fmt)
        rec += struct.pack("<i", fmt)
        rec += name8.encode("ascii", errors="replace")[:8].ljust(8)
        if label:
            label_bytes = label.encode("utf-8", errors="replace")
            rec += struct.pack("<i", len(label_bytes))
            rec += label_bytes
            rec += b" " * (_pad4(len(label_bytes)) - len(label_bytes))
        body += rec

    for col in col_info:
        slot_counter += 1
        slot_start_1based[col["name"]] = slot_counter
        label = var_labels.get(col["name"], "")
        if col["kind"] == "num":
            write_var_record(0, col["short"], label)
        else:
            write_var_record(col["width"], col["short"], label)
            n_extra = -(-col["width"] // 8) - 1
            for _ in range(n_extra):
                slot_counter += 1
                write_var_record(-1, "", "")

    nominal_case_size = slot_counter

    # -------- VALUE LABELS (rec type 3 + 4) --------
    for name, table in value_labels.items():
        if name not in slot_start_1based or not table:
            continue
        kind = next(c["kind"] for c in col_info if c["name"] == name)
        rec = struct.pack("<i", 3)
        rec += struct.pack("<i", len(table))
        for code, label in table.items():
            if kind == "num":
                try:
                    code_val = float(code)
                except (TypeError, ValueError):
                    continue
                rec += struct.pack("<d", code_val)
            else:
                raw = str(code).encode("utf-8", errors="replace")[:8]
                rec += raw.ljust(8, b" ")
            label_bytes = str(label).encode("utf-8", errors="replace")
            label_len = min(len(label_bytes), 255)
            label_bytes = label_bytes[:label_len]
            rec += bytes([label_len])
            rec += label_bytes
            consumed = 1 + label_len
            rec += b" " * (((consumed + 7) // 8 * 8) - consumed)
        rec += struct.pack("<i", 4)
        rec += struct.pack("<i", 1)
        rec += struct.pack("<i", slot_start_1based[name])
        body += rec

    # -------- subtype 13: nomes longos --------
    long_names_txt = "\t".join(f"{col['short']}={col['name']}" for col in col_info)
    payload13 = long_names_txt.encode("utf-8", errors="replace")
    body += struct.pack("<i", 7)
    body += struct.pack("<i", 13)
    body += struct.pack("<i", 1)
    body += struct.pack("<i", len(payload13))
    body += payload13

    # -------- subtype 20: encoding --------
    payload20 = b"UTF-8"
    body += struct.pack("<i", 7)
    body += struct.pack("<i", 20)
    body += struct.pack("<i", 1)
    body += struct.pack("<i", len(payload20))
    body += payload20

    # -------- fim do dicionário --------
    body += struct.pack("<i", 999)
    body += struct.pack("<i", 0)

    # -------- dados (sem compressão) --------
    data = bytearray()
    for _, row in df.iterrows():
        for col in col_info:
            val = row[col["name"]]
            if col["kind"] == "num":
                if val is None or (isinstance(val, float) and val != val):  # NaN
                    data += struct.pack("<d", SYSMISS)
                else:
                    try:
                        data += struct.pack("<d", float(val))
                    except (TypeError, ValueError):
                        data += struct.pack("<d", SYSMISS)
            else:
                if val is None or (isinstance(val, float) and val != val):
                    text_bytes = b""
                else:
                    text_bytes = str(val).encode("utf-8", errors="replace")[: col["width"]]
                n_segs = -(-col["width"] // 8)
                padded = text_bytes.ljust(n_segs * 8, b" ")
                data += padded

    # agora que sabemos o nominal_case_size real, regrava no cabeçalho
    header = bytearray(header)
    struct.pack_into("<i", header, 68, nominal_case_size)

    return bytes(header) + bytes(body) + bytes(data), warnings
