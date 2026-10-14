"""
Leitor mínimo, em Python puro, de arquivos .sav (SPSS), sem depender de pyreadstat
(usado apenas porque o ambiente de execução não tem acesso à internet para instalar
pyreadstat). Suporta:

- dicionário de variáveis (nome curto, tipo numérico/string, largura)
- nomes longos de variável (subtype 13)
- labels de variável
- VALUE LABELS numéricos e de string curta (rec type 3/4)
- VALUE LABELS de string longa (subtype 21)
- dados comprimidos (bytecode) e não comprimidos
- strings "muito longas" (subtype 14), reconstituindo os segmentos de 252 bytes

Não implementa: datas SPSS como tipo especial (fica double serial), pesos,
multiple response sets, etc. Suficiente para leitura/auditoria de variáveis
de respostas e VALUE LABELS.
"""
from __future__ import annotations

import struct
from dataclasses import dataclass, field


SYSMISS = -1e308  # sentinel usado internamente para system-missing numérico


@dataclass
class SavMeta:
    column_names: list
    column_labels: dict  # short name -> label
    variable_value_labels: dict  # short name -> {code: label}
    var_types: dict  # short name -> 0 (numeric) or width in bytes (string)
    number_rows: int


class _Reader:
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0

    def read(self, n):
        b = self.data[self.pos:self.pos + n]
        self.pos += n
        return b

    def read_i32(self):
        return struct.unpack_from('<i', self.data, self.pos)[0], self._adv(4)

    def i32(self):
        v = struct.unpack_from('<i', self.data, self.pos)[0]
        self.pos += 4
        return v

    def d(self):
        v = struct.unpack_from('<d', self.data, self.pos)[0]
        self.pos += 8
        return v

    def _adv(self, n):
        self.pos += n


def _decode_text(b: bytes) -> str:
    """Decodifica texto do .sav. A maioria dos arquivos modernos (modo Unicode)
    usa UTF-8 independentemente do que o subtype 20 diga; tentamos UTF-8 primeiro
    e caímos para latin-1 apenas se não for UTF-8 válido."""
    try:
        return b.decode('utf-8')
    except UnicodeDecodeError:
        return b.decode('latin-1', errors='replace')


def _pad4(n):
    return (n + 3) // 4 * 4


def _pad8(n):
    return (n + 7) // 8 * 8


def read_sav_bytes(raw: bytes):
    """Retorna (df, meta) igual, na forma usada pelo app, ao pyreadstat.read_sav."""
    r = _Reader(raw)
    rec_type = r.read(4)
    if rec_type not in (b'$FL2', b'$FL3'):
        raise ValueError("Arquivo não parece ser um .sav válido (assinatura ausente).")
    r.read(60)  # prod_name
    layout_code = r.i32()
    nominal_case_size = r.i32()
    compression_switch = r.i32()
    case_weight_index = r.i32()
    ncases_hdr = r.i32()
    bias = r.d()
    r.read(9)   # creation date
    r.read(8)   # creation time
    r.read(64)  # file label
    r.read(3)   # padding

    # -------- dicionário --------
    # cada variável "curta" (<=8 bytes de string, ou numérica) = 1 slot.
    # strings > 8 bytes usam slots extras (type == -1) que devem ser ignorados
    # como variáveis, mas contam para o índice usado nos value-label var-index records.
    var_slots = []  # um item por slot no dicionário (inclui slots de continuação)
    real_vars = []  # ordem de entrada: dicts com name, type(0=numeric,str width real), label
    short_name_order = []
    extra_records = {}
    pending_value_labels = []

    while True:
        rt = r.i32()
        if rt == 2:
            vtype = r.i32()
            has_label = r.i32()
            n_missing = r.i32()
            r.i32()  # print format
            r.i32()  # write format
            name = r.read(8).decode('latin-1').rstrip()
            label = None
            if has_label:
                label_len = r.i32()
                label_raw = r.read(label_len)
                label = _decode_text(label_raw)
                r.read(_pad4(label_len) - label_len)
            nmiss_abs = abs(n_missing)
            for _ in range(nmiss_abs):
                r.d()

            if vtype == -1:
                # continuação de string longa: apenas ocupa um slot
                var_slots.append({'continuation': True})
                continue

            slot = {
                'continuation': False,
                'name': name,
                'type': vtype,  # 0 numérica, 1-8 string curta (até 8 bytes neste slot)
                'label': label,
            }
            var_slots.append(slot)
            real_vars.append(slot)
            short_name_order.append(name)
        elif rt == 3:
            n_labels = r.i32()
            pairs = []
            for _ in range(n_labels):
                raw_val = r.read(8)
                label_len = r.read(1)[0]
                label_txt = _decode_text(r.read(label_len))
                consumed = 1 + label_len
                pad = _pad8(consumed) - consumed
                r.read(pad)
                pairs.append((raw_val, label_txt))
            rt4 = r.i32()
            assert rt4 == 4, f"esperado rec type 4 após 3, obtido {rt4}"
            n_idx = r.i32()
            idxs = [r.i32() for _ in range(n_idx)]
            pending_value_labels.append((pairs, idxs))
        elif rt == 4:
            # não deveria ocorrer isolado (sempre vem logo após um rec 3)
            n_idx = r.i32()
            for _ in range(n_idx):
                r.i32()
        elif rt == 6:
            n_lines = r.i32()
            r.read(80 * n_lines)
        elif rt == 7:
            subtype = r.i32()
            size = r.i32()
            count = r.i32()
            total = size * count
            payload = r.read(total)
            extra_records[subtype] = payload
            pad = 0
            # registros tipo 7 não têm padding adicional definido pelo spec além do próprio total
        elif rt == 999:
            r.i32()  # filler
            break
        else:
            raise ValueError(f"Registro de dicionário desconhecido: {rt} em offset {r.pos-4}")

    data_start = r.pos

    # -------- pós-processamento do dicionário --------
    long_names = {}
    if 13 in extra_records:
        txt = _decode_text(extra_records[13])
        for part in txt.split('\t'):
            if '=' in part:
                short, long_ = part.split('=', 1)
                long_names[short] = long_

    long_string_lengths = {}
    if 14 in extra_records:
        txt = _decode_text(extra_records[14].rstrip(b'\x00'))
        for part in txt.split('\t'):
            part = part.strip()
            if not part or '=' not in part:
                continue
            short, ln = part.split('=', 1)
            ln = ln.strip().rstrip('\x00')
            try:
                long_string_lengths[short] = int(ln)
            except ValueError:
                pass

    encoding = 'latin-1'
    if 20 in extra_records:
        try:
            encoding = extra_records[20].decode('ascii', errors='ignore').strip() or 'latin-1'
        except Exception:
            encoding = 'latin-1'

    # monta lista final de variáveis reais na ordem do dicionário, com nome longo se houver,
    # e largura final de string (considerando slots de continuação e "muito longas")
    variables = []  # dicts: short, name(final), type('num'/'str'), width, label, slot_index(0-based no var_slots)
    slot_index_by_realvar = []
    idx = 0
    n_slots = len(var_slots)
    while idx < n_slots:
        slot = var_slots[idx]
        if slot.get('continuation'):
            idx += 1
            continue
        short = slot['name']
        vtype = slot['type']
        label = slot['label']
        start_slot = idx
        if vtype == 0:
            variables.append({
                'short': short, 'type': 'num', 'width': 0, 'label': label,
                'slot_start': start_slot, 'n_slots': 1,
            })
            idx += 1
        else:
            # string: vtype = largura do 1º segmento (<=8). Slots seguintes de
            # continuação (type -1) estendem em blocos de 8 bytes.
            n_slots_used = 1
            j = idx + 1
            while j < n_slots and var_slots[j].get('continuation'):
                n_slots_used += 1
                j += 1
            # vtype já é a largura total declarada deste "pedaço" (até 255 bytes);
            # os slots de continuação (-1) apenas reservam os 8 bytes extras de dado,
            # não somam largura adicional ao vtype.
            chunk_width = vtype
            variables.append({
                'short': short, 'type': 'str', 'width': chunk_width, 'label': label,
                'slot_start': start_slot, 'n_slots': n_slots_used,
            })
            idx = j

    # aplica nomes longos (mantendo short como chave de dados por simplicidade;
    # o app usa os nomes de variável como estão no banco, então preservamos 'short'
    # como nome principal — value labels/labels seguem o short name)
    for v in variables:
        v['final_name'] = long_names.get(v['short'], v['short'])

    # -------- "strings muito longas" (subtype 14) --------
    # SPSS representa uma string > 255 bytes como uma cadeia de variáveis
    # "curtas" consecutivas (cada uma com seus próprios slots de continuação
    # tipo -1), com nomes auto-gerados após a 1ª. O subtype 14 informa o
    # comprimento lógico real; aqui fundimos essas variáveis em uma só.
    if long_string_lengths:
        merged_variables = []
        i = 0
        n_vars = len(variables)
        while i < n_vars:
            v = variables[i]
            total_len = long_string_lengths.get(v['short'])
            if v['type'] == 'str' and total_len is not None and total_len > v['width']:
                n_chunks = -(-total_len // 255)  # ceto (ceil)
                chunks = variables[i:i + n_chunks]
                merged = dict(v)
                merged['width'] = total_len
                merged['chunks'] = chunks
                merged_variables.append(merged)
                i += n_chunks
            else:
                merged_variables.append(v)
                i += 1
        variables = merged_variables

    # -------- VALUE LABELS numéricos/string-curta (rec 3/4) --------
    # os índices em pending_value_labels referem-se a slots do dicionário (1-based,
    # contando inclusive slots de continuação de string).
    slot_to_var = {}
    for v in variables:
        chunks = v.get('chunks', [v])
        for chunk in chunks:
            for s in range(chunk['slot_start'], chunk['slot_start'] + chunk['n_slots']):
                slot_to_var[s + 1] = v  # 1-based

    value_labels = {}  # final_name -> {code: label}
    for pairs, idxs in pending_value_labels:
        target_vars = []
        seen = set()
        for i in idxs:
            v = slot_to_var.get(i)
            if v is not None and id(v) not in seen:
                seen.add(id(v))
                target_vars.append(v)
        for v in target_vars:
            table = value_labels.setdefault(v['final_name'], {})
            for raw_val, label_txt in pairs:
                if v['type'] == 'num':
                    code = struct.unpack('<d', raw_val)[0]
                    if code == int(code):
                        code = int(code)
                    table[code] = label_txt
                else:
                    code = _decode_text(raw_val).rstrip()
                    table[code] = label_txt

    # -------- VALUE LABELS de string longa (subtype 21) --------
    if 21 in extra_records:
        payload = extra_records[21]
        p = 0
        n = len(payload)
        while p < n:
            var_name_len = struct.unpack_from('<i', payload, p)[0]; p += 4
            var_name = _decode_text(payload[p:p+var_name_len]); p += var_name_len
            p += _pad4(var_name_len) - var_name_len
            var_width = struct.unpack_from('<i', payload, p)[0]; p += 4
            n_labels = struct.unpack_from('<i', payload, p)[0]; p += 4
            table = value_labels.setdefault(long_names.get(var_name, var_name), {})
            for _ in range(n_labels):
                val_len = struct.unpack_from('<i', payload, p)[0]; p += 4
                val = _decode_text(payload[p:p+val_len]); p += val_len
                lab_len = struct.unpack_from('<i', payload, p)[0]; p += 4
                lab = _decode_text(payload[p:p+lab_len]); p += lab_len
                table[val] = lab

    # -------- dados --------
    # cada "elemento" de 8 bytes no fluxo de dados corresponde, na ordem, a:
    #   variável numérica -> 1 elemento
    #   variável string -> ceil(width/8) elementos (segmentos)
    elements = []  # ('num', var) ou ('strseg', var, seg_index, n_segs)
    for v in variables:
        if v['type'] == 'num':
            elements.append(('num', v))
        else:
            chunks = v.get('chunks', [v])
            chunk_info = []  # (chunk_width, n_segs_this_chunk)
            for chunk in chunks:
                cw = chunk['width']
                n_segs_c = max(1, -(-cw // 8)) if cw > 0 else 1
                chunk_info.append((cw, n_segs_c))
            v['_chunk_info'] = chunk_info
            n_segs = sum(c[1] for c in chunk_info)
            for seg in range(n_segs):
                elements.append(('strseg', v, seg, n_segs))

    n_elements = len(elements)
    import os
    if os.environ.get('MINISAV_DEBUG'):
        print('DEBUG n_elements', n_elements, 'nominal_case_size', nominal_case_size, 'n_var_slots', len(var_slots))
    rows = {v['final_name']: [] for v in variables}
    str_buffers = {v['final_name']: None for v in variables if v['type'] == 'str'}

    def flush_string(v):
        buf = str_buffers[v['final_name']]
        if buf is None:
            rows[v['final_name']].append(None)
        else:
            chunk_info = v.get('_chunk_info', [(v['width'], max(1, -(-v['width'] // 8)))])
            parts = []
            offset = 0
            for cw, n_segs_c in chunk_info:
                seg_len = n_segs_c * 8
                parts.append(buf[offset:offset + cw])
                offset += seg_len
            txt = _decode_text(b''.join(parts)).rstrip('\x00').rstrip()
            rows[v['final_name']].append(txt)
        str_buffers[v['final_name']] = None

    pos = data_start
    total_len = len(raw)

    if compression_switch == 0:
        elem_i = 0
        while pos < total_len and elem_i < n_elements * 10**9:
            kind = elements[elem_i % n_elements]
            if kind[0] == 'num':
                v = kind[1]
                val = struct.unpack_from('<d', raw, pos)[0]
                pos += 8
                rows[v['final_name']].append(None if val <= -1e307 else val)
            else:
                _, v, seg, n_segs = kind
                chunk = raw[pos:pos+8]
                pos += 8
                buf = str_buffers[v['final_name']]
                buf = (buf or b'') + chunk
                str_buffers[v['final_name']] = buf
                if seg == n_segs - 1:
                    flush_string(v)
            elem_i += 1
            if elem_i % n_elements == 0 and pos >= total_len:
                break
    else:
        elem_i = 0
        finished = False
        while pos < total_len and not finished:
            codes = raw[pos:pos+8]
            pos += 8
            if len(codes) < 8:
                break
            for code in codes:
                if finished:
                    break
                kind = elements[elem_i % n_elements]
                if code == 0:
                    # padding: não avança elem_i (ocorre entre casos, ex. após system-missing
                    # em certos writers) -- na prática tratamos como "sem dado", ignorar.
                    continue
                elif code == 252:
                    finished = True
                    break
                elif code == 253:
                    raw_bytes = raw[pos:pos+8]
                    pos += 8
                    if kind[0] == 'num':
                        v = kind[1]
                        val = struct.unpack('<d', raw_bytes)[0]
                        rows[v['final_name']].append(None if val <= -1e307 else val)
                    else:
                        _, v, seg, n_segs = kind
                        buf = str_buffers[v['final_name']]
                        buf = (buf or b'') + raw_bytes
                        str_buffers[v['final_name']] = buf
                        if seg == n_segs - 1:
                            flush_string(v)
                    elem_i += 1
                elif code == 254:
                    # segmento de string em branco
                    _, v, seg, n_segs = kind
                    buf = str_buffers[v['final_name']]
                    buf = (buf or b'') + b' ' * 8
                    str_buffers[v['final_name']] = buf
                    if seg == n_segs - 1:
                        flush_string(v)
                    elem_i += 1
                elif code == 255:
                    if kind[0] == 'num':
                        v = kind[1]
                        rows[v['final_name']].append(None)
                    else:
                        _, v, seg, n_segs = kind
                        buf = str_buffers[v['final_name']]
                        buf = (buf or b'') + b'\x00' * 8
                        str_buffers[v['final_name']] = buf
                        if seg == n_segs - 1:
                            flush_string(v)
                    elem_i += 1
                else:
                    val = code - bias
                    if kind[0] == 'num':
                        v = kind[1]
                        rows[v['final_name']].append(val)
                    else:
                        _, v, seg, n_segs = kind
                        # bytecode numérico não se aplica normalmente a string, mas por
                        # segurança tratamos como branco caso ocorra
                        buf = str_buffers[v['final_name']]
                        buf = (buf or b'') + b' ' * 8
                        str_buffers[v['final_name']] = buf
                        if seg == n_segs - 1:
                            flush_string(v)
                    elem_i += 1

    import pandas as pd

    ncases = ncases_hdr if ncases_hdr and ncases_hdr > 0 else None
    max_len = max((len(v) for v in rows.values()), default=0)
    if ncases:
        max_len = min(max_len, ncases) if max_len >= ncases else max_len

    data = {}
    for v in variables:
        col = rows[v['final_name']]
        if ncases and len(col) > ncases:
            col = col[:ncases]
        data[v['final_name']] = col

    df = pd.DataFrame(data)

    column_labels = {v['final_name']: (v['label'] or '') for v in variables}
    var_types = {v['final_name']: (0 if v['type'] == 'num' else v['width']) for v in variables}

    meta = SavMeta(
        column_names=[v['final_name'] for v in variables],
        column_labels=column_labels,
        variable_value_labels=value_labels,
        var_types=var_types,
        number_rows=len(df),
    )
    return df, meta


