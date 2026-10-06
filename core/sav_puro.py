"""
Leitor de arquivos SPSS (.sav / .zsav) em Python puro — sem DLL.

Motivo: no Windows do Instituto, a política de Controle de Aplicativo
(Smart App Control / WDAC) bloqueia a DLL compilada do pyreadstat
(`_readstat_parser`) dentro da .venv, e nenhuma tela conseguia ler .sav
localmente. Este módulo lê o mesmo formato só com a biblioteca padrão +
numpy/pandas, e devolve `(DataFrame, metadados)` no MESMO formato do
`pyreadstat.read_sav`, para servir de substituto transparente (ver
core/pyreadstat_seguro.py, que só o usa quando o pyreadstat não carrega).

Formato implementado conforme a documentação pública do PSPP ("System
File Format"):
- cabeçalho $FL2 (.sav) e $FL3 (.zsav, compressão zlib);
- sem compressão, compressão por bytecode e zlib;
- variáveis numéricas e de texto (inclusive textos > 255 caracteres,
  que o SPSS divide em segmentos);
- rótulos de variáveis, rótulos de valores (numéricos e de texto, inclusive
  de textos longos), nomes longos de variáveis, missing de usuário
  (valores discretos e faixas), codificação de caracteres;
- conversão de formatos de data/hora como o pyreadstat faz por padrão
  (DATE/ADATE/EDATE/SDATE/JDATE -> date, DATETIME -> datetime,
  TIME/DTIME -> time).

Compatibilidade com o pyreadstat.read_sav (parâmetros usados no app):
  read_sav(caminho, encoding=None, user_missing=False,
           apply_value_formats=False, metadataonly=False)
- user_missing=False: códigos de missing de usuário viram NaN (como o
  pyreadstat); user_missing=True: mantém os códigos e informa as faixas em
  `meta.missing_ranges`.
- metadados: column_names, column_labels, column_names_to_labels,
  variable_value_labels, value_labels, variable_to_label, missing_ranges,
  file_encoding, number_rows, number_columns, file_label.
"""
from __future__ import annotations

import datetime as _dt
import io
import math
import struct
import zlib
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd

SYSMIS = -1.7976931348623157e308  # -DBL_MAX: valor ausente do sistema
_SPSS_EPOCA = _dt.datetime(1582, 10, 14)

# códigos de formato do SPSS que o pyreadstat converte em data/hora
_FMT_DATA = {20, 23, 24, 28, 29, 30, 38, 39}   # DATE ADATE JDATE MOYR QYR WKYR EDATE SDATE
_FMT_DATAHORA = {22, 41, 42}                    # DATETIME YMDHMS MTIME?/DATETIME variantes
_FMT_HORA = {21, 25}                            # TIME DTIME

_CODEPAGES = {
    65001: "utf-8", 20127: "ascii", 28591: "latin-1", 28592: "iso8859-2", 28605: "iso8859-15",
    1250: "cp1250", 1251: "cp1251", 1252: "cp1252", 1253: "cp1253", 1254: "cp1254",
    1255: "cp1255", 1256: "cp1256", 1257: "cp1257", 1258: "cp1258", 437: "cp437", 850: "cp850",
    874: "cp874", 932: "cp932", 936: "gbk", 949: "cp949", 950: "cp950", 2: "cp1252", 3: "cp1252",
}


class ErroSav(ValueError):
    pass


@dataclass
class MetaSav:
    column_names: List[str] = field(default_factory=list)
    column_labels: List[Optional[str]] = field(default_factory=list)
    column_names_to_labels: Dict[str, Optional[str]] = field(default_factory=dict)
    variable_value_labels: Dict[str, Dict[Any, str]] = field(default_factory=dict)
    value_labels: Dict[str, Dict[Any, str]] = field(default_factory=dict)
    variable_to_label: Dict[str, str] = field(default_factory=dict)
    missing_ranges: Dict[str, List[Dict[str, Any]]] = field(default_factory=dict)
    file_encoding: Optional[str] = None
    file_label: Optional[str] = None
    number_rows: Optional[int] = None
    number_columns: int = 0
    original_variable_types: Dict[str, str] = field(default_factory=dict)
    readstat_variable_types: Dict[str, str] = field(default_factory=dict)
    lido_por: str = "sav_puro"


# ----------------------------------------------------------------------------
class _Leitor:
    def __init__(self, dados: bytes):
        self.b = dados
        self.p = 0
        self.e = "<"

    def ler(self, n: int) -> bytes:
        if self.p + n > len(self.b):
            raise ErroSav("Arquivo .sav truncado ou corrompido.")
        r = self.b[self.p:self.p + n]
        self.p += n
        return r

    def i32(self) -> int:
        v = struct.unpack_from(self.e + "i", self.b, self.p)[0]
        self.p += 4
        return v

    def i64(self) -> int:
        v = struct.unpack_from(self.e + "q", self.b, self.p)[0]
        self.p += 8
        return v

    def f64(self) -> float:
        v = struct.unpack_from(self.e + "d", self.b, self.p)[0]
        self.p += 8
        return v


@dataclass
class _Slot:
    """Uma entrada do dicionário (registro tipo 2), inclusive continuações."""
    nome_curto: bytes
    tipo: int              # 0 numérico, >0 largura de texto, -1 continuação
    rotulo: Optional[bytes]
    n_missing: int
    missing_raw: List[bytes]
    fmt_print: int


def _decodificar(b: bytes, enc: str) -> str:
    return b.decode(enc, errors="replace")


def _encoding_do_codepage(cp: Optional[int]) -> Optional[str]:
    if cp is None:
        return None
    return _CODEPAGES.get(cp, f"cp{cp}")


def _ler_dicionario(r: _Leitor):
    assin = r.ler(4)
    if assin not in (b"$FL2", b"$FL3"):
        raise ErroSav("Não é um arquivo SPSS (.sav/.zsav) válido.")
    r.ler(60)  # produto
    layout = struct.unpack_from("<i", r.b, r.p)[0]
    r.e = "<" if layout in (2, 3) else ">"
    r.i32()
    _nominal = r.i32()
    compressao = r.i32()
    _peso = r.i32()
    ncasos = r.i32()
    bias = r.f64()
    r.ler(9 + 8)
    rotulo_arquivo = r.ler(64)
    r.ler(3)

    slots: List[_Slot] = []
    rotulos_valor: List[Tuple[List[Tuple[bytes, bytes]], List[int]]] = []
    ext: Dict[int, List[bytes]] = {}
    while True:
        tipo = r.i32()
        if tipo == 2:
            t = r.i32()
            tem_rotulo = r.i32()
            n_miss = r.i32()
            fmt_print = r.i32()
            r.i32()  # formato de escrita
            nome = r.ler(8)
            rot = None
            if tem_rotulo:
                tam = r.i32()
                rot = r.ler(tam)
                r.ler((4 - tam % 4) % 4)
            miss = [r.ler(8) for _ in range(abs(n_miss))]
            slots.append(_Slot(nome, t, rot, n_miss, miss, fmt_print))
        elif tipo == 3:
            n = r.i32()
            pares = []
            for _ in range(n):
                val = r.ler(8)
                tam = r.ler(1)[0]
                rot = r.ler(tam)
                r.ler((8 - (tam + 1) % 8) % 8)
                pares.append((val, rot))
            if r.i32() != 4:
                raise ErroSav("Registro de rótulos de valor sem a lista de variáveis.")
            nv = r.i32()
            idxs = [r.i32() for _ in range(nv)]
            rotulos_valor.append((pares, idxs))
        elif tipo == 6:
            n = r.i32()
            r.ler(80 * n)
        elif tipo == 7:
            sub = r.i32()
            tam = r.i32()
            cont = r.i32()
            ext.setdefault(sub, []).append(r.ler(tam * cont))
        elif tipo == 999:
            r.i32()
            break
        else:
            raise ErroSav(f"Registro desconhecido no dicionário (tipo {tipo}).")
    return compressao, ncasos, bias, rotulo_arquivo, slots, rotulos_valor, ext


# ----------------------------------------------------------------------------
def _valores_brutos(r: _Leitor, compressao: int, bias: float, n_slots: int, ncasos: int) -> List[List[Any]]:
    """Devolve as linhas como listas de 'células de 8 bytes': float para
    numérico (None = sysmis) ou bytes(8) para pedaços de texto."""
    e = r.e
    if compressao == 2:  # zsav: blocos zlib com o fluxo bytecode dentro
        r.i64()  # zheader_ofs
        ztrailer_ofs = r.i64()
        ztrailer_len = r.i64()
        t = _Leitor(r.b)
        t.e = e
        t.p = ztrailer_ofs
        t.i64()  # bias
        t.i64()  # zero
        t.i32()  # block size
        nblocos = t.i32()
        partes = []
        for _ in range(nblocos):
            t.i64()  # uncompressed ofs
            cofs = t.i64()
            t.i32()  # uncompressed size
            csize = t.i32()
            partes.append(zlib.decompress(r.b[cofs:cofs + csize]))
        fluxo = b"".join(partes)
        _ = ztrailer_len
        return _bytecode(fluxo, 0, e, bias, n_slots, ncasos)
    if compressao == 1:
        return _bytecode(r.b, r.p, e, bias, n_slots, ncasos)
    # sem compressão: n_slots blocos de 8 bytes por caso
    linhas = []
    p = r.p
    fim = len(r.b)
    while (ncasos < 0 or len(linhas) < ncasos) and p + 8 * n_slots <= fim:
        linhas.append([r.b[p + 8 * k:p + 8 * k + 8] for k in range(n_slots)])
        p += 8 * n_slots
    return linhas


def _bytecode(b: bytes, p: int, e: str, bias: float, n_slots: int, ncasos: int) -> List[List[Any]]:
    linhas: List[List[Any]] = []
    atual: List[Any] = []
    fim = len(b)
    unpack_d = struct.Struct(e + "d").unpack_from
    espacos = b" " * 8
    while p + 8 <= fim:
        cmds = b[p:p + 8]
        p += 8
        acabou = False
        for c in cmds:
            if c == 0:
                continue
            if c == 252:
                acabou = True
                break
            if c == 253:
                atual.append(b[p:p + 8])
                p += 8
            elif c == 254:
                atual.append(espacos)
            elif c == 255:
                atual.append(None)
            else:
                atual.append(float(c - bias))
            if len(atual) == n_slots:
                linhas.append(atual)
                atual = []
                if 0 <= ncasos == len(linhas):
                    acabou = True
                    break
        if acabou:
            break
    return linhas


# ----------------------------------------------------------------------------
def read_sav(
    arquivo: Union[str, bytes, io.IOBase],
    encoding: Optional[str] = None,
    user_missing: bool = False,
    apply_value_formats: bool = False,
    metadataonly: bool = False,
    disable_datetime_conversion: bool = False,
    **_ignorados,
):
    """Mesma assinatura (no que o app usa) e mesmo retorno do pyreadstat.read_sav."""
    if isinstance(arquivo, (bytes, bytearray)):
        dados = bytes(arquivo)
    elif hasattr(arquivo, "read"):
        dados = arquivo.read()
    else:
        with open(arquivo, "rb") as f:
            dados = f.read()

    r = _Leitor(dados)
    compressao, ncasos, bias, rot_arq, slots, rotulos_valor, ext = _ler_dicionario(r)

    # ---- codificação
    enc = encoding
    if not enc and 20 in ext:
        enc = ext[20][0].decode("ascii", errors="ignore").strip("\x00 ").lower() or None
    if not enc and 3 in ext and len(ext[3][0]) >= 32:
        cp = struct.unpack_from(r.e + "8i", ext[3][0])[7]
        enc = _encoding_do_codepage(cp)
    enc = enc or "cp1252"
    try:
        "".encode(enc)
    except LookupError:
        enc = "cp1252"

    # ---- variáveis "lógicas" (pula continuações) e índice do slot
    vars_: List[Dict[str, Any]] = []
    slot_para_var: Dict[int, int] = {}
    for i, s in enumerate(slots):
        if s.tipo == -1:
            continue
        slot_para_var[i + 1] = len(vars_)  # índices do dicionário são 1-based
        n_slots_var = 1 if s.tipo == 0 else math.ceil(s.tipo / 8)
        curto = _decodificar(s.nome_curto, enc).rstrip()
        vars_.append({
            "curto": curto, "nome": curto, "largura": s.tipo, "slot": i,
            "n_slots": n_slots_var, "rotulo": _decodificar(s.rotulo, enc) if s.rotulo is not None else None,
            "slot_rec": s, "fmt": (s.fmt_print >> 16) & 0xFF,
        })

    # nomes longos (subtipo 13)
    if 13 in ext:
        mapa = {}
        for par in _decodificar(ext[13][0], enc).split("\t"):
            if "=" in par:
                k, v = par.split("=", 1)
                mapa[k.strip().upper()] = v.strip("\x00")
        for v in vars_:
            v["nome"] = mapa.get(v["curto"].upper(), v["curto"])

    # textos muito longos (subtipo 14): junta os segmentos de 255 em 1 variável
    remover = set()
    if 14 in ext:
        larguras = {}
        for par in _decodificar(ext[14][0], enc).split("\t"):
            par = par.strip("\x00")
            if "=" in par:
                k, v = par.split("=", 1)
                try:
                    larguras[k.strip().upper()] = int(v.strip("\x00 "))
                except ValueError:
                    pass
        for idx, v in enumerate(vars_):
            w = larguras.get(v["curto"].upper())
            if not w or idx in remover:
                continue
            n_seg = math.ceil(w / 252)
            segs = [idx + k for k in range(n_seg) if idx + k < len(vars_)]
            v["segmentos"] = [(vars_[s]["slot"], vars_[s]["n_slots"]) for s in segs]
            v["largura"] = w
            remover.update(segs[1:])
    vars_logicas = [v for i, v in enumerate(vars_) if i not in remover]

    meta = MetaSav()
    meta.file_encoding = enc
    meta.file_label = _decodificar(rot_arq, enc).strip() or None
    unpack_d = struct.Struct(r.e + "d").unpack

    # ---- missing de usuário
    def _missing_da_var(v) -> Tuple[List[Dict[str, Any]], Any]:
        s = v["slot_rec"]
        if not s.n_missing:
            return [], None
        if v["largura"] > 0:
            vals = [_decodificar(m, enc).rstrip() for m in s.missing_raw]
            return [{"lo": x, "hi": x} for x in vals], ("str", set(vals))
        nums = [unpack_d(m)[0] for m in s.missing_raw]
        faixas: List[Dict[str, Any]] = []
        if s.n_missing in (-2, -3):
            faixas.append({"lo": nums[0], "hi": nums[1]})
            discretos = nums[2:]
        else:
            discretos = nums
        faixas += [{"lo": x, "hi": x} for x in discretos]
        return faixas, ("num", faixas)

    info_missing = {}
    for v in vars_logicas:
        faixas, regra = _missing_da_var(v)
        if faixas:
            info_missing[v["nome"]] = regra
            if user_missing:  # o pyreadstat só informa as faixas com user_missing=True
                meta.missing_ranges[v["nome"]] = faixas

    # ---- rótulos de valor
    largura_por_slot = {v["slot"] + 1: v for v in vars_logicas}
    for pares, idxs in rotulos_valor:
        alvo = [largura_por_slot.get(i) for i in idxs]
        alvo = [a for a in alvo if a is not None]
        if not alvo:
            continue
        texto = alvo[0]["largura"] > 0
        mapa: Dict[Any, str] = {}
        for val, rot in pares:
            chave = _decodificar(val, enc).rstrip() if texto else unpack_d(val)[0]
            mapa[chave] = _decodificar(rot, enc)
        for a in alvo:
            meta.variable_value_labels.setdefault(a["nome"], {}).update(mapa)
    if 21 in ext:  # rótulos de valor de textos longos
        for bloco in ext[21]:
            t = _Leitor(bloco)
            t.e = r.e
            while t.p < len(bloco):
                nome = _decodificar(t.ler(t.i32()), enc)
                t.i32()  # largura
                n = t.i32()
                mapa = {}
                for _ in range(n):
                    val = _decodificar(t.ler(t.i32()), enc).rstrip()
                    mapa[val] = _decodificar(t.ler(t.i32()), enc)
                alvo = next((v["nome"] for v in vars_logicas if v["curto"].upper() == nome.upper()
                             or v["nome"].upper() == nome.upper()), nome)
                meta.variable_value_labels.setdefault(alvo, {}).update(mapa)

    nomes = [v["nome"] for v in vars_logicas]
    meta.column_names = nomes
    meta.column_labels = [v["rotulo"] for v in vars_logicas]
    meta.column_names_to_labels = dict(zip(nomes, meta.column_labels))
    meta.number_columns = len(nomes)
    meta.number_rows = ncasos if ncasos >= 0 else None
    for v in vars_logicas:
        tipo = "string" if v["largura"] > 0 else "double"
        meta.readstat_variable_types[v["nome"]] = tipo
        meta.original_variable_types[v["nome"]] = f"A{v['largura']}" if v["largura"] > 0 else "F8.2"
    for nome, mapa in meta.variable_value_labels.items():
        meta.value_labels[f"labels{len(meta.value_labels)}"] = mapa
        meta.variable_to_label[nome] = f"labels{len(meta.value_labels) - 1}"

    if metadataonly:
        return pd.DataFrame(columns=nomes), meta

    # ---- dados
    linhas = _valores_brutos(r, compressao, bias, len(slots), ncasos)
    if meta.number_rows is None:
        meta.number_rows = len(linhas)

    colunas: Dict[str, Any] = {}
    for v in vars_logicas:
        if v["largura"] == 0:
            k = v["slot"]
            col = np.empty(len(linhas), dtype=float)
            for i, ln in enumerate(linhas):
                cel = ln[k]
                if cel is None:
                    col[i] = np.nan
                elif isinstance(cel, float):
                    col[i] = cel
                else:
                    x = unpack_d(cel)[0]
                    col[i] = np.nan if x == SYSMIS else x
            regra = info_missing.get(v["nome"])
            if regra and not user_missing:
                for fx in regra[1]:
                    lo, hi = fx["lo"], fx["hi"]
                    col[(col >= lo) & (col <= hi)] = np.nan
            colunas[v["nome"]] = col
        else:
            segs = v.get("segmentos") or [(v["slot"], v["n_slots"])]
            valores = []
            for ln in linhas:
                pedacos = []
                for si, (ini, n) in enumerate(segs):
                    bruto = b"".join(c if isinstance(c, bytes) else b" " * 8 for c in ln[ini:ini + n])
                    pedacos.append(bruto[:252] if si < len(segs) - 1 else bruto)
                txt = _decodificar(b"".join(pedacos)[: v["largura"]], enc).rstrip().rstrip("\ufffd").rstrip()
                valores.append(txt)
            regra = info_missing.get(v["nome"])
            if regra and not user_missing:
                valores = [None if x in regra[1] else x for x in valores]
            colunas[v["nome"]] = valores

    df = pd.DataFrame(colunas, columns=nomes)

    # datas/horas como o pyreadstat (por padrão)
    if not disable_datetime_conversion:
        for v in vars_logicas:
            if v["largura"] != 0:
                continue
            f = v["fmt"]
            col = df[v["nome"]]
            if f in _FMT_DATA or f in _FMT_DATAHORA:
                convertidos = []
                for x in col:
                    if pd.isna(x):
                        convertidos.append(None)
                        continue
                    try:
                        d = _SPSS_EPOCA + _dt.timedelta(seconds=float(x))
                    except (OverflowError, ValueError):
                        convertidos.append(None)
                        continue
                    convertidos.append(d.date() if f in _FMT_DATA else d)
                if f in _FMT_DATAHORA and all(c is None or c.year >= 1678 for c in convertidos):
                    df[v["nome"]] = pd.to_datetime(pd.Series(convertidos, dtype=object))
                else:
                    df[v["nome"]] = pd.Series(convertidos, dtype=object)
            elif f in _FMT_HORA:
                df[v["nome"]] = [(_dt.datetime.min + _dt.timedelta(seconds=float(x))).time()
                                 if pd.notna(x) and 0 <= x < 86400 else None for x in col]
    return df, meta
