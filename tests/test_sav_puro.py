"""Leitor de .sav em Python puro (plano B quando o pyreadstat é bloqueado).

Na criação, o leitor foi conferido contra os arquivos de teste oficiais do
pyreadstat e do pandas (sample.sav/.zsav, missing, UTF-8, hebraico, textos
longos) — mesmos valores, rótulos e missing. Aqui geramos arquivos .sav
pequenos (sem e com compressão) para fixar o comportamento.
"""
import math
import struct
import unittest

import numpy as np

from core.sav_puro import ErroSav, read_sav

SYSMIS = -1.7976931348623157e308


def _sav(vars_, linhas, *, comprimir=True, rotulos=None, missing=None, encoding="UTF-8"):
    """Gera um .sav mínimo. vars_: [(nome, largura, rótulo)] (largura 0 = numérica)."""
    rotulos = rotulos or {}
    missing = missing or {}
    slots = sum(1 if w == 0 else math.ceil(w / 8) for _, w, _ in vars_)
    out = bytearray(b"$FL2" + b"teste".ljust(60))
    out += struct.pack("<iiiii", 2, slots, 1 if comprimir else 0, 0, len(linhas))
    out += struct.pack("<d", 100.0) + b"01 Jan 26" + b"00:00:00" + b" " * 64 + b"\0\0\0"
    idx_slot = {}
    k = 1
    for nome, w, rot in vars_:
        idx_slot[nome] = k
        lab = rot.encode("utf-8")
        mv = missing.get(nome, [])
        fmt = (1 << 16) | (w << 8) if w else (5 << 16) | (8 << 8)
        out += struct.pack("<iiiiii", 2, w, 1, len(mv), fmt, fmt) + nome.upper()[:8].ljust(8).encode()
        out += struct.pack("<i", len(lab)) + lab + b"\0" * ((4 - len(lab) % 4) % 4)
        for m in mv:
            out += struct.pack("<d", m) if w == 0 else m.encode().ljust(8)
        k += 1
        for _ in range((math.ceil(w / 8) if w else 1) - 1):
            out += struct.pack("<iiiiii", 2, -1, 0, 0, 0, 0) + b" " * 8
            k += 1
    for nome, mapa in rotulos.items():
        out += struct.pack("<ii", 3, len(mapa))
        for v, l in mapa.items():
            lb = l.encode("utf-8")
            out += (struct.pack("<d", v) if isinstance(v, float) else v.encode().ljust(8))
            out += bytes([len(lb)]) + lb + b" " * ((8 - (len(lb) + 1) % 8) % 8)
        out += struct.pack("<iii", 4, 1, idx_slot[nome])
    longos = "\t".join(f"{n.upper()[:8]}={n}" for n, _, _ in vars_).encode()
    out += struct.pack("<iiii", 7, 13, 1, len(longos)) + longos
    e = encoding.encode()
    out += struct.pack("<iiii", 7, 20, 1, len(e)) + e
    out += struct.pack("<ii", 999, 0)

    celulas = []
    for linha in linhas:
        for (nome, w, _), x in zip(vars_, linha):
            if w == 0:
                celulas.append(("num", SYSMIS if x is None else float(x)))
            else:
                b = x.encode("utf-8").ljust(math.ceil(w / 8) * 8)
                celulas += [("txt", b[i:i + 8]) for i in range(0, len(b), 8)]
    if not comprimir:
        for t, x in celulas:
            out += struct.pack("<d", x) if t == "num" else x
        return bytes(out)
    cmds, raws = [], []
    for t, x in celulas + [("fim", None)]:
        if t == "fim":
            cmds.append(252)
        elif t == "num" and x == SYSMIS:
            cmds.append(255)
        elif t == "num" and float(x).is_integer() and -99 <= x <= 151:
            cmds.append(int(x) + 100)
        elif t == "txt" and x == b" " * 8:
            cmds.append(254)
        else:
            cmds.append(253)
            raws.append(struct.pack("<d", x) if t == "num" else x)
        if len(cmds) == 8 or t == "fim":
            out += bytes(cmds + [0] * (8 - len(cmds))) + b"".join(raws)
            cmds, raws = [], []
    return bytes(out)


VARS = [("sexo", 0, "Sexo"), ("Peso_final", 0, "Peso"), ("cidade", 12, "Cidade"), ("nota", 0, "Avaliação")]
LINHAS = [[1, 1.25, "São Paulo", 8], [2, 0.8, "Barueri", 99], [1, None, "", 5]]
ROT = {"sexo": {1.0: "Masculino", 2.0: "Feminino"}, "nota": {99.0: "Não sabe"}}


class TestSavPuro(unittest.TestCase):
    def _ler(self, **kw):
        return read_sav(_sav(VARS, LINHAS, rotulos=ROT, missing={"nota": [99.0]}), **kw)

    def test_dados_rotulos_e_nomes_longos(self):
        for comprimir in (True, False):
            df, meta = read_sav(_sav(VARS, LINHAS, rotulos=ROT, comprimir=comprimir))
            self.assertEqual(list(df.columns), ["sexo", "Peso_final", "cidade", "nota"])
            self.assertEqual(df["cidade"].tolist(), ["São Paulo", "Barueri", ""])
            self.assertEqual(df["Peso_final"].tolist()[:2], [1.25, 0.8])
            self.assertTrue(np.isnan(df["Peso_final"][2]))  # sysmis
            self.assertEqual(meta.column_names_to_labels["nota"], "Avaliação")
            self.assertEqual(meta.variable_value_labels["sexo"], {1.0: "Masculino", 2.0: "Feminino"})
            self.assertEqual(meta.file_encoding, "utf-8")

    def test_missing_de_usuario(self):
        df, meta = self._ler()
        self.assertTrue(np.isnan(df["nota"][1]))   # 99 vira NaN (padrão do pyreadstat)
        self.assertEqual(meta.missing_ranges, {})
        df, meta = self._ler(user_missing=True)
        self.assertEqual(df["nota"][1], 99.0)      # mantém o código
        self.assertEqual(meta.missing_ranges["nota"], [{"lo": 99.0, "hi": 99.0}])

    def test_somente_metadados_e_bytes_invalidos(self):
        df, meta = self._ler(metadataonly=True)
        self.assertEqual(len(df), 0)
        self.assertEqual(meta.number_rows, 3)
        with self.assertRaises(ErroSav):
            read_sav(b"isto nao e um sav")

    def test_integracao_com_carregar_banco(self):
        from core.analitica_dados import carregar_banco
        from core.pyreadstat_seguro import pyreadstat

        if not pyreadstat.usando_leitor_puro:
            self.skipTest("pyreadstat disponível: o plano B não é usado neste ambiente")
        b = carregar_banco(_sav(VARS, LINHAS, rotulos=ROT, missing={"nota": [99.0]}), "banco.sav", user_missing=True)
        self.assertEqual(b.var_labels["sexo"], "Sexo")
        self.assertEqual(b.missing_ranges["nota"][0]["lo"], 99.0)


if __name__ == "__main__":
    unittest.main()
