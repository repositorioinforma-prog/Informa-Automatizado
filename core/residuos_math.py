"""
Análise Residual (A.R) — cálculos do app desktop "Relatória Analítica".

O trecho entre as marcas "INÍCIO/FIM DO CÓDIGO ORIGINAL" é cópia SEM
ALTERAÇÕES de app/pages/residuals_page.py (linhas 27–211): ordem dos itens
fixos no fim, percentuais, resíduo ajustado (SPSS), cores do VBA e nomes de
abas. A lógica dos @Slot (prévia, fila e exportação) está em
analises/analise_residual.py, também com o código original.
"""
from __future__ import annotations

import math
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# ===================== INÍCIO DO CÓDIGO ORIGINAL (residuals_page.py) =====================
# ============================================================
# CONFIG: itens sugeridos no final (igual Crosstabs)
# ============================================================
FIXED_BOTTOM_LABELS = [
    "Não sei",
    "Branco/Nulo",
    "Branco / Nulo",
    "Não votarei",
    "Não irá votar",
    "Não sabe",
    "Sem resposta",
    "(Sem resposta)",
]




def _round_half_up(x: float, ndigits: int = 1) -> float:
    q = Decimal("1").scaleb(-ndigits)
    return float(Decimal(str(x)).quantize(q, rounding=ROUND_HALF_UP))


def _canon(s: str) -> str:
    s = (s or "").strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def _is_peso_name(name: str) -> bool:
    return "peso" in _canon(name)


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    try:
        if isinstance(v, float) and math.isnan(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _apply_fixed_bottom_order(labels: List[str], fixed_bottom: List[str]) -> List[str]:
    fixed_canon = [_canon(x) for x in fixed_bottom]
    present = labels[:]

    fixed_present: List[str] = []
    for f in fixed_canon:
        for lab in present:
            if _canon(lab) == f and lab not in fixed_present:
                fixed_present.append(lab)
                break

    rest = [lab for lab in present if _canon(lab) not in set(fixed_canon)]
    return rest + fixed_present


def _sheet_name_safe(name: str) -> str:
    bad = r"[\[\]\:\*\?\/\\]"
    name = re.sub(bad, " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "A.R"


def _unique_sheet_name(base: str, used: set) -> str:
    base = _sheet_name_safe(base)
    if len(base) > 31:
        base = base[:31].rstrip()

    if base not in used:
        used.add(base)
        return base

    i = 2
    while True:
        suffix = f" ({i})"
        cut = 31 - len(suffix)
        cand = (base[:cut].rstrip() + suffix)
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


# VBA colors
FILL_RED = "FBE2D5"    # RGB(251, 226, 213)
FONT_RED = "7E350E"    # RGB(126, 53, 14)
FILL_BLUE = "DAE9F8"   # RGB(218, 233, 248)
FONT_BLUE = "153D64"   # RGB(21, 61, 100)


def _cell_style_for_z(z: float) -> Optional[Tuple[str, str]]:
    if z is None or not np.isfinite(z):
        return None
    if z <= -1.85:
        return (FILL_RED, FONT_RED)
    if z >= 1.85:
        return (FILL_BLUE, FONT_BLUE)
    return None


def _round_obs_to_integer(ct: pd.DataFrame) -> pd.DataFrame:
    return ct.apply(
        lambda s: s.map(
            lambda v: float(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        )
    )


def _pct_matrix(ct: pd.DataFrame, mode: str) -> pd.DataFrame:
    """
    Igual Crosstabs:
      - col: percentual por coluna
      - row: percentual por linha
      - total: percentual no total
    """
    if mode == "col":
        denom = ct.sum(axis=0).replace(0, float("nan"))
        p = ct.divide(denom, axis=1) * 100.0
    elif mode == "row":
        denom = ct.sum(axis=1).replace(0, float("nan"))
        p = ct.divide(denom, axis=0) * 100.0
    elif mode == "total":
        denom = float(ct.values.sum())
        p = (ct / denom) * 100.0 if denom else ct * 0.0
    else:
        p = ct * 0.0

    p = p.fillna(0.0)
    p = p.apply(lambda s: s.map(lambda v: _round_half_up(float(v), 1)))
    return p


def _adjusted_residuals(obs: np.ndarray) -> np.ndarray:
    """
    Adjusted Residual (SPSS-like):
      z_ij = (O_ij - E_ij) / sqrt( E_ij * (1 - r_i) * (1 - c_j) )
    """
    O = obs.astype(float)
    total = float(np.nansum(O))
    if not np.isfinite(total) or total <= 0:
        return np.full_like(O, np.nan)

    row_sum = np.nansum(O, axis=1)
    col_sum = np.nansum(O, axis=0)
    E = np.outer(row_sum, col_sum) / total

    with np.errstate(divide="ignore", invalid="ignore"):
        ri = row_sum / total
        cj = col_sum / total
        denom = np.sqrt(E * (1.0 - ri[:, None]) * (1.0 - cj[None, :]))
        z = (O - E) / denom

    z[~np.isfinite(z)] = np.nan
    return z


# ============================================================
# DEBUG-ONLY START
# Helpers usados apenas para os arquivos de debug.
# Pode remover este bloco quando o problema estiver resolvido.
# ============================================================
def _expected_from_ct(ct: pd.DataFrame) -> np.ndarray:
    O = ct.to_numpy(dtype=float)
    total = float(np.nansum(O))
    if not np.isfinite(total) or total <= 0:
        return np.full_like(O, np.nan)

    row_sum = np.nansum(O, axis=1)
    col_sum = np.nansum(O, axis=0)
    E = np.outer(row_sum, col_sum) / total
    E[~np.isfinite(E)] = np.nan
    return E


def _safe_label_from_map(map_: Dict[Any, str], code: Any) -> str:
    return _safe_str(map_.get(code, _safe_str(code)))
# ============================================================
# DEBUG-ONLY END
# ============================================================

# ======================= FIM DO CÓDIGO ORIGINAL (residuals_page.py) =======================
