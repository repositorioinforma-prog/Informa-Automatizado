"""
Matriz de Correlação — cálculos do app desktop "Relatória Analítica".

O trecho entre as marcas "INÍCIO/FIM DO CÓDIGO ORIGINAL" é cópia SEM
ALTERAÇÕES de app/pages/correlation_page.py (linhas 24–298): Pearson e
Spearman ponderados, Sig. (2-tailed), filtro 0,000–0,005 e cores do VBA.
As contas dos @Slot (computeCorr, getUniqueValues e exportações) estão em
analises/correlacao.py, também sem mudar as regras.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# ===================== INÍCIO DO CÓDIGO ORIGINAL (correlation_page.py) =====================

try:
    from scipy import stats as scipy_stats
except Exception:
    scipy_stats = None




def _canon(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = (
        s.replace("á", "a").replace("à", "a").replace("â", "a").replace("ã", "a")
         .replace("é", "e").replace("ê", "e")
         .replace("í", "i")
         .replace("ó", "o").replace("ô", "o").replace("õ", "o")
         .replace("ú", "u").replace("ç", "c")
    )
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def _is_peso_name(name: str) -> bool:
    return "peso" in _canon(name)


def _numeric_candidates(df: pd.DataFrame) -> List[str]:
    cols: List[str] = []
    n = len(df)

    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            cols.append(str(c))
            continue

        if pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s):
            x = pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")
            if x.notna().sum() >= max(3, int(0.2 * n)):
                cols.append(str(c))

    return sorted(cols, key=lambda x: str(x).lower())


def _parse_missing_list(text: str) -> Tuple[List[str], List[float]]:
    tokens = [t.strip() for t in (text or "").split(",") if t.strip()]
    missing_str: List[str] = []
    missing_num: List[float] = []

    for t in tokens:
        tt = t.replace(".", "").replace(",", ".") if re.match(r"^[0-9\.\,]+$", t) else t
        try:
            v = float(tt)
            missing_num.append(v)
        except Exception:
            missing_str.append(t.strip().lower())

    return missing_str, missing_num


def _looks_numeric_token(t: str) -> bool:
    if t is None:
        return False
    s = str(t).strip()
    if not s:
        return False
    return bool(re.match(r"^[\+\-]?\d+([\,\.]\d+)?$", s))


def _token_to_float(t: str) -> Optional[float]:
    try:
        s = str(t).strip().replace(",", ".")
        return float(s)
    except Exception:
        return None


def _weighted_pearson(x: np.ndarray, y: np.ndarray, w: np.ndarray) -> float:
    sw = float(np.sum(w))
    if not np.isfinite(sw) or sw <= 0:
        return float("nan")

    mx = float(np.sum(w * x) / sw)
    my = float(np.sum(w * y) / sw)

    dx = x - mx
    dy = y - my

    vx = float(np.sum(w * dx * dx) / sw)
    vy = float(np.sum(w * dy * dy) / sw)
    if vx <= 0 or vy <= 0:
        return float("nan")

    cov = float(np.sum(w * dx * dy) / sw)
    return float(cov / np.sqrt(vx * vy))


SIG_P_MIN = 0.0
SIG_P_MAX = 0.005


def _sig_keep_value(p: float) -> bool:
    return p is not None and np.isfinite(p) and (SIG_P_MIN <= float(p) <= SIG_P_MAX)


def _sig_style_rgb(p: float) -> Optional[Tuple[str, str]]:
    """
    Estilo da folha Sig. (2-tailed).
    Destaca apenas os p-values que passam no critério 0,000 a 0,005.
    """
    if not _sig_keep_value(p):
        return None
    return ("D9EAD3", "274E13")  # verde claro / verde escuro


def _two_tailed_p_from_r(r: float, n: float) -> float:
    """
    p-valor bicaudal a partir de r e N.
    Aqui foi usado N ponderado como soma dos pesos válidos do par.
    """
    if scipy_stats is None:
        return float("nan")

    if r is None or not np.isfinite(r):
        return float("nan")
    if n is None or not np.isfinite(n) or float(n) <= 2.0:
        return float("nan")

    rr = float(r)
    rr = max(min(rr, 0.999999999999), -0.999999999999)

    den = 1.0 - (rr * rr)
    if den <= 0:
        return 0.0

    t = rr * np.sqrt((float(n) - 2.0) / den)
    return float(2.0 * scipy_stats.t.sf(abs(t), df=float(n) - 2.0))


def _weighted_corr_and_sig_matrices(
    df: pd.DataFrame,
    cols: List[str],
    w: pd.Series,
    method: str
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      corr_df      -> matriz de correlação "bruta"
      sig_df       -> matriz de Sig. (2-tailed)
      filtered_df  -> matriz da correlação já filtrada pelo p-value
    """
    method = (method or "pearson").lower().strip()
    n = len(cols)

    mat = np.full((n, n), np.nan, dtype=float)
    pmat = np.full((n, n), np.nan, dtype=float)
    filtered = np.full((n, n), np.nan, dtype=float)

    wv = pd.to_numeric(w, errors="coerce").fillna(0.0).astype(float)

    for i in range(n):
        mat[i, i] = 1.0
        filtered[i, i] = 1.0
        pmat[i, i] = np.nan

    for i in range(n):
        xi = pd.to_numeric(df[cols[i]], errors="coerce")

        for j in range(i + 1, n):
            yj = pd.to_numeric(df[cols[j]], errors="coerce")

            mask = xi.notna() & yj.notna() & (wv > 0)
            if int(mask.sum()) < 2:
                continue

            ww = wv[mask].to_numpy(dtype=float)
            xx = xi[mask]
            yy = yj[mask]

            if method == "spearman":
                xx = xx.rank(method="average")
                yy = yy.rank(method="average")

            r = _weighted_pearson(xx.to_numpy(dtype=float), yy.to_numpy(dtype=float), ww)

            # N ponderado do par
            n_pair = float(np.sum(ww))
            p = _two_tailed_p_from_r(r, n_pair)

            mat[i, j] = r
            mat[j, i] = r

            pmat[i, j] = p
            pmat[j, i] = p

            if _sig_keep_value(p):
                filtered[i, j] = r
                filtered[j, i] = r

    return (
        pd.DataFrame(mat, index=cols, columns=cols),
        pd.DataFrame(pmat, index=cols, columns=cols),
        pd.DataFrame(filtered, index=cols, columns=cols),
    )


def _corr_style_rgb(v: float) -> Optional[Tuple[str, str]]:
    if v is None or not np.isfinite(v):
        return None
    if abs(float(v) - 1.0) < 1e-12:
        return None

    if v <= -0.5:
        return ("BE5014", "FFFFFF")
    if v <= -0.4:
        return ("F1A983", "BE5014")
    if v <= -0.3:
        return ("F7C7AC", "BE5014")
    if v <= -0.2:
        return ("FBE2D5", "BE5014")

    if v >= 0.5:
        return ("104861", "FFFFFF")
    if v >= 0.4:
        return ("44B3E1", "104861")
    if v >= 0.3:
        return ("83CCEB", "104861")
    if v >= 0.2:
        return ("C0E6F5", "104861")

    return None


def _format_token_for_ui(v: Any) -> Tuple[str, str]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ("__NA__", "⟂ (NA)")
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return ("__EMPTY__", "<vazio>")
        return (s, s)

    try:
        if isinstance(v, (np.integer, int)):
            s = str(int(v))
            return (s, s)
        if isinstance(v, (np.floating, float)):
            if float(v).is_integer():
                s = str(int(v))
            else:
                s = f"{float(v):g}"
            return (s, s)
    except Exception:
        pass

    s = str(v).strip()
    if s == "":
        return ("__EMPTY__", "<vazio>")
    return (s, s)

def _jsonable_matrix(mat: np.ndarray) -> List[List[Optional[float]]]:
    out: List[List[Optional[float]]] = []
    arr = np.asarray(mat, dtype=float)

    for row in arr:
        new_row: List[Optional[float]] = []
        for v in row:
            if np.isfinite(v):
                new_row.append(float(v))
            else:
                new_row.append(None)
        out.append(new_row)

    return out
# ======================= FIM DO CÓDIGO ORIGINAL (correlation_page.py) =======================
