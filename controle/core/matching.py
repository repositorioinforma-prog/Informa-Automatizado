import difflib
import re
from dataclasses import dataclass
from .normalization import normalize_text


@dataclass(frozen=True)
class MatchResult:
    code: float | int | None
    label: str | None
    method: str
    score: float
    candidates: tuple[float | int, ...] = ()


def build_lookup(label_dict: dict):
    by_norm = {}
    for code, label in label_dict.items():
        key = normalize_text(label)
        if key:
            by_norm.setdefault(key, []).append(code)
    return by_norm, list(by_norm)


def _single_code(keys, by_norm):
    codes = []
    for key in keys:
        codes.extend(by_norm[key])
    unique = list(dict.fromkeys(codes))
    return (unique[0] if len(unique) == 1 else None), unique


def match_text(text, label_dict: dict, cutoff: float = 0.82) -> MatchResult:
    tn = normalize_text(text)
    if not tn:
        return MatchResult(None, None, "vazio", 0.0)
    by_norm, norm_list = build_lookup(label_dict)
    if tn in by_norm:
        code, candidates = _single_code([tn], by_norm)
        if code is not None:
            return MatchResult(code, label_dict[code], "exata", 1.0)
        return MatchResult(None, None, "ambiguo", 0.0, tuple(candidates))
    starts = [k for k in norm_list if k.startswith(tn + " ")]
    if starts:
        code, candidates = _single_code(starts, by_norm)
        if code is not None:
            return MatchResult(code, label_dict[code], "prefixo", 0.95)
        return MatchResult(None, None, "ambiguo", 0.0, tuple(candidates))
    contains = [k for k in norm_list if re.search(r"\b" + re.escape(tn) + r"\b", k)]
    if contains:
        code, candidates = _single_code(contains, by_norm)
        if code is not None:
            return MatchResult(code, label_dict[code], "substring", 0.90)
        return MatchResult(None, None, "ambiguo", 0.0, tuple(candidates))
    # Direção inversa: a resposta traz texto extra em volta do nome válido, ex.:
    # "CAPITAL/Rio de Janeiro" contendo o label do projeto "Rio de Janeiro".
    contains_reverse = [k for k in norm_list if k and re.search(r"\b" + re.escape(k) + r"\b", tn)]
    if contains_reverse:
        # entre rótulos que aparecem dentro do texto, prioriza o mais longo (mais específico)
        contains_reverse.sort(key=len, reverse=True)
        best_len = len(contains_reverse[0])
        top = [k for k in contains_reverse if len(k) == best_len]
        code, candidates = _single_code(top, by_norm)
        if code is not None:
            return MatchResult(code, label_dict[code], "substring_invertido", 0.88)
        return MatchResult(None, None, "ambiguo", 0.0, tuple(candidates))
    close = difflib.get_close_matches(tn, norm_list, n=3, cutoff=cutoff)
    if close:
        best = close[0]
        score = difflib.SequenceMatcher(None, tn, best).ratio()
        candidates = []
        for key in close:
            candidates.extend(by_norm[key])
        candidates = list(dict.fromkeys(candidates))
        if len(by_norm[best]) == 1:
            code = by_norm[best][0]
            return MatchResult(code, label_dict[code], "fuzzy", score, tuple(candidates[1:]))
        return MatchResult(None, None, "ambiguo", score, tuple(candidates))
    return MatchResult(None, None, "sem_correspondencia", 0.0)


def confidence_status(result: MatchResult, auto_fuzzy_threshold: float = 0.93) -> str:
    if result.method == "vazio": return "VAZIO"
    if result.method == "ambiguo": return "AMBÍGUO"
    if result.method == "sem_correspondencia": return "NÃO IDENTIFICADO"
    if result.method == "exata": return "CONFIRMADO"
    if result.method in {"prefixo", "substring", "substring_invertido"}: return "AJUSTADO AUTOMATICAMENTE"
    if result.method == "fuzzy" and result.score >= auto_fuzzy_threshold: return "AJUSTADO AUTOMATICAMENTE"
    return "SUGESTÃO"
