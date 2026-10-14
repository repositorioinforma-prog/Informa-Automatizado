import os
import tempfile
from core.pyreadstat_seguro import pyreadstat  # [Informa] lê .sav mesmo com a DLL bloqueada


def read_sav_bytes(data: bytes):
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = os.path.join(tmp_dir, "input.sav")
        with open(path, "wb") as handle:
            handle.write(data)
        return pyreadstat.read_sav(path, apply_value_formats=False)


def variable_catalog(df, meta):
    labels = getattr(meta, "column_names_to_labels", {}) or {}
    value_labels = getattr(meta, "variable_value_labels", {}) or {}
    return [{
        "variavel": c,
        "label": labels.get(c, "") or "",
        "tipo_pandas": str(df[c].dtype),
        "n_preenchidos": int(df[c].notna().sum()),
        "n_value_labels": len(value_labels.get(c, {})),
    } for c in df.columns]
