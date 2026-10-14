import os
import tempfile

from core.pyreadstat_seguro import pyreadstat  # [Informa] com a DLL bloqueada, grava em Python puro


def write_sav_bytes(df, column_labels: dict | None = None, variable_value_labels: dict | None = None,
                     file_label: str = "") -> bytes:
    """
    Grava o DataFrame como .sav e retorna os bytes do arquivo, prontos para
    download. Usa pyreadstat (mesma biblioteca usada na importação), então
    preserva corretamente compressão, valores ausentes e variáveis string
    longas -- diferente de escrever um .sav "na unha".

    column_labels: nome_da_coluna -> label da variável (dicionário; ordem não importa).
    variable_value_labels: nome_da_coluna -> {codigo: label} (VALUE LABELS).
    """
    column_labels = column_labels or {}
    variable_value_labels = variable_value_labels or {}
    if pyreadstat.usando_leitor_puro:
        # [Informa] DLL do pyreadstat bloqueada no Windows: grava pelo escritor
        # em Python puro (exporters/sav_sem_pyreadstat.py, sem compressão)
        from controle.exporters.sav_sem_pyreadstat import write_sav_bytes as _gravar_puro
        dados, _avisos = _gravar_puro(
            df, {}, var_labels=column_labels,
            value_labels={v: dict(t) for v, t in variable_value_labels.items() if t},
            file_label=file_label or "Banco controlado - Gerador de Controle Geral",
        )
        return dados
    # pyreadstat exige dict puro (não aceita OrderedDict nem outros mapeamentos),
    # tanto no nível externo quanto no dict de código->label de cada variável.
    variable_value_labels = {
        var: dict(table) for var, table in variable_value_labels.items() if table
    }

    # pyreadstat espera os labels de variável como uma lista alinhada à ordem
    # das colunas do DataFrame, não como um dicionário.
    labels_list = [column_labels.get(col, "") for col in df.columns]

    with tempfile.TemporaryDirectory() as tmp_dir:
        path = os.path.join(tmp_dir, "output.sav")
        pyreadstat.write_sav(
            df,
            path,
            file_label=file_label or "Banco controlado - Gerador de Controle Geral",
            column_labels=labels_list,
            variable_value_labels=variable_value_labels,
            compress=True,
        )
        with open(path, "rb") as handle:
            return handle.read()
