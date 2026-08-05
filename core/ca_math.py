"""
Motor matemático da Análise de Correspondência (CA), via SVD manual.
Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import numpy as np
import pandas as pd


def _ca_manual(tabela_contingencia, n_components=2):
    """
    Implementação manual de Análise de Correspondência via SVD.
    """
    N = tabela_contingencia.to_numpy(dtype=float)
    n = N.sum()
    if n <= 0:
        raise ValueError("Tabela de contingência vazia.")

    P = N / n
    r = P.sum(axis=1)
    c = P.sum(axis=0)

    row_mask = r > 0
    col_mask = c > 0

    if row_mask.sum() < 2 or col_mask.sum() < 2:
        raise ValueError("Linhas/colunas insuficientes com massa > 0 para a CA.")

    P_rc = P[np.ix_(row_mask, col_mask)]
    r_rc = r[row_mask]
    c_rc = c[col_mask]

    D_r_inv_sqrt = np.diag(1.0 / np.sqrt(r_rc))
    D_c_inv_sqrt = np.diag(1.0 / np.sqrt(c_rc))

    expected = np.outer(r_rc, c_rc)
    S = D_r_inv_sqrt @ (P_rc - expected) @ D_c_inv_sqrt

    U, singvals, VT = np.linalg.svd(S, full_matrices=False)
    eigenvalues = singvals ** 2

    F = D_r_inv_sqrt @ U @ np.diag(singvals)
    G = D_c_inv_sqrt @ VT.T @ np.diag(singvals)

    k = min(n_components, F.shape[1])
    F = F[:, :k]
    G = G[:, :k]
    eigenvalues = eigenvalues[:k]

    row_index = tabela_contingencia.index[row_mask]
    col_index = tabela_contingencia.columns[col_mask]
    cols = [f"Dim{i+1}" for i in range(k)]

    coordenadas_linhas = pd.DataFrame(F, index=row_index, columns=cols)
    coordenadas_colunas = pd.DataFrame(G, index=col_index, columns=cols)

    return coordenadas_linhas, coordenadas_colunas, eigenvalues

