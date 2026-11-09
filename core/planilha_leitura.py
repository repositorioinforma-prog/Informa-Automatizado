"""
Leitura de planilhas de resultados em .xlsx ou .xls (lógica pura, sem Streamlit).

O SPSS/Excel do Instituto entrega parte dos arquivos em .xls (formato antigo). Ordem de tentativa:
  1. .xlsx / .xlsm  -> openpyxl;
  2. .xls           -> xlrd (está no requirements.txt);
  3. .xls sem xlrd  -> LibreOffice (soffice), se existir na máquina;
  4. nada disso     -> ErroLeitura com a instrução de salvar o arquivo como .xlsx.
"""
from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
from typing import Any, List

import openpyxl

ASSINATURA_XLS = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # OLE2 (formato .xls antigo)
ASSINATURA_ZIP = b"PK\x03\x04"                          # .xlsx / .pptx / .zip


class ErroLeitura(Exception):
    """Arquivo ilegível (mensagem já pronta para mostrar à pessoa)."""


def eh_xls(conteudo: bytes) -> bool:
    return conteudo[:8] == ASSINATURA_XLS


def eh_zip(conteudo: bytes) -> bool:
    return conteudo[:4] == ASSINATURA_ZIP


def _linhas_xlrd(conteudo: bytes) -> List[List[Any]]:
    import xlrd  # noqa: WPS433 (import tardio: dependência opcional)

    wb = xlrd.open_workbook(file_contents=conteudo)
    ws = wb.sheet_by_index(0)
    saida = []
    for i in range(ws.nrows):
        linha = []
        for v in ws.row_values(i):
            linha.append(None if v == "" else v)
        saida.append(linha)
    return saida


def _converter_com_soffice(conteudo: bytes) -> bytes:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise FileNotFoundError("soffice")
    with tempfile.TemporaryDirectory() as pasta:
        origem = os.path.join(pasta, "entrada.xls")
        with open(origem, "wb") as f:
            f.write(conteudo)
        subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", pasta, origem],
                       check=True, capture_output=True, timeout=120)
        with open(os.path.join(pasta, "entrada.xlsx"), "rb") as f:
            return f.read()


def _linhas_openpyxl(conteudo: bytes) -> List[List[Any]]:
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def linhas_da_planilha(conteudo: bytes) -> List[List[Any]]:
    """Linhas (listas de valores) da primeira aba. Levanta ErroLeitura se não der para ler."""
    try:
        if eh_xls(conteudo):
            try:
                return _linhas_xlrd(conteudo)
            except ImportError:
                try:
                    return _linhas_openpyxl(_converter_com_soffice(conteudo))
                except (FileNotFoundError, subprocess.SubprocessError, OSError):
                    raise ErroLeitura(
                        "Este arquivo é .xls (formato antigo) e o leitor não está instalado. "
                        "Instale com «pip install xlrd» ou abra o arquivo no Excel e salve como .xlsx."
                    ) from None
        return _linhas_openpyxl(conteudo)
    except ErroLeitura:
        raise
    except Exception as exc:  # arquivo corrompido, protegido, não é planilha...
        raise ErroLeitura(f"Não consegui abrir a planilha ({type(exc).__name__}).") from exc
