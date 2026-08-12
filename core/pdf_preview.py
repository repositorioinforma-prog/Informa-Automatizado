"""
Pré-visualização em PDF do relatório inteiro (com paginação e quebras de
página de verdade — o mais próximo de "abrir o PDF final") usando
wkhtmltopdf (via pdfkit) pra converter o HTML gerado a partir da
planilha.

Escolhido em vez do LibreOffice por ser bem mais leve — um binário só,
focado em conversão HTML->PDF, em vez de uma suíte de escritório
inteira — o que é mais adequado pro Streamlit Community Cloud (build
mais rápido, menos uso de recursos).

Precisa do pacote "wkhtmltopdf" instalado no sistema (ver packages.txt
na raiz do projeto — é isso que faz o Streamlit Community Cloud
instalar ele automaticamente via apt no deploy — um build completo,
não só um reboot, é necessário pra pegar um pacote de sistema novo).

Rodando localmente no Windows, precisa do wkhtmltopdf instalado na
máquina (https://wkhtmltopdf.org/downloads.html) — e, assim como o
LibreOffice, o instalador não garante que o binário fique no PATH,
então procuramos também no caminho de instalação padrão do Windows.

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import os
import shutil

_CAMINHOS_WINDOWS = [
    r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
    r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
]


def _encontrar_binario_wkhtmltopdf():
    """Acha o executável do wkhtmltopdf: primeiro no PATH, depois nos
    caminhos de instalação padrão do Windows como reforço."""
    encontrado = shutil.which("wkhtmltopdf")
    if encontrado:
        return encontrado
    for caminho in _CAMINHOS_WINDOWS:
        if os.path.isfile(caminho):
            return caminho
    return None


def wkhtmltopdf_disponivel():
    """Confere rapidamente se o binário do wkhtmltopdf existe no sistema
    (no PATH, ou nos caminhos de instalação padrão do Windows)."""
    return _encontrar_binario_wkhtmltopdf() is not None


def gerar_pdf_preview(bytes_xlsx, aba=None, timeout=90):
    """
    Converte os bytes de um .xlsx pra PDF paginado — abre a planilha com
    openpyxl, gera um HTML fiel (mesclagem, cor, fonte, borda, quebra de
    página) via `core.planilha_utils.worksheet_para_html_paginado`, e
    converte esse HTML pra PDF com wkhtmltopdf.

    Args:
        bytes_xlsx: conteúdo do arquivo .xlsx.
        aba: nome da aba a converter. Se None, usa a primeira.
        timeout: tempo máximo (segundos) de espera pela conversão.

    Returns:
        bytes do PDF gerado.

    Raises:
        RuntimeError: se o wkhtmltopdf não estiver instalado ou se a
            conversão falhar — sempre com uma mensagem clara do que
            aconteceu, pra exibir na tela em vez de travar o app.
    """
    import io

    import openpyxl
    import pdfkit

    binario = _encontrar_binario_wkhtmltopdf()
    if not binario:
        raise RuntimeError(
            "wkhtmltopdf não está instalado neste ambiente (ou não está "
            "no PATH nem nos caminhos padrão de instalação do Windows) — "
            "a pré-visualização em PDF não está disponível aqui."
        )

    wb = openpyxl.load_workbook(io.BytesIO(bytes_xlsx), rich_text=True)
    ws = wb[aba] if aba else wb[wb.sheetnames[0]]

    from core.planilha_utils import worksheet_para_html_paginado
    html = worksheet_para_html_paginado(ws)

    config = pdfkit.configuration(wkhtmltopdf=binario)
    opcoes = {
        "quiet": "",
        "disable-smart-shrinking": "",
        "print-media-type": "",
    }

    try:
        pdf_bytes = pdfkit.from_string(html, False, configuration=config, options=opcoes)
    except OSError as e:
        raise RuntimeError(f"Falha ao converter pra PDF: {e}")

    return pdf_bytes
