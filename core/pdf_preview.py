"""
Pré-visualização em PDF do relatório inteiro (com paginação, margens e
quebras de página de verdade — o mais próximo de "abrir o PDF final")
usando o LibreOffice em modo headless pra converter o .xlsx.

Precisa do pacote "libreoffice" instalado no sistema (ver packages.txt
na raiz do projeto — é isso que faz o Streamlit Community Cloud instalar
ele automaticamente no deploy). Se não estiver disponível, as funções
aqui avisam claramente em vez de travar o resto do app.

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import os
import shutil
import subprocess
import tempfile


def libreoffice_disponivel():
    """Confere rapidamente se o binário do LibreOffice existe no sistema."""
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def gerar_pdf_preview(bytes_xlsx, timeout=90):
    """
    Converte os bytes de um .xlsx pra PDF usando o LibreOffice headless.

    Args:
        bytes_xlsx: conteúdo do arquivo .xlsx.
        timeout: tempo máximo (segundos) de espera pela conversão —
            relatórios grandes (milhares de linhas, várias páginas)
            podem demorar; 90s costuma ser suficiente, mas em arquivos
            enormes pode precisar de mais.

    Returns:
        bytes do PDF gerado.

    Raises:
        RuntimeError: se o LibreOffice não estiver instalado, se a
            conversão estourar o tempo limite, ou se o processo falhar
            por qualquer outro motivo — sempre com uma mensagem clara
            do que aconteceu, pra exibir na tela em vez de travar o app.
    """
    binario = shutil.which("soffice") or shutil.which("libreoffice")
    if not binario:
        raise RuntimeError(
            "LibreOffice não está instalado neste ambiente — a "
            "pré-visualização em PDF não está disponível aqui."
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        caminho_entrada = os.path.join(tmpdir, "relatorio.xlsx")
        with open(caminho_entrada, "wb") as f:
            f.write(bytes_xlsx)

        # cada chamada usa um perfil de usuário isolado (-env:UserInstallation)
        # pra não dar conflito se duas conversões rodarem "ao mesmo tempo"
        # (ex.: duas pessoas usando o app simultaneamente)
        perfil = os.path.join(tmpdir, "lo_profile")
        try:
            resultado = subprocess.run(
                [
                    binario, "--headless", "--invisible", "--nocrashreport",
                    "--nodefault", "--norestore", "--nologo", "--nofirststartwizard",
                    f"-env:UserInstallation=file://{perfil}",
                    "--convert-to", "pdf", "--outdir", tmpdir, caminho_entrada,
                ],
                capture_output=True, timeout=timeout,
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError(
                f"A conversão pra PDF demorou mais que {timeout}s — o "
                "relatório pode ser grande demais pra essa pré-visualização."
            )

        caminho_pdf = os.path.join(tmpdir, "relatorio.pdf")
        if resultado.returncode != 0 or not os.path.exists(caminho_pdf):
            detalhe = resultado.stderr.decode("utf-8", errors="replace")[-500:]
            raise RuntimeError(f"Falha ao converter pra PDF: {detalhe or 'motivo desconhecido'}")

        with open(caminho_pdf, "rb") as f:
            return f.read()
