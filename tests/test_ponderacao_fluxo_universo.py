from pathlib import Path


def test_fluxo_do_universo_segue_selecao_download_importacao():
    """A interface deve seguir o fluxo natural pedido pelo usuário."""
    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    fonte = caminho.read_text(encoding="utf-8")

    pos_selecao = fonte.index('variaveis = st.multiselect(')
    pos_download = fonte.index('key="ponderacao_download_modelo_universo"')
    pos_upload = fonte.index('key="ponderacao_upload_universo_xlsx"')

    assert pos_selecao < pos_download < pos_upload


def test_download_e_upload_do_universo_ficam_na_mesma_secao():
    caminho = Path(__file__).resolve().parents[1] / "analises" / "ponderacao.py"
    fonte = caminho.read_text(encoding="utf-8")

    inicio_secao = fonte.index('st.subheader("2. Modelo e importação do Universo")')
    fim_secao = fonte.index('st.subheader("3. Componentes de peso")', inicio_secao)
    secao = fonte[inicio_secao:fim_secao]

    assert 'Baixar modelo do Universo (.xlsx)' in secao
    assert 'Importar Universo preenchido (.xlsx)' in secao
    assert 'Fluxo recomendado: selecione as variáveis acima, baixe o modelo' in secao
