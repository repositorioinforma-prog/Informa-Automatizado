"""
Insere uma imagem no cabeçalho/rodapé de impressão de uma planilha
(o logo que aparece repetido em toda página impressa/exportada em PDF).

O openpyxl NÃO tem suporte nativo pra isso: imagem em cabeçalho/rodapé de
impressão usa um mecanismo legado do Excel baseado em VML (o mesmo tipo
de desenho usado por comentários de célula antigos), completamente
separado do sistema de imagens "soltas" na planilha que o openpyxl sabe
lidar (`ws.add_image()`). Por isso esta função opera diretamente nos
bytes do `.xlsx` já salvo (um arquivo `.xlsx` é um `.zip`), adicionando
as partes que faltam:
    - a imagem em xl/media/
    - um desenho VML (xl/drawings/vmlDrawingN.vml) descrevendo a
      posição/tamanho da imagem
    - as relações (.rels) ligando planilha -> VML -> imagem
    - a tag <legacyDrawingHF> na planilha, apontando pro VML
    - o texto do cabeçalho/rodapé com o código "&G" (placeholder de
      imagem do Excel) na posição desejada

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import io
import re
import zipfile

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CM_PARA_PX = 37.7952755906  # 1 cm = 37.7952755906 px a 96 DPI (padrão Windows/Excel)

_TIPO_CONTEUDO_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"
_TIPO_RELACAO_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
_TIPO_RELACAO_IMAGEM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

_EXTENSAO_CONTENT_TYPE = {
    "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "png": "image/png", "gif": "image/gif", "bmp": "image/bmp",
}


def _proximo_indice_disponivel(nomes_zip, prefixo, sufixo):
    """Acha o próximo número livre pra um nome tipo 'xl/media/image{N}.jpg'."""
    usados = set()
    padrao = re.compile(re.escape(prefixo) + r"(\d+)" + re.escape(sufixo) + r"$")
    for nome in nomes_zip:
        m = padrao.match(nome)
        if m:
            usados.add(int(m.group(1)))
    n = 1
    while n in usados:
        n += 1
    return n


def _proximo_rid(xml_rels):
    """Acha o próximo Id='rIdN' livre num arquivo .rels existente (texto XML ou None)."""
    if not xml_rels:
        return 1
    usados = [int(m) for m in re.findall(r'Id="rId(\d+)"', xml_rels)]
    n = 1
    while n in usados:
        n += 1
    return n


def _sheet_path_e_rid(zin, nomes_zip, nome_aba):
    """
    Descobre o caminho do XML da aba (ex.: 'xl/worksheets/sheet2.xml') e o
    r:id correspondente, a partir do nome da aba — segue workbook.xml
    (nome -> r:id) e workbook.xml.rels (r:id -> caminho do arquivo).
    """
    workbook_xml = zin.read("xl/workbook.xml").decode("utf-8")
    m = re.search(
        r'<sheet\b[^>]*\bname="' + re.escape(nome_aba) + r'"[^>]*\br:id="(rId\d+)"',
        workbook_xml,
    )
    if not m:
        # tenta a ordem inversa dos atributos (name depois de r:id)
        m = re.search(
            r'<sheet\b[^>]*\br:id="(rId\d+)"[^>]*\bname="' + re.escape(nome_aba) + r'"',
            workbook_xml,
        )
    if not m:
        raise ValueError(f"Aba '{nome_aba}' não encontrada em workbook.xml")
    rid_wb = m.group(1)

    rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m2 = re.search(r'<Relationship\b[^>]*\bId="' + rid_wb + r'"[^>]*\bTarget="([^"]+)"', rels_xml)
    if not m2:
        m2 = re.search(r'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="' + rid_wb + r'"', rels_xml)
    if not m2:
        raise ValueError(f"Relacionamento '{rid_wb}' não encontrado em workbook.xml.rels")
    alvo = m2.group(1).lstrip("/")
    if not alvo.startswith("xl/"):
        alvo = "xl/" + alvo
    return alvo


def inserir_imagem_cabecalho(
    bytes_xlsx, caminho_imagem, largura_cm=5.0, posicao="R", aba=None,
    aplicar_em_todas_abas=False,
):
    """
    Devolve uma nova versão (bytes) do .xlsx com a imagem inserida no
    cabeçalho de impressão — repete automaticamente em toda página, é
    assim que cabeçalho/rodapé de impressão funciona no Excel.

    Args:
        bytes_xlsx: conteúdo do .xlsx (já salvo pelo openpyxl) como bytes.
        caminho_imagem: caminho no disco pra imagem (jpg/png/etc.).
        largura_cm: largura desejada da imagem no cabeçalho, em cm — a
            altura é calculada automaticamente pra manter a proporção
            original da imagem.
        posicao: "L" (esquerda), "C" (centro) ou "R" (direita) — em qual
            parte do cabeçalho a imagem entra. Não mexe no texto que já
            estiver nas OUTRAS posições (ex.: o aviso legal do código 14
            fica em "L", a imagem pode ir em "R" sem conflito).
        aba: nome da aba a receber a imagem. Se None, usa a primeira aba
            do arquivo.
        aplicar_em_todas_abas: se True, insere a mesma imagem no
            cabeçalho de TODAS as abas do arquivo (cada uma com sua
            própria cópia do desenho VML — o Excel não permite
            compartilhar um VML entre abas diferentes).

    Returns:
        bytes do novo .xlsx.
    """
    from PIL import Image as _PILImage

    with open(caminho_imagem, "rb") as f:
        dados_imagem = f.read()
    extensao = caminho_imagem.rsplit(".", 1)[-1].lower()
    if extensao not in _EXTENSAO_CONTENT_TYPE:
        raise ValueError(f"Extensão de imagem não suportada: {extensao}")

    with _PILImage.open(io.BytesIO(dados_imagem)) as img:
        img_largura_px, img_altura_px = img.size
    altura_cm = largura_cm * (img_altura_px / img_largura_px)
    largura_px_alvo = largura_cm * CM_PARA_PX
    altura_px_alvo = altura_cm * CM_PARA_PX

    buf_entrada = io.BytesIO(bytes_xlsx)
    with zipfile.ZipFile(buf_entrada) as zin:
        nomes_zip = zin.namelist()
        conteudos = {nome: zin.read(nome) for nome in nomes_zip}

        content_types = conteudos["[Content_Types].xml"].decode("utf-8")

        # abas alvo
        if aplicar_em_todas_abas:
            workbook_xml = conteudos["xl/workbook.xml"].decode("utf-8")
            nomes_abas = re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', workbook_xml)
        else:
            if aba is None:
                workbook_xml = conteudos["xl/workbook.xml"].decode("utf-8")
                nomes_abas = [re.search(r'<sheet\b[^>]*\bname="([^"]+)"', workbook_xml).group(1)]
            else:
                nomes_abas = [aba]

        # imagem só precisa ser adicionada uma vez, mesmo se usada em várias abas
        idx_img = _proximo_indice_disponivel(conteudos.keys(), "xl/media/image", f".{extensao}")
        caminho_img = f"xl/media/image{idx_img}.{extensao}"
        conteudos[caminho_img] = dados_imagem
        if 'ContentType="' + _EXTENSAO_CONTENT_TYPE[extensao] not in content_types:
            if f'Extension="{extensao}"' not in content_types:
                content_types = content_types.replace(
                    "</Types>",
                    f'<Default Extension="{extensao}" '
                    f'ContentType="{_EXTENSAO_CONTENT_TYPE[extensao]}"/></Types>',
                )
        if _TIPO_CONTEUDO_VML not in content_types:
            content_types = content_types.replace(
                "</Types>",
                f'<Default Extension="vml" ContentType="{_TIPO_CONTEUDO_VML}"/></Types>',
            )

        for nome_aba in nomes_abas:
            caminho_sheet = _sheet_path_e_rid(zin, nomes_zip, nome_aba)
            sheet_xml = conteudos[caminho_sheet].decode("utf-8")

            # ---- VML drawing (um por aba) ----
            idx_vml = _proximo_indice_disponivel(conteudos.keys(), "xl/drawings/vmlDrawing", ".vml")
            caminho_vml = f"xl/drawings/vmlDrawing{idx_vml}.vml"
            vml_xml = f"""<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"
  o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
  <v:stroke joinstyle="miter"/>
  <v:formulas>
   <v:f eqn="if lineDrawn pixelLineWidth 0"/>
   <v:f eqn="sum @0 1 0"/>
   <v:f eqn="sum 0 0 @1"/>
   <v:f eqn="prod @2 1 2"/>
   <v:f eqn="prod @3 21600 pixelWidth"/>
   <v:f eqn="prod @3 21600 pixelHeight"/>
   <v:f eqn="sum @0 0 1"/>
   <v:f eqn="prod @6 1 2"/>
   <v:f eqn="prod @7 21600 pixelWidth"/>
   <v:f eqn="sum @8 21600 0"/>
   <v:f eqn="prod @7 21600 pixelHeight"/>
   <v:f eqn="sum @10 21600 0"/>
  </v:formulas>
  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" aspectratio="t"/>
 </v:shapetype><v:shape id="imagem_cabecalho_{idx_vml}" o:spid="_x0000_s{idx_vml}025" \
type="#_x0000_t75" style="position:absolute;\
margin-left:0px;margin-top:0px;width:{largura_px_alvo:.0f}px;height:{altura_px_alvo:.0f}px;z-index:1">
  <v:imagedata o:relid="rId1" o:title=""/>
  <o:lock v:ext="edit" textRotation="t"/>
 </v:shape>
</xml>"""
            conteudos[caminho_vml] = vml_xml.encode("utf-8")

            caminho_vml_rels = f"xl/drawings/_rels/vmlDrawing{idx_vml}.vml.rels"
            conteudos[caminho_vml_rels] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{_TIPO_RELACAO_IMAGEM}" '
                f'Target="../media/image{idx_img}.{extensao}"/>'
                "</Relationships>"
            ).encode("utf-8")

            # ---- relação da PLANILHA -> VML ----
            pasta_sheet = caminho_sheet.rsplit("/", 1)[0]
            nome_sheet_arquivo = caminho_sheet.rsplit("/", 1)[-1]
            caminho_sheet_rels = f"{pasta_sheet}/_rels/{nome_sheet_arquivo}.rels"
            sheet_rels_xml = conteudos.get(caminho_sheet_rels, b"").decode("utf-8") or None
            rid_sheet = _proximo_rid(sheet_rels_xml)
            if sheet_rels_xml:
                novo_sheet_rels = sheet_rels_xml.replace(
                    "</Relationships>",
                    f'<Relationship Id="rId{rid_sheet}" Type="{_TIPO_RELACAO_VML}" '
                    f'Target="../drawings/vmlDrawing{idx_vml}.vml"/></Relationships>',
                )
            else:
                novo_sheet_rels = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId{rid_sheet}" Type="{_TIPO_RELACAO_VML}" '
                    f'Target="../drawings/vmlDrawing{idx_vml}.vml"/>'
                    "</Relationships>"
                )
            conteudos[caminho_sheet_rels] = novo_sheet_rels.encode("utf-8")

            # ---- <legacyDrawingHF> na planilha — o schema do Excel exige
            # uma ordem específica de elementos dentro de <worksheet>:
            # ...pageSetup, headerFooter, rowBreaks, colBreaks,
            # customProperties, cellWatches, ignoredErrors, smartTags,
            # drawing, legacyDrawingHF, ... Colocar logo depois de
            # </headerFooter> (ignorando rowBreaks/colBreaks, que o
            # código 07 já pode ter inserido) deixa o arquivo fora de
            # ordem — sintaticamente ainda é XML válido, mas o Excel
            # rejeita e "repara" removendo a planilha inteira, sem erro
            # nenhum durante a geração (só ao abrir de verdade no
            # Excel). Insere sempre depois do ÚLTIMO desses elementos que
            # existir no arquivo, nunca só depois de headerFooter.
            tag_legacy = f'<legacyDrawingHF xmlns:r="{NS_R}" r:id="rId{rid_sheet}"/>'
            ponto_insercao = None
            for tag_fim in ("</colBreaks>", "</rowBreaks>", "</headerFooter>"):
                idx = sheet_xml.rfind(tag_fim)
                if idx != -1:
                    ponto_insercao = idx + len(tag_fim)
                    break
            if ponto_insercao is not None:
                sheet_xml = sheet_xml[:ponto_insercao] + tag_legacy + sheet_xml[ponto_insercao:]
            else:
                sheet_xml = sheet_xml.replace("</worksheet>", tag_legacy + "</worksheet>")

            # ---- texto do cabeçalho: adiciona "&G" na posição pedida,
            # sem apagar o que já estiver nas outras posições ----
            marcador = {"L": "&amp;L", "C": "&amp;C", "R": "&amp;R"}[posicao]
            m_header = re.search(r"<oddHeader>(.*?)</oddHeader>", sheet_xml, re.DOTALL)
            if m_header:
                texto_atual = m_header.group(1)
                if marcador in texto_atual:
                    novo_texto = texto_atual.replace(marcador, marcador + "&amp;G", 1)
                else:
                    novo_texto = texto_atual + marcador + "&amp;G"
                sheet_xml = sheet_xml.replace(
                    f"<oddHeader>{texto_atual}</oddHeader>", f"<oddHeader>{novo_texto}</oddHeader>"
                )
            else:
                novo_header = f"<oddHeader>{marcador}&amp;G</oddHeader>"
                if "<headerFooter>" in sheet_xml:
                    sheet_xml = sheet_xml.replace("<headerFooter>", "<headerFooter>" + novo_header)
                else:
                    # não existe headerFooter nenhum ainda — cria um bloco mínimo
                    bloco = f"<headerFooter>{novo_header}</headerFooter>"
                    if "</worksheet>" in sheet_xml:
                        sheet_xml = sheet_xml.replace("</worksheet>", bloco + "</worksheet>")

            conteudos[caminho_sheet] = sheet_xml.encode("utf-8")

        conteudos["[Content_Types].xml"] = content_types.encode("utf-8")

    buf_saida = io.BytesIO()
    with zipfile.ZipFile(buf_saida, "w", zipfile.ZIP_DEFLATED) as zout:
        for nome, dados in conteudos.items():
            zout.writestr(nome, dados)

    return buf_saida.getvalue()
