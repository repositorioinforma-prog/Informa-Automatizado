"""Exportação Excel do Gerador de Amostra.

O arquivo final contém:
    - aba ``Informações`` com metadados, memória de cálculo e validações;
    - aba ``Amostra Completa``;
    - controles por camada, quando aplicável;
    - uma aba por campanha, quando configurada.

As cotas demográficas podem ser exibidas como totais, percentuais dentro de
cada unidade ou nos dois formatos. As colunas de amostra total e percentual da
amostra permanecem sempre visíveis.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
COR_CABECALHO_ESCURO = "1F4E78"
COR_CABECALHO_VERDE = "70AD47"
COR_CABECALHO_TEAL = "0F6B78"
COR_SECAO = "D9EAF7"
COR_LINHA_ZEBRA = "F2F2F2"
COR_ALERTA = "FCE4D6"
COR_OK = "E2F0D9"
COR_BORDA = "808080"
COR_TEXTO_CINZA = "666666"

FONTE_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONTE_CAMPANHA = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONTE_CORPO = Font(name="Calibri", size=10)
FONTE_TITULO = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
FONTE_SUBTITULO = Font(name="Calibri", size=11, bold=True, color=COR_CABECALHO_ESCURO)

FILL_HEADER = PatternFill(
    start_color=COR_CABECALHO_ESCURO,
    end_color=COR_CABECALHO_ESCURO,
    fill_type="solid",
)
FILL_CAMPANHA = PatternFill(
    start_color=COR_CABECALHO_VERDE,
    end_color=COR_CABECALHO_VERDE,
    fill_type="solid",
)
FILL_TEAL = PatternFill(
    start_color=COR_CABECALHO_TEAL,
    end_color=COR_CABECALHO_TEAL,
    fill_type="solid",
)
FILL_SECAO = PatternFill(start_color=COR_SECAO, end_color=COR_SECAO, fill_type="solid")
FILL_ZEBRA = PatternFill(
    start_color=COR_LINHA_ZEBRA,
    end_color=COR_LINHA_ZEBRA,
    fill_type="solid",
)
FILL_ALERTA = PatternFill(start_color=COR_ALERTA, end_color=COR_ALERTA, fill_type="solid")
FILL_OK = PatternFill(start_color=COR_OK, end_color=COR_OK, fill_type="solid")

ALINHA_CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALINHA_ESQUERDA = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALINHA_DIREITA = Alignment(horizontal="right", vertical="center")

_borda_fina = Side(border_style="thin", color=COR_BORDA)
BORDA = Border(top=_borda_fina, bottom=_borda_fina, left=_borda_fina, right=_borda_fina)
BORDA_INFERIOR = Border(bottom=_borda_fina)

FORMATOS_COTAS_VALIDOS = {"totais", "percentuais", "ambos"}
ROTULOS_FORMATO_COTAS = {
    "totais": "Somente totais (n)",
    "percentuais": "Somente percentuais (%)",
    "ambos": "Totais e percentuais",
}

CATEGORIAS_COTAS = [
    ("Gênero", "Masculino", "cota_masc"),
    ("Gênero", "Feminino", "cota_fem"),
    ("Idade", "16 a 19", "cota_16_19"),
    ("Idade", "20 a 29", "cota_20_29"),
    ("Idade", "30 a 39", "cota_30_39"),
    ("Idade", "40 a 49", "cota_40_49"),
    ("Idade", "50 ou mais", "cota_50mais"),
    ("Renda Média Domiciliar Familiar", "Até 2sm", "cota_renda_ate_2sm"),
    ("Renda Média Domiciliar Familiar", "2 a 5sm", "cota_renda_2_a_5sm"),
    ("Renda Média Domiciliar Familiar", "5 a 10sm", "cota_renda_5_a_10sm"),
    ("Renda Média Domiciliar Familiar", "Mais de 10sm", "cota_renda_mais_10sm"),
]

HIERARQUIA_EXPORTACAO = [
    ("campanha", "Campanhas", "Campanhas", 28),
    ("regiao_intermediaria", "Mesorregiões (Intermediárias)", "Mesorregiões (Intermediárias)", 24),
    ("regiao_imediata", "Microrregiões (Imediatas)", "Microrregiões (Imediatas)", 24),
    ("municipio", "Municípios", "Municípios", 24),
    ("distrito", "Distritos", "Distritos", 24),
]


def _normalizar_formato_cotas(valor: str | None) -> str:
    formato = str(valor or "ambos").strip().lower()
    if formato not in FORMATOS_COTAS_VALIDOS:
        raise ValueError(
            f"Formato de cotas inválido: {valor!r}. Use totais, percentuais ou ambos."
        )
    return formato


def _coluna(
    chave: str,
    grupo: str,
    rotulo: str,
    tipo: str,
    *,
    fonte: str | None = None,
    largura: int = 12,
) -> dict[str, Any]:
    return {
        "chave": chave,
        "grupo": grupo,
        "rotulo": rotulo,
        "tipo": tipo,
        "fonte": fonte,
        "largura": largura,
    }


def _montar_layout(formato_cotas: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    formato_cotas = _normalizar_formato_cotas(formato_cotas)
    visiveis: list[dict[str, Any]] = []
    helpers: list[dict[str, Any]] = []

    for chave, grupo, rotulo, largura in HIERARQUIA_EXPORTACAO:
        visiveis.append(_coluna(chave, grupo, rotulo, "texto", largura=largura))

    visiveis.extend(
        [
            _coluna("populacao_total", "População real", "Total", "inteiro", largura=14),
            _coluna("populacao_base", "População real", "Base usada", "inteiro", largura=14),
            _coluna("percentual_populacao", "População real", "% do recorte", "percentual", largura=13),
            _coluna("amostra", "Amostra", "Total", "inteiro", largura=10),
            _coluna("percentual_amostra", "Amostra", "% da amostra", "percentual", largura=13),
        ]
    )

    for grupo, rotulo, fonte in CATEGORIAS_COTAS:
        if formato_cotas in {"totais", "ambos"}:
            label = rotulo if formato_cotas == "totais" else f"{rotulo} (n)"
            visiveis.append(
                _coluna(f"{fonte}__n", grupo, label, "cota_inteiro", fonte=fonte, largura=12)
            )
        if formato_cotas in {"percentuais", "ambos"}:
            label = rotulo if formato_cotas == "percentuais" else f"{rotulo} (%)"
            visiveis.append(
                _coluna(f"{fonte}__pct", grupo, label, "cota_percentual", fonte=fonte, largura=13)
            )
            if formato_cotas == "percentuais":
                helpers.append(
                    _coluna(f"_helper_{fonte}", "_helpers", fonte, "helper", fonte=fonte, largura=2)
                )
    return visiveis, helpers


def _indice_por_chave(colunas: list[dict[str, Any]]) -> dict[str, int]:
    return {spec["chave"]: indice for indice, spec in enumerate(colunas, start=1)}


def _escrever_cabecalho(ws, colunas: list[dict[str, Any]]) -> None:
    grupos: list[tuple[str, list[dict[str, Any]]]] = []
    for spec in colunas:
        if grupos and grupos[-1][0] == spec["grupo"]:
            grupos[-1][1].append(spec)
        else:
            grupos.append((spec["grupo"], [spec]))

    coluna_atual = 1
    for grupo, specs in grupos:
        largura = len(specs)
        ws.cell(row=1, column=coluna_atual, value=grupo)
        if largura == 1:
            ws.merge_cells(
                start_row=1,
                start_column=coluna_atual,
                end_row=2,
                end_column=coluna_atual,
            )
        else:
            ws.merge_cells(
                start_row=1,
                start_column=coluna_atual,
                end_row=1,
                end_column=coluna_atual + largura - 1,
            )
            for deslocamento, spec in enumerate(specs):
                ws.cell(row=2, column=coluna_atual + deslocamento, value=spec["rotulo"])
        coluna_atual += largura

    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=len(colunas)):
        for cell in row:
            cell.font = FONTE_HEADER
            cell.fill = FILL_HEADER
            cell.alignment = ALINHA_CENTRO
            cell.border = BORDA
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 32


def _valor_inteiro_ou_vazio(valor: Any) -> int | None:
    if pd.isna(valor):
        return None
    return int(round(float(valor)))


def _escrever_dados(
    ws,
    df: pd.DataFrame,
    colunas: list[dict[str, Any]],
    helpers: list[dict[str, Any]],
    *,
    nome_campanha: str,
) -> tuple[int, int, dict[str, int], dict[str, int]]:
    inicio = 3
    fim = inicio + len(df) - 1
    mapa = _indice_por_chave(colunas)
    mapa_helper = {
        spec["fonte"]: len(colunas) + indice
        for indice, spec in enumerate(helpers, start=1)
    }

    for offset, (_, row) in enumerate(df.iterrows()):
        linha = inicio + offset
        for indice, spec in enumerate(colunas, start=1):
            chave = spec["chave"]
            tipo = spec["tipo"]
            if chave == "campanha":
                valor = nome_campanha
            elif chave in {"regiao_intermediaria", "regiao_imediata", "municipio", "distrito"}:
                valor = row.get(chave, "")
            elif chave in {"populacao_total", "populacao_base", "amostra"}:
                valor = _valor_inteiro_ou_vazio(row.get(chave))
            elif tipo == "cota_inteiro":
                valor = _valor_inteiro_ou_vazio(row.get(spec["fonte"]))
            else:
                valor = None

            cell = ws.cell(row=linha, column=indice, value=valor)
            cell.font = FONTE_CORPO
            cell.alignment = ALINHA_ESQUERDA if tipo == "texto" else ALINHA_CENTRO
            cell.border = BORDA
            if offset % 2 == 1:
                cell.fill = FILL_ZEBRA

        for spec in helpers:
            helper_col = mapa_helper[spec["fonte"]]
            helper_cell = ws.cell(
                row=linha,
                column=helper_col,
                value=_valor_inteiro_ou_vazio(row.get(spec["fonte"])),
            )
            helper_cell.number_format = "0"

    if len(df):
        pop_base_col = get_column_letter(mapa["populacao_base"])
        amostra_col = get_column_letter(mapa["amostra"])
        for linha in range(inicio, fim + 1):
            pop_pct_cell = ws.cell(row=linha, column=mapa["percentual_populacao"])
            pop_pct_cell.value = (
                f"=IFERROR({pop_base_col}{linha}/SUM(${pop_base_col}${inicio}:${pop_base_col}${fim}),0)"
            )
            pop_pct_cell.number_format = "0.00%"

            amostra_pct_cell = ws.cell(row=linha, column=mapa["percentual_amostra"])
            amostra_pct_cell.value = (
                f"=IFERROR({amostra_col}{linha}/SUM(${amostra_col}${inicio}:${amostra_col}${fim}),0)"
            )
            amostra_pct_cell.number_format = "0.00%"

            for spec in colunas:
                if spec["tipo"] != "cota_percentual":
                    continue
                pct_idx = mapa[spec["chave"]]
                count_key = f"{spec['fonte']}__n"
                if count_key in mapa:
                    count_idx = mapa[count_key]
                else:
                    count_idx = mapa_helper[spec["fonte"]]
                count_ref = f"{get_column_letter(count_idx)}{linha}"
                sample_ref = f"{amostra_col}{linha}"
                cell = ws.cell(row=linha, column=pct_idx)
                cell.value = f'=IF({count_ref}="","",IFERROR({count_ref}/{sample_ref},0))'
                cell.number_format = "0.00%"

    for indice, spec in enumerate(colunas, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = spec["largura"]
    for indice in range(len(colunas) + 1, len(colunas) + len(helpers) + 1):
        letra = get_column_letter(indice)
        ws.column_dimensions[letra].hidden = True
        ws.column_dimensions[letra].width = 2

    return inicio, fim, mapa, mapa_helper


def _mesclar_hierarquia(ws, quantidade: int, mapa: dict[str, int]) -> None:
    if quantidade <= 1:
        return
    inicio_dados = 3
    for chave in ["campanha", "regiao_intermediaria", "regiao_imediata", "municipio"]:
        coluna = mapa.get(chave)
        if coluna is None:
            continue
        inicio = inicio_dados
        atual = ws.cell(row=inicio, column=coluna).value
        for deslocamento in range(1, quantidade + 1):
            linha = inicio_dados + deslocamento
            valor = ws.cell(row=linha, column=coluna).value if deslocamento < quantidade else None
            if valor != atual or deslocamento == quantidade:
                fim = linha - 1
                if fim > inicio:
                    ws.merge_cells(
                        start_row=inicio,
                        start_column=coluna,
                        end_row=fim,
                        end_column=coluna,
                    )
                    ws.cell(row=inicio, column=coluna).alignment = ALINHA_CENTRO
                inicio = linha
                atual = valor


def _estilizar_campanha(ws, quantidade: int, mapa: dict[str, int]) -> None:
    coluna = mapa.get("campanha")
    if coluna is None:
        return
    for linha in range(3, 3 + quantidade):
        cell = ws.cell(row=linha, column=coluna)
        cell.fill = FILL_CAMPANHA
        cell.font = FONTE_CAMPANHA


def _escrever_totais(
    ws,
    colunas: list[dict[str, Any]],
    helpers: list[dict[str, Any]],
    inicio: int,
    fim: int,
    mapa: dict[str, int],
    mapa_helper: dict[str, int],
) -> None:
    linha_total = max(inicio, fim + 1)
    linha_pct = linha_total + 1
    ultima_hierarquia = mapa["distrito"]

    ws.cell(row=linha_total, column=1, value="TOTAL DA AMOSTRA")
    ws.merge_cells(
        start_row=linha_total,
        start_column=1,
        end_row=linha_total,
        end_column=ultima_hierarquia,
    )
    ws.cell(row=linha_pct, column=1, value="% DA AMOSTRA")
    ws.merge_cells(
        start_row=linha_pct,
        start_column=1,
        end_row=linha_pct,
        end_column=ultima_hierarquia,
    )

    amostra_total_ref = f"{get_column_letter(mapa['amostra'])}{linha_total}"
    possui_dados = fim >= inicio

    for spec in colunas:
        coluna = mapa[spec["chave"]]
        letra = get_column_letter(coluna)
        tipo = spec["tipo"]
        total_cell = ws.cell(row=linha_total, column=coluna)
        pct_cell = ws.cell(row=linha_pct, column=coluna)

        if tipo == "inteiro" and possui_dados:
            total_cell.value = f"=SUM({letra}{inicio}:{letra}{fim})"
            total_cell.number_format = "#,##0"
        elif tipo == "cota_inteiro" and possui_dados:
            total_cell.value = (
                f'=IF(COUNTBLANK({letra}{inicio}:{letra}{fim})>0,"",'
                f'SUM({letra}{inicio}:{letra}{fim}))'
            )
            total_cell.number_format = "#,##0"
        elif tipo == "percentual" and spec["chave"] in {
            "percentual_populacao",
            "percentual_amostra",
        }:
            total_cell.value = 1
            total_cell.number_format = "0.00%"

        if tipo == "cota_inteiro" and possui_dados:
            pct_cell.value = (
                f'=IF({letra}{linha_total}="","",'
                f'IFERROR({letra}{linha_total}/{amostra_total_ref},0))'
            )
            pct_cell.number_format = "0.00%"
        elif tipo == "cota_percentual" and possui_dados:
            fonte = spec["fonte"]
            count_key = f"{fonte}__n"
            if count_key in mapa:
                count_total_ref = f"{get_column_letter(mapa[count_key])}{linha_total}"
            else:
                helper_col = mapa_helper[fonte]
                helper_total_ref = f"{get_column_letter(helper_col)}{linha_total}"
                count_total_ref = helper_total_ref
            pct_cell.value = (
                f'=IF({count_total_ref}="","",'
                f'IFERROR({count_total_ref}/{amostra_total_ref},0))'
            )
            pct_cell.number_format = "0.00%"
        elif spec["chave"] == "percentual_amostra":
            pct_cell.value = 1
            pct_cell.number_format = "0.00%"

    for fonte, helper_col in mapa_helper.items():
        letra = get_column_letter(helper_col)
        if possui_dados:
            ws.cell(
                row=linha_total,
                column=helper_col,
                value=(
                    f'=IF(COUNTBLANK({letra}{inicio}:{letra}{fim})>0,"",'
                    f'SUM({letra}{inicio}:{letra}{fim}))'
                ),
            )

    for coluna in range(1, len(colunas) + 1):
        cell = ws.cell(row=linha_total, column=coluna)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = FILL_HEADER
        cell.alignment = ALINHA_CENTRO
        cell.border = BORDA

        cell_pct = ws.cell(row=linha_pct, column=coluna)
        cell_pct.font = Font(name="Calibri", size=10, bold=True, italic=True)
        cell_pct.fill = FILL_ZEBRA
        cell_pct.alignment = ALINHA_CENTRO
        cell_pct.border = BORDA


def _criar_aba_amostra(
    wb: Workbook,
    *,
    titulo: str,
    df: pd.DataFrame,
    nome_campanha: str,
    formato_cotas: str,
) -> None:
    ws = wb.create_sheet(title=titulo[:31])
    colunas, helpers = _montar_layout(formato_cotas)
    _escrever_cabecalho(ws, colunas)
    inicio, fim, mapa, mapa_helper = _escrever_dados(
        ws,
        df,
        colunas,
        helpers,
        nome_campanha=nome_campanha,
    )
    _mesclar_hierarquia(ws, len(df), mapa)
    _estilizar_campanha(ws, len(df), mapa)
    _escrever_totais(ws, colunas, helpers, inicio, fim, mapa, mapa_helper)
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# Informações e memória de cálculo
# ---------------------------------------------------------------------------
def _texto_metodo(metodo: str) -> str:
    if metodo == "manual":
        return (
            "Metas manuais por unidade: o sistema calculou e preservou uma sugestão "
            "proporcional como referência. O usuário informou apenas a quantidade bruta "
            "de entrevistas de cada unidade; depois, gênero, idade e renda foram "
            "recalculados automaticamente dentro de cada meta. A soma das unidades foi "
            "validada para fechar exatamente na amostra total."
        )
    return (
        "Distribuição proporcional pelo método do Maior Resto (Hamilton): calcula-se a "
        "amostra ideal decimal a partir da população-base, utiliza-se inicialmente a parte "
        "inteira e as entrevistas restantes são entregues às maiores frações decimais."
    )


def _codigo_unidade(row: pd.Series) -> str:
    for coluna in [
        "cod_distrito",
        "cod_municipio",
        "cod_regiao_imediata",
        "cod_regiao_intermediaria",
        "uf",
    ]:
        valor = row.get(coluna)
        if valor is not None and not pd.isna(valor) and str(valor).strip():
            return str(valor)
    return ""


def _escrever_secao(ws, linha: int, titulo: str, ultima_coluna: int = 19) -> None:
    ws.cell(row=linha, column=1, value=titulo)
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=ultima_coluna)
    cell = ws.cell(row=linha, column=1)
    cell.font = FONTE_SUBTITULO
    cell.fill = FILL_SECAO
    cell.alignment = ALINHA_ESQUERDA
    cell.border = BORDA_INFERIOR


def _criar_aba_informacoes(
    wb: Workbook,
    df: pd.DataFrame,
    *,
    dict_campanhas: dict[str, pd.DataFrame] | None,
    informacoes: dict[str, Any] | None,
    formato_cotas: str,
) -> None:
    ws = wb.active
    ws.title = "Informações"
    info = dict(informacoes or {})
    metodo = str(
        info.get("metodo_distribuicao")
        or (
            df["metodo_distribuicao"].iloc[0]
            if "metodo_distribuicao" in df.columns and len(df)
            else "proporcional"
        )
    )
    total_informado = int(
        info.get("amostra_total")
        or pd.to_numeric(df["amostra"], errors="coerce").fillna(0).sum()
    )

    # No modo por camadas, a amostra ideal de uma unidade final depende da
    # quantidade proporcional recebida pelo nível pai. Esses índices apontam
    # para as colunas visíveis usadas nos SUMIFS da memória de cálculo.
    camadas = list(info.get("camadas") or [])
    nivel_pai = camadas[-2] if len(camadas) >= 2 else None
    criterios_por_nivel = {
        "uf": [1],
        "regiao_intermediaria": [1, 2],
        "regiao_imediata": [1, 2, 3],
        "municipio": [1, 2, 3, 4],
    }
    colunas_criterio_pai = criterios_por_nivel.get(nivel_pai, [])

    ws.merge_cells("A1:S1")
    ws["A1"] = "INFORMAÇÕES E MEMÓRIA DE CÁLCULO DA AMOSTRA"
    ws["A1"].font = FONTE_TITULO
    ws["A1"].fill = FILL_HEADER
    ws["A1"].alignment = ALINHA_CENTRO
    ws.row_dimensions[1].height = 30

    _escrever_secao(ws, 3, "Configuração do projeto")
    metadados = [
        ("Nome do projeto", info.get("nome_projeto", "Projeto de amostra")),
        ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")),
        (
            "Fonte dos dados",
            info.get("fonte_dados_label", info.get("fonte_dados", "Não informado")),
        ),
        ("Modo de montagem", info.get("modo", "Não informado")),
        (
            "Método de distribuição",
            "Metas manuais por unidade" if metodo == "manual" else "Proporcional",
        ),
        (
            "Base populacional",
            "16 anos ou mais"
            if info.get("base_populacional") == "16_mais"
            else "População total",
        ),
        (
            "Nível final",
            info.get("nivel_quebra_label", info.get("nivel_quebra", "Não informado")),
        ),
        ("Recorte territorial", info.get("recorte_descricao", "Não informado")),
        ("Amostra total informada", total_informado),
        (
            "Exibição das cotas",
            ROTULOS_FORMATO_COTAS[_normalizar_formato_cotas(formato_cotas)],
        ),
    ]
    if metodo == "manual":
        nivel_meta = info.get(
            "nivel_alocacao_manual_label",
            info.get("nivel_quebra_label", info.get("nivel_quebra", "Não informado")),
        )
        metadados.insert(7, ("Unidade das metas manuais", nivel_meta))
        nivel_solicitado = info.get("nivel_saida_solicitado_label")
        if nivel_solicitado and nivel_solicitado != nivel_meta:
            metadados.insert(7, ("Nível inicialmente configurado", nivel_solicitado))
    linha = 4
    celula_total_informado = None
    for rotulo, valor in metadados:
        ws.cell(row=linha, column=1, value=rotulo).font = Font(
            name="Calibri", size=10, bold=True
        )
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=2)
        ws.cell(row=linha, column=1).alignment = ALINHA_ESQUERDA
        ws.cell(row=linha, column=3, value=valor)
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=8)
        ws.cell(row=linha, column=3).alignment = ALINHA_ESQUERDA
        if rotulo == "Amostra total informada":
            celula_total_informado = f"$C${linha}"
            ws.cell(row=linha, column=3).number_format = "#,##0"
        linha += 1

    _escrever_secao(ws, linha + 1, "Como a amostra foi calculada")
    linha += 2
    ws.cell(row=linha, column=1, value=_texto_metodo(metodo))
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha + 1, end_column=19)
    ws.cell(row=linha, column=1).alignment = ALINHA_ESQUERDA
    ws.cell(row=linha, column=1).font = Font(
        name="Calibri", size=10, color=COR_TEXTO_CINZA
    )
    linha += 3

    if dict_campanhas:
        _escrever_secao(ws, linha, "Resumo das campanhas")
        linha += 1
        cabecalhos = [
            (1, 3, "Campanha"),
            (4, 4, "Unidades"),
            (5, 5, "Amostra"),
            (6, 6, "% da amostra completa"),
        ]
        for inicio_col, fim_col, rotulo in cabecalhos:
            ws.cell(row=linha, column=inicio_col, value=rotulo)
            if fim_col > inicio_col:
                ws.merge_cells(
                    start_row=linha,
                    start_column=inicio_col,
                    end_row=linha,
                    end_column=fim_col,
                )
            cell = ws.cell(row=linha, column=inicio_col)
            cell.font = FONTE_HEADER
            cell.fill = FILL_TEAL
            cell.alignment = ALINHA_CENTRO
            cell.border = BORDA
        ws.row_dimensions[linha].height = 30
        linha_inicio_campanhas = linha + 1
        for nome, df_campanha in dict_campanhas.items():
            linha += 1
            ws.cell(row=linha, column=1, value=nome)
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
            ws.cell(row=linha, column=4, value=len(df_campanha))
            ws.cell(row=linha, column=5, value=int(df_campanha["amostra"].sum()))
            ws.cell(
                row=linha,
                column=6,
                value=f"=IFERROR(E{linha}/{celula_total_informado},0)",
            ).number_format = "0.00%"
            for col in range(1, 7):
                ws.cell(row=linha, column=col).border = BORDA
                ws.cell(row=linha, column=col).alignment = (
                    ALINHA_CENTRO if col > 3 else ALINHA_ESQUERDA
                )
        linha_total_camp = linha + 1
        ws.cell(row=linha_total_camp, column=1, value="TOTAL ATRIBUÍDO")
        ws.merge_cells(
            start_row=linha_total_camp,
            start_column=1,
            end_row=linha_total_camp,
            end_column=3,
        )
        ws.cell(
            row=linha_total_camp,
            column=4,
            value=f"=SUM(D{linha_inicio_campanhas}:D{linha})",
        )
        ws.cell(
            row=linha_total_camp,
            column=5,
            value=f"=SUM(E{linha_inicio_campanhas}:E{linha})",
        )
        ws.cell(
            row=linha_total_camp,
            column=6,
            value=f"=IFERROR(E{linha_total_camp}/{celula_total_informado},0)",
        )
        ws.cell(row=linha_total_camp, column=6).number_format = "0.00%"
        for col in range(1, 7):
            ws.cell(row=linha_total_camp, column=col).font = Font(
                name="Calibri", size=10, bold=True
            )
            ws.cell(row=linha_total_camp, column=col).fill = FILL_ZEBRA
            ws.cell(row=linha_total_camp, column=col).border = BORDA
        linha = linha_total_camp + 2

    _escrever_secao(ws, linha, "Memória de cálculo por unidade territorial")
    linha_header = linha + 1
    headers = [
        "UF",
        "Região Intermediária",
        "Região Imediata",
        "Município",
        "Distrito",
        "Código da unidade",
        "População total",
        "População-base",
        "% real do recorte",
        "% real no pai",
        "Amostra proporcional do pai",
        "Amostra ideal",
        "Sugestão proporcional",
        "Amostra aplicada",
        "% da amostra",
        "Diferença vs proporcional",
        "Mínimo",
        "Máximo",
        "Método",
    ]
    for coluna, header in enumerate(headers, start=1):
        cell = ws.cell(row=linha_header, column=coluna, value=header)
        cell.font = FONTE_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALINHA_CENTRO
        cell.border = BORDA
    ws.row_dimensions[linha_header].height = 40

    inicio_dados = linha_header + 1
    fim_dados = inicio_dados + len(df) - 1
    for offset, (_, row) in enumerate(df.iterrows()):
        linha_atual = inicio_dados + offset
        valores = [
            row.get("uf", ""),
            row.get("regiao_intermediaria", ""),
            row.get("regiao_imediata", ""),
            row.get("municipio", ""),
            row.get("distrito", ""),
            _codigo_unidade(row),
            _valor_inteiro_ou_vazio(row.get("populacao_total")),
            _valor_inteiro_ou_vazio(row.get("populacao_base")),
            None,  # % real do recorte
            None,  # % real no pai
            None,  # amostra proporcional do pai
            None,  # amostra ideal
            _valor_inteiro_ou_vazio(
                row.get("amostra_proporcional", row.get("amostra"))
            ),
            _valor_inteiro_ou_vazio(row.get("amostra")),
            None,  # % da amostra
            None,  # diferença
            _valor_inteiro_ou_vazio(row.get("minimo")),
            _valor_inteiro_ou_vazio(row.get("maximo")),
            "Manual" if metodo == "manual" else "Proporcional",
        ]
        for coluna, valor in enumerate(valores, start=1):
            cell = ws.cell(row=linha_atual, column=coluna, value=valor)
            cell.font = FONTE_CORPO
            cell.alignment = (
                ALINHA_ESQUERDA
                if coluna in {1, 2, 3, 4, 5, 6, 19}
                else ALINHA_CENTRO
            )
            cell.border = BORDA
            if offset % 2 == 1:
                cell.fill = FILL_ZEBRA

        ws.cell(
            row=linha_atual,
            column=9,
            value=(
                f"=IFERROR(H{linha_atual}/"
                f"SUM($H${inicio_dados}:$H${fim_dados}),0)"
            ),
        ).number_format = "0.00%"

        if colunas_criterio_pai:
            argumentos = []
            for col_idx in colunas_criterio_pai:
                letra = get_column_letter(col_idx)
                argumentos.extend(
                    [
                        f"${letra}${inicio_dados}:${letra}${fim_dados}",
                        f"{letra}{linha_atual}",
                    ]
                )
            criterios = ",".join(argumentos)
            ws.cell(
                row=linha_atual,
                column=10,
                value=(
                    f"=IFERROR(H{linha_atual}/SUMIFS("
                    f"$H${inicio_dados}:$H${fim_dados},{criterios}),0)"
                ),
            ).number_format = "0.00%"
            ws.cell(
                row=linha_atual,
                column=11,
                value=(
                    f"=SUMIFS($M${inicio_dados}:$M${fim_dados},{criterios})"
                ),
            ).number_format = "#,##0"
        else:
            ws.cell(row=linha_atual, column=10, value=f"=I{linha_atual}").number_format = (
                "0.00%"
            )
            ws.cell(
                row=linha_atual,
                column=11,
                value=f"={celula_total_informado}",
            ).number_format = "#,##0"

        ws.cell(
            row=linha_atual,
            column=12,
            value=f"=J{linha_atual}*K{linha_atual}",
        ).number_format = "0.000"
        ws.cell(
            row=linha_atual,
            column=15,
            value=f"=IFERROR(N{linha_atual}/{celula_total_informado},0)",
        ).number_format = "0.00%"
        ws.cell(row=linha_atual, column=16, value=f"=N{linha_atual}-M{linha_atual}")

    linha_total = fim_dados + 1 if len(df) else inicio_dados
    ws.cell(row=linha_total, column=1, value="TOTAL")
    ws.merge_cells(start_row=linha_total, start_column=1, end_row=linha_total, end_column=6)
    for col in [7, 8, 12, 13, 14, 16, 17, 18]:
        letra = get_column_letter(col)
        if len(df):
            ws.cell(
                row=linha_total,
                column=col,
                value=f"=SUM({letra}{inicio_dados}:{letra}{fim_dados})",
            )
    ws.cell(row=linha_total, column=9, value=1).number_format = "0.00%"
    ws.cell(row=linha_total, column=15, value=1).number_format = "0.00%"
    ws.cell(row=linha_total, column=19, value="FECHAMENTO")
    for col in range(1, 20):
        cell = ws.cell(row=linha_total, column=col)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = FILL_HEADER
        cell.alignment = ALINHA_CENTRO
        cell.border = BORDA

    linha_validacao = linha_total + 3
    _escrever_secao(ws, linha_validacao, "Validações automáticas")
    verificacoes = [
        (
            "Fechamento da amostra",
            f'=IF(N{linha_total}={celula_total_informado},"OK","DIVERGÊNCIA")',
        ),
        (
            "Fechamento percentual",
            f'=IF(ABS(O{linha_total}-1)<0.000001,"OK","DIVERGÊNCIA")',
        ),
        (
            "Diferença total vs proporcional",
            f'=IF(P{linha_total}=0,"OK","REVISAR")',
        ),
        (
            "Unidades com amostra zero",
            f"=COUNTIF(N{inicio_dados}:N{fim_dados},0)" if len(df) else 0,
        ),
    ]
    for indice, (rotulo, formula) in enumerate(verificacoes, start=1):
        linha_v = linha_validacao + indice
        ws.cell(row=linha_v, column=1, value=rotulo).font = Font(
            name="Calibri", size=10, bold=True
        )
        ws.merge_cells(start_row=linha_v, start_column=1, end_row=linha_v, end_column=2)
        ws.cell(row=linha_v, column=1).alignment = ALINHA_ESQUERDA
        ws.cell(row=linha_v, column=3, value=formula)
        ws.merge_cells(start_row=linha_v, start_column=3, end_row=linha_v, end_column=6)
        ws.cell(row=linha_v, column=3).alignment = ALINHA_CENTRO
        ws.cell(row=linha_v, column=3).border = BORDA
        ws.row_dimensions[linha_v].height = 22

    larguras = {
        "A": 10,
        "B": 24,
        "C": 24,
        "D": 24,
        "E": 24,
        "F": 18,
        "G": 15,
        "H": 15,
        "I": 14,
        "J": 14,
        "K": 17,
        "L": 14,
        "M": 17,
        "N": 15,
        "O": 14,
        "P": 18,
        "Q": 10,
        "R": 10,
        "S": 15,
    }
    for letra, largura in larguras.items():
        ws.column_dimensions[letra].width = largura
    ws.freeze_panes = f"A{inicio_dados}"
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# Controles por camada
# ---------------------------------------------------------------------------
_ROTULOS_RESUMO = {
    "uf": "UF",
    "regiao_intermediaria": "Região Intermediária",
    "regiao_imediata": "Região Imediata",
    "municipio": "Município",
    "distrito": "Distrito",
    "populacao_total": "População total",
    "populacao_base": "População-base",
    "percentual": "% real do total",
    "percentual_no_pai": "% real no pai",
    "amostra_ideal": "Amostra ideal",
    "amostra_proporcional": "Sugestão proporcional",
    "amostra": "Amostra aplicada",
    "percentual_amostra": "% da amostra",
    "diferenca_vs_proporcional": "Diferença",
}


def _escrever_resumo_camada(wb: Workbook, nivel: str, df: pd.DataFrame) -> None:
    nomes_aba = {
        "uf": "Controle - UF",
        "regiao_intermediaria": "Controle - Reg Interm",
        "regiao_imediata": "Controle - Reg Imed",
        "municipio": "Controle - Município",
        "distrito": "Controle - Distrito",
    }
    ws = wb.create_sheet(title=nomes_aba.get(nivel, f"Controle - {nivel}")[:31])
    colunas = [coluna for coluna in _ROTULOS_RESUMO if coluna in df.columns]

    for indice, coluna in enumerate(colunas, start=1):
        cell = ws.cell(row=1, column=indice, value=_ROTULOS_RESUMO[coluna])
        cell.font = FONTE_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALINHA_CENTRO
        cell.border = BORDA

    for linha_excel, (_, row) in enumerate(df.iterrows(), start=2):
        for indice, coluna in enumerate(colunas, start=1):
            valor = row[coluna]
            cell = ws.cell(row=linha_excel, column=indice)
            if coluna in {"percentual", "percentual_no_pai", "percentual_amostra"}:
                cell.value = None if pd.isna(valor) else float(valor) / 100
                cell.number_format = "0.00%"
            elif coluna == "amostra_ideal":
                cell.value = None if pd.isna(valor) else float(valor)
                cell.number_format = "0.000"
            elif coluna in {
                "populacao_total",
                "populacao_base",
                "amostra_proporcional",
                "amostra",
                "diferenca_vs_proporcional",
            }:
                cell.value = _valor_inteiro_ou_vazio(valor)
                cell.number_format = "#,##0"
            else:
                cell.value = valor
            cell.font = FONTE_CORPO
            cell.alignment = ALINHA_CENTRO if coluna in {
                "populacao_total",
                "populacao_base",
                "percentual",
                "percentual_no_pai",
                "amostra_ideal",
                "amostra_proporcional",
                "amostra",
                "percentual_amostra",
                "diferenca_vs_proporcional",
            } else ALINHA_ESQUERDA
            cell.border = BORDA
            if linha_excel % 2 == 1:
                cell.fill = FILL_ZEBRA

    linha_total = len(df) + 2
    ws.cell(row=linha_total, column=1, value="TOTAL")
    ws.cell(row=linha_total, column=1).font = Font(name="Calibri", size=10, bold=True)
    for coluna in [
        "populacao_total",
        "populacao_base",
        "amostra_proporcional",
        "amostra",
        "diferenca_vs_proporcional",
    ]:
        if coluna not in colunas or not len(df):
            continue
        indice = colunas.index(coluna) + 1
        letra = get_column_letter(indice)
        ws.cell(row=linha_total, column=indice, value=f"=SUM({letra}2:{letra}{linha_total - 1})")
        ws.cell(row=linha_total, column=indice).font = Font(name="Calibri", size=10, bold=True)
    if "percentual" in colunas:
        indice = colunas.index("percentual") + 1
        ws.cell(row=linha_total, column=indice, value=1).number_format = "0.00%"
    if "percentual_amostra" in colunas:
        indice = colunas.index("percentual_amostra") + 1
        ws.cell(row=linha_total, column=indice, value=1).number_format = "0.00%"

    for indice, coluna in enumerate(colunas, start=1):
        largura = 14
        if coluna in {"regiao_intermediaria", "regiao_imediata", "municipio", "distrito"}:
            largura = 25
        ws.column_dimensions[get_column_letter(indice)].width = largura
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _nome_aba_unico(wb: Workbook, nome: str) -> str:
    base = (nome or "Campanha")[:31]
    candidato = base
    contador = 2
    while candidato in wb.sheetnames:
        sufixo = f" ({contador})"
        candidato = f"{base[:31 - len(sufixo)]}{sufixo}"
        contador += 1
    return candidato


# ---------------------------------------------------------------------------
# Função pública
# ---------------------------------------------------------------------------
def exportar_amostra(
    df_completo: pd.DataFrame,
    dict_campanhas: dict[str, pd.DataFrame] | None = None,
    caminho_saida: str | Path = "amostra.xlsx",
    resumos_camadas: dict[str, pd.DataFrame] | None = None,
    *,
    informacoes: dict[str, Any] | None = None,
    exibicao_cotas: str = "ambos",
):
    """Gera o arquivo Excel final e retorna o caminho gravado."""
    formato = _normalizar_formato_cotas(exibicao_cotas)
    if df_completo is None or df_completo.empty:
        raise ValueError("df_completo precisa possuir ao menos uma unidade territorial.")

    wb = Workbook()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass

    _criar_aba_informacoes(
        wb,
        df_completo,
        dict_campanhas=dict_campanhas,
        informacoes=informacoes,
        formato_cotas=formato,
    )
    _criar_aba_amostra(
        wb,
        titulo="Amostra Completa",
        df=df_completo,
        nome_campanha="AMOSTRA COMPLETA",
        formato_cotas=formato,
    )

    if resumos_camadas:
        for nivel, df_resumo in resumos_camadas.items():
            _escrever_resumo_camada(wb, nivel, df_resumo)

    if dict_campanhas:
        for nome, df_campanha in dict_campanhas.items():
            _criar_aba_amostra(
                wb,
                titulo=_nome_aba_unico(wb, nome),
                df=df_campanha,
                nome_campanha=nome,
                formato_cotas=formato,
            )

    caminho = Path(caminho_saida)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho)
    return str(caminho)


if __name__ == "__main__":
    from core.amostra_math import calcular_amostra

    resultado = calcular_amostra(
        uf="SP",
        amostra_total=1000,
        nivel_quebra="regiao_imediata",
        base_populacional="16_mais",
    )
    caminho = exportar_amostra(
        resultado,
        caminho_saida="teste_amostra_SP_1000.xlsx",
        informacoes={
            "nome_projeto": "Teste SP",
            "amostra_total": 1000,
            "metodo_distribuicao": "proporcional",
            "base_populacional": "16_mais",
            "nivel_quebra_label": "Região Imediata",
        },
        exibicao_cotas="ambos",
    )
    print(f"OK: {caminho}")
