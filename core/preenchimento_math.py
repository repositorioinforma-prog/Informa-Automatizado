"""
Preenchimento do esqueleto da apresentação (lógica pura, sem Streamlit).

Fluxo: esqueleto .pptx (gerado pelo app) + arquivos do estudo -> .pptx preenchido + planilha de conferência.

  1. `reunir_arquivos`   abre zips (inclusive zips dentro de zips) e devolve [(nome, bytes)];
  2. `classificar`       identifica cada arquivo pelo CONTEÚDO (esqueleto, perfil, descritivo...);
  3. `preencher`         aplica cada fonte encontrada no esqueleto; o que faltar fica com o aviso
                         «Arquivo de X não enviado – conteúdo pendente»;
  4. `Conferencia`       acumula o que foi escrito (slide, item, fonte, valor), pendências,
                         divergências e avisos, e gera o .xlsx de conferência.

Fontes já suportadas: Perfil. (Intenção de voto e Análise das perguntas entram na sequência.)
Regras do projeto: nunca inventar valor, nunca usar zero como preenchimento.
"""
from __future__ import annotations

import io
import posixpath
import zipfile
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pptx import Presentation

from core.perfil_math import eh_perfil, ler_perfil, preencher_slide_perfil
from core.planilha_leitura import ErroLeitura, eh_zip, linhas_da_planilha
from core.verificador_relatorio_math import chave

ARQUIVOS_ESPERADOS = (  # tipo -> rótulo mostrado à pessoa
    ("perfil", "Perfil (Excel)"),
)
TIPOS_CONHECIDOS = {"esqueleto": "Esqueleto da apresentação", "perfil": "Perfil", "descritivo": "Relatório Descritivo"}


# ------------------------------------------------------------------ conferência
class Conferencia:
    def __init__(self) -> None:
        self.linhas: List[tuple] = []        # (slide, item, fonte, valor)
        self.pendencias: List[tuple] = []    # (slide, item, motivo)
        self.divergencias: List[tuple] = []  # (item, valor A, valor B, observação)
        self.avisos: List[str] = []

    def linha(self, slide, item: str, fonte: str, valor) -> None:
        self.linhas.append((slide, item, fonte, valor))

    def pendencia(self, slide, item: str, motivo: str) -> None:
        self.pendencias.append((slide, item, motivo))

    def divergencia(self, item: str, a, b, obs: str = "") -> None:
        self.divergencias.append((item, a, b, obs))

    def aviso(self, texto: str) -> None:
        if texto not in self.avisos:
            self.avisos.append(texto)

    def para_xlsx(self, estudo: Optional[dict] = None) -> bytes:
        wb = openpyxl.Workbook()
        cab = PatternFill("solid", fgColor="002060")

        def aba(ws, titulo, colunas, linhas, larguras):
            ws.title = titulo
            ws.append(colunas)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = cab
                c.alignment = Alignment(vertical="center", wrap_text=True)
            for r in linhas:
                ws.append(list(r))
            for i, w in enumerate(larguras, 1):
                ws.column_dimensions[get_column_letter(i)].width = w
            ws.freeze_panes = "A2"
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.alignment = Alignment(vertical="top", wrap_text=True)

        aba(wb.active, "Conferência", ["Slide", "Item", "Fonte", "Valor"], self.linhas, [8, 52, 34, 16])
        aba(wb.create_sheet(), "Divergências", ["Item", "Valor A", "Valor B", "Observação"], self.divergencias,
            [44, 20, 20, 60])
        aba(wb.create_sheet(), "Pendências e observações", ["Slide", "Item", "Motivo / observação"],
            list(self.pendencias) + [("", "Aviso", a) for a in self.avisos], [8, 44, 90])
        res = wb.create_sheet("Resumo", 0)
        res.append(["Conferência do preenchimento da apresentação"])
        res["A1"].font = Font(bold=True, size=14)
        est = estudo or {}
        for rot, val in (("Estudo", est.get("local", "")), ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")),
                         ("Valores preenchidos", len(self.linhas)), ("Pendências", len(self.pendencias)),
                         ("Divergências", len(self.divergencias)), ("Avisos", len(self.avisos))):
            res.append([rot, val])
        res.column_dimensions["A"].width = 24
        res.column_dimensions["B"].width = 40
        saida = io.BytesIO()
        wb.save(saida)
        return saida.getvalue()


# ------------------------------------------------------------------ arquivos
def reunir_arquivos(arquivos: List[Tuple[str, bytes]], _nivel: int = 0) -> List[Tuple[str, bytes]]:
    """Abre zips (até 3 níveis) e descarta lixo (pastas, __MACOSX, ~$temporários). Devolve [(nome, bytes)]."""
    saida: List[Tuple[str, bytes]] = []
    for nome, dados in arquivos:
        base = posixpath.basename(nome.replace("\\", "/"))
        if not base or base.startswith(("~$", ".")) or "__MACOSX" in nome:
            continue
        if base.lower().endswith(".zip") and eh_zip(dados) and _nivel < 3:
            try:
                with zipfile.ZipFile(io.BytesIO(dados)) as z:
                    internos = [(i.filename, z.read(i)) for i in z.infolist() if not i.is_dir()]
            except zipfile.BadZipFile:
                continue
            saida.extend(reunir_arquivos(internos, _nivel + 1))
        else:
            saida.append((base, dados))
    return saida


def _tem_marca_pendente(dados: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(dados)) as z:
            return any(n.startswith("ppt/slides/slide") and n.endswith(".xml") and b"INFORMA_PENDENTE" in z.read(n)
                       for n in z.namelist())
    except zipfile.BadZipFile:
        return False


def tipo_do_arquivo(nome: str, dados: bytes) -> str:
    """'esqueleto' | 'perfil' | 'descritivo' | 'desconhecido' — decidido pelo conteúdo, não pelo nome."""
    baixo = nome.lower()
    if baixo.endswith(".pptx") and eh_zip(dados):
        # esqueleto = deck gerado pelo app (tem a marca INFORMA_PENDENTE); decks prontos do estudo não entram
        return "esqueleto" if _tem_marca_pendente(dados) else "desconhecido"
    if baixo.endswith((".xlsx", ".xlsm", ".xls")):
        try:
            linhas = linhas_da_planilha(dados)
        except ErroLeitura:
            return "desconhecido"
        if eh_perfil(linhas):
            return "perfil"
        marcadores = {chave(l[0]) for l in linhas[:5000] if l and isinstance(l[0], str)}
        if {"resultados pelo total", "resultados pelos segmentos"} & marcadores:
            return "descritivo"
    return "desconhecido"


def classificar(arquivos: List[Tuple[str, bytes]]) -> dict:
    """
    {"por_tipo": {tipo: (nome, bytes)}, "duplicados": {tipo: [nomes]}, "ignorados": [nomes], "erros": [...]}.
    Se houver dois arquivos do mesmo tipo, vale o primeiro e os outros ficam em `duplicados`.
    """
    por_tipo: Dict[str, Tuple[str, bytes]] = {}
    duplicados: Dict[str, List[str]] = {}
    ignorados: List[str] = []
    for nome, dados in reunir_arquivos(arquivos):
        tipo = tipo_do_arquivo(nome, dados)
        if tipo == "desconhecido":
            ignorados.append(nome)
        elif tipo in por_tipo:
            duplicados.setdefault(tipo, []).append(nome)
        else:
            por_tipo[tipo] = (nome, dados)
    return {"por_tipo": por_tipo, "duplicados": duplicados, "ignorados": ignorados}


# ------------------------------------------------------------------ preenchimento
def preencher(pptx_bytes: bytes, fontes: Dict[str, Tuple[str, bytes]], base_esperada: Optional[int] = None,
              estudo: Optional[dict] = None) -> Tuple[bytes, Conferencia]:
    """
    Preenche o esqueleto com as fontes disponíveis (`fontes` = por_tipo de `classificar`).
    `base_esperada` (opcional, vinda do Descritivo/esqueleto) é comparada com a base do Perfil.
    """
    conf = Conferencia()
    prs = Presentation(io.BytesIO(pptx_bytes))

    perfil = None
    nome_perfil = "Perfil"
    if "perfil" in fontes:
        nome_perfil, dados = fontes["perfil"]
        try:
            perfil = ler_perfil(dados)
        except ErroLeitura as exc:
            conf.aviso(f"{nome_perfil}: {exc}")
    preencher_slide_perfil(prs, perfil, conf, nome_perfil)
    if perfil and base_esperada and perfil.get("base") and perfil["base"] != base_esperada:
        conf.divergencia("Base de entrevistas", f"Perfil: {perfil['base']}", f"Esqueleto/Descritivo: {base_esperada}",
                         "A base do Perfil deveria ser igual à do Relatório Descritivo")

    saida = io.BytesIO()
    prs.save(saida)
    return saida.getvalue(), conf
