"""Standalone engine to detect and split report tables across printed pages.

The module is intentionally isolated from the existing report workflow. It reads
workbooks produced by the app, detects semantic table blocks, estimates page
breaks when Excel automatic breaks are not stored in the XLSX, and replaces
only tables that would cross a page boundary.
"""
from __future__ import annotations

from copy import copy
from dataclasses import dataclass
import math
import re
from typing import Iterable

from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.pagebreak import Break

from core.planilha_utils import inserir_linhas_seguro


DEFAULT_ROW_HEIGHT = 15.0
GAP_ROWS = 4
BREAK_AFTER_GAP_ROWS = 2
# Excel automatic page breaks depend on printer metrics that are not persisted in XLSX.
# Keep a small vertical reserve so borderline tables are not falsely considered to fit.
PAGINATION_RESERVE_POINTS = 18.0
CONTINUATION_TEXT = " Continua\u00e7\u00e3o"


@dataclass
class TableInfo:
    sheet: str
    start_row: int
    title_row: int
    header_start: int
    label_start: int
    label_end: int
    base_row: int
    footer_end: int
    legend_start: int | None
    legend_end: int | None
    title: str

    @property
    def end_row(self) -> int:
        """Last row that belongs to the regional table footprint.

        For regional tables the legend is part of the table. Therefore any
        page break after the title and before the final legend row must trigger
        a split, even when it falls on Base, Pergunta or the spacer before
        LEGENDA.
        """
        return self.legend_end if self.legend_end is not None else self.footer_end

    @property
    def label_count(self) -> int:
        return max(0, self.label_end - self.label_start + 1)

    @property
    def has_legend(self) -> bool:
        return self.legend_start is not None and self.legend_end is not None


@dataclass
class SplitResult:
    sheet: str
    title: str
    original_start: int
    original_end: int
    labels: int
    parts: int
    part_sizes: list[int]
    reason: str


def _norm(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    repl = {
        "A": "A", "E": "E", "I": "I", "O": "O", "U": "U", "C": "C",
    }
    try:
        import unicodedata
        text = "".join(
            ch for ch in unicodedata.normalize("NFD", text)
            if unicodedata.category(ch) != "Mn"
        )
    except Exception:
        pass
    return re.sub(r"\s+", " ", text)


def _row_height(ws, row: int) -> float:
    dim = ws.row_dimensions.get(row)
    if dim and dim.height is not None:
        return float(dim.height)
    if ws.sheet_format.defaultRowHeight is not None:
        return float(ws.sheet_format.defaultRowHeight)
    return DEFAULT_ROW_HEIGHT


def _row_has_content(ws, row: int) -> bool:
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None and str(v).strip() != "":
            return True
    return False


def _row_has_total(ws, row: int) -> bool:
    return any(_norm(ws.cell(row=row, column=c).value) == "TOTAL" for c in range(1, ws.max_column + 1))


def _row_looks_like_data(ws, row: int) -> bool:
    label = ws.cell(row=row, column=1).value
    if label is None or str(label).strip() == "":
        return False
    if _norm(label) in {"BASE", "BASE REDUZIDA"}:
        return False
    for c in range(2, ws.max_column + 1):
        value = ws.cell(row=row, column=c).value
        if isinstance(value, (int, float)):
            return True
        if isinstance(value, str) and value.startswith("="):
            return True
    return False


def _find_footer_end(ws, base_row: int) -> int:
    """Return the final fixed row between Base and the legend.

    Production reports are not uniform: some tables have Pergunta immediately
    after Base, while evaluation tables include rows such as ``Média`` between
    Base and Pergunta. These post-base summary rows are part of the table and
    must be repeated in every split.
    """
    end = base_row
    r = base_row + 1
    probe_limit = min(ws.max_row, base_row + 8)
    found_question = False

    while r <= probe_limit:
        a = ws.cell(row=r, column=1).value
        n = _norm(a)

        if n.startswith("LEGENDA"):
            break
        if _row_has_total(ws, r) or n == "BASE":
            break

        if _row_has_content(ws, r):
            end = r
            if n.startswith("PERGUNTA:") or str(a or "").lstrip().startswith("*"):
                found_question = True
                r += 1
                # Preserve wrapped/multiline textual footer rows immediately
                # after the question, but never absorb the legend/new table.
                while r <= ws.max_row and _row_has_content(ws, r):
                    nn = _norm(ws.cell(row=r, column=1).value)
                    if nn.startswith("LEGENDA") or _row_has_total(ws, r):
                        break
                    end = r
                    r += 1
                break
        elif found_question:
            break

        r += 1

    return end


def _find_legend_block(ws, footer_end: int) -> tuple[int | None, int | None]:
    """Find a LEGENDA block immediately after the table footer.

    Report legends are allowed to start after a small blank spacer and continue
    for multiple rows (including rich text). The block ends at the first blank
    row after legend content or at the next table-like content.
    """
    r = footer_end + 1
    # The report convention uses a small spacer between Pergunta and LEGENDA.
    max_probe = min(ws.max_row, footer_end + 4)
    while r <= max_probe and not _row_has_content(ws, r):
        r += 1
    legend_marker = _norm(ws.cell(row=r, column=1).value) if r <= ws.max_row else ""
    # Production reports use markers such as "LEGENDA REGIÕES" and
    # "LEGENDA CAPITAL", not only the literal word "LEGENDA". Treat any
    # marker that starts with LEGENDA as the beginning of the protected block.
    if r > ws.max_row or not legend_marker.startswith("LEGENDA"):
        return None, None

    legend_start = r
    legend_end = r
    r += 1
    while r <= ws.max_row:
        if not _row_has_content(ws, r):
            break
        # Do not absorb a new report table if a malformed legend has no spacer.
        if _row_has_total(ws, r):
            break
        legend_end = r
        r += 1
    return legend_start, legend_end


def _find_title_row(ws, header_start: int) -> int:
    r = header_start - 1
    while r >= 1 and not _row_has_content(ws, r):
        r -= 1
    if r < 1:
        return header_start
    # A title is normally the nearest non-empty row immediately above the header.
    return r


def detectar_tabelas(ws) -> list[TableInfo]:
    """Detect report tables using the semantic markers Total + Base + data rows."""
    tables: list[TableInfo] = []
    used_bases: set[int] = set()

    for base_row in range(1, ws.max_row + 1):
        if _norm(ws.cell(row=base_row, column=1).value) != "BASE":
            continue
        if base_row in used_bases:
            continue

        header_start = None
        for r in range(base_row - 1, max(0, base_row - 100), -1):
            # A large table may have dozens of labels between its header and Base.
            # Stop only if another Base is reached, which means we crossed into
            # the preceding table.
            if _norm(ws.cell(row=r, column=1).value) == "BASE":
                break
            if _row_has_total(ws, r):
                header_start = r
                break
        if header_start is None:
            continue

        data_rows = [r for r in range(header_start + 1, base_row) if _row_looks_like_data(ws, r)]
        if not data_rows:
            continue
        label_start = min(data_rows)
        label_end = max(data_rows)
        # Require labels to be essentially contiguous. This filters out unrelated Base rows.
        if any(not _row_looks_like_data(ws, r) for r in range(label_start, label_end + 1)):
            continue

        title_row = _find_title_row(ws, header_start)
        footer_end = _find_footer_end(ws, base_row)
        legend_start, legend_end = _find_legend_block(ws, footer_end)
        title_value = ws.cell(row=title_row, column=1).value
        title = str(title_value).strip() if title_value is not None else f"Tabela linha {title_row}"

        tables.append(
            TableInfo(
                sheet=ws.title,
                start_row=title_row,
                title_row=title_row,
                header_start=header_start,
                label_start=label_start,
                label_end=label_end,
                base_row=base_row,
                footer_end=footer_end,
                legend_start=legend_start,
                legend_end=legend_end,
                title=title,
            )
        )
        used_bases.add(base_row)

    # De-duplicate accidental overlaps, keeping the earliest valid table.
    out: list[TableInfo] = []
    last_end = 0
    for table in sorted(tables, key=lambda t: t.start_row):
        if table.start_row <= last_end:
            continue
        out.append(table)
        last_end = table.end_row
    return out


def detectar_tabelas_regionais(ws) -> list[TableInfo]:
    """Fast-path detector anchored on LEGENDA markers.

    The divider only operates on regional/capital tables that have a legend,
    so scanning all Base rows in a multi-thousand-row report is unnecessary.
    Anchoring on ``LEGENDA*`` is both faster and more faithful to the business
    rule.
    """
    tables: list[TableInfo] = []
    legend_rows = [
        r for r in range(1, ws.max_row + 1)
        if _norm(ws.cell(row=r, column=1).value).startswith("LEGENDA")
    ]

    for legend_start in legend_rows:
        # Base is normally 2-4 rows above the legend (Pergunta, optional Média,
        # spacer). Keep a wider probe for future report variants.
        base_row = None
        for r in range(legend_start - 1, max(0, legend_start - 15), -1):
            if _norm(ws.cell(row=r, column=1).value) == "BASE":
                base_row = r
                break
        if base_row is None:
            continue

        header_start = None
        for r in range(base_row - 1, max(0, base_row - 100), -1):
            if _norm(ws.cell(row=r, column=1).value) == "BASE":
                break
            if _row_has_total(ws, r):
                header_start = r
                break
        if header_start is None:
            continue

        data_rows = [r for r in range(header_start + 1, base_row) if _row_looks_like_data(ws, r)]
        if not data_rows:
            continue
        label_start = min(data_rows)
        label_end = max(data_rows)
        if any(not _row_looks_like_data(ws, r) for r in range(label_start, label_end + 1)):
            continue

        title_row = _find_title_row(ws, header_start)
        footer_end = _find_footer_end(ws, base_row)
        found_legend_start, legend_end = _find_legend_block(ws, footer_end)
        if found_legend_start != legend_start or legend_end is None:
            continue

        title_value = ws.cell(row=title_row, column=1).value
        title = str(title_value).strip() if title_value is not None else f"Tabela linha {title_row}"
        tables.append(TableInfo(
            sheet=ws.title, start_row=title_row, title_row=title_row,
            header_start=header_start, label_start=label_start, label_end=label_end,
            base_row=base_row, footer_end=footer_end, legend_start=legend_start,
            legend_end=legend_end, title=title,
        ))

    return sorted(tables, key=lambda t: t.start_row)


def _print_area_rows(ws) -> tuple[int, int]:
    if ws.print_area:
        try:
            text = str(ws.print_area)
            first = text.split(",")[0].replace("'", "")
            if "!" in first:
                first = first.split("!", 1)[1]
            min_col, min_row, max_col, max_row = range_boundaries(first.replace("$", ""))
            return max(1, min_row), min(ws.max_row, max_row)
        except Exception:
            pass
    return 1, ws.max_row


def _paper_height_points(ws) -> float:
    # Point heights for common paper sizes. The app uses A4; fall back to A4.
    paper = str(ws.page_setup.paperSize or "9")
    heights_portrait = {
        "1": 792.0,   # Letter 11 in
        "5": 1008.0,  # Legal 14 in
        "9": 841.89,  # A4 297 mm
    }
    widths_portrait = {
        "1": 612.0,
        "5": 612.0,
        "9": 595.28,
    }
    orientation = (ws.page_setup.orientation or "portrait").lower()
    if orientation == "landscape":
        return widths_portrait.get(paper, 595.28)
    return heights_portrait.get(paper, 841.89)


def _usable_page_height(ws) -> float:
    page_height = _paper_height_points(ws)
    top = float(ws.page_margins.top or 0.75) * 72.0
    bottom = float(ws.page_margins.bottom or 0.75) * 72.0
    usable = max(120.0, page_height - top - bottom)
    scale = ws.page_setup.scale
    try:
        scale_factor = float(scale or 100) / 100.0
    except Exception:
        scale_factor = 1.0
    if scale_factor <= 0:
        scale_factor = 1.0

    # Automatic/dashed page breaks are calculated by Excel using the active
    # printer driver. Those printer metrics are not stored in the XLSX, so a
    # mathematically exact A4 calculation can differ by a few points from what
    # Excel displays. Reserve a small amount of vertical space before undoing
    # the worksheet scale. This intentionally classifies borderline cases as
    # crossing the page instead of leaving Pergunta/LEGENDA on the next page.
    usable = max(120.0, usable - PAGINATION_RESERVE_POINTS)
    return usable / scale_factor


def estimar_quebras_automaticas(ws) -> list[int]:
    """Estimate horizontal automatic breaks from row heights and page setup.

    Returned IDs follow openpyxl/Excel convention: Break(id=n) means the next
    printed page starts at row n + 1.
    """
    start_row, end_row = _print_area_rows(ws)
    manual = sorted({int(b.id) for b in ws.row_breaks.brk if b.id is not None})
    manual_set = set(manual)
    usable = _usable_page_height(ws)
    breaks: list[int] = []
    current = 0.0

    for r in range(start_row, end_row + 1):
        h = _row_height(ws, r)
        if current > 0 and current + h > usable:
            br = r - 1
            if br >= start_row:
                breaks.append(br)
            current = 0.0
        current += h
        if r in manual_set:
            current = 0.0

    return sorted(set(breaks))


def _page_context(ws, row: int, effective_breaks: Iterable[int]) -> tuple[int, float, float]:
    """Return page start row, used height before row, and usable height."""
    start_print, _ = _print_area_rows(ws)
    previous = [b for b in effective_breaks if b < row]
    page_start = (max(previous) + 1) if previous else start_print
    used = sum(_row_height(ws, r) for r in range(page_start, row))
    return page_start, used, _usable_page_height(ws)


def _fixed_part_height(ws, table: TableInfo, continuation: bool) -> float:
    # Everything outside the variable label block is repeated in every part.
    # This includes Base, Pergunta, the spacer and the full LEGENDA block.
    rows = list(range(table.start_row, table.label_start)) + list(range(table.base_row, table.end_row + 1))
    return sum(_row_height(ws, r) for r in rows)


def _balanced_sizes(total: int, parts: int) -> list[int]:
    base = total // parts
    rem = total % parts
    return [base + (1 if i < rem else 0) for i in range(parts)]


def _choose_part_sizes(ws, table: TableInfo, effective_breaks: list[int]) -> list[int]:
    label_heights = [_row_height(ws, r) for r in range(table.label_start, table.label_end + 1)]
    total = len(label_heights)
    if total <= 1:
        return [total]

    crossing = _breaks_inside_table(table, effective_breaks)
    parts = max(2, len(crossing) + 1)
    _, used_before, usable = _page_context(ws, table.start_row, effective_breaks)
    first_capacity = max(1.0, usable - used_before)
    full_capacity = usable - (BREAK_AFTER_GAP_ROWS * DEFAULT_ROW_HEIGHT)
    fixed = _fixed_part_height(ws, table, continuation=False)

    while parts <= total:
        sizes = _balanced_sizes(total, parts)
        cursor = 0
        fits = True
        for i, size in enumerate(sizes):
            labels_h = sum(label_heights[cursor:cursor + size])
            cursor += size
            cap = first_capacity if i == 0 else full_capacity
            if fixed + labels_h > cap + 0.01:
                fits = False
                break
        if fits:
            return sizes
        parts += 1
    return [1] * total


def _capture_block(ws, table: TableInfo):
    cells = {}
    for r in range(table.start_row, table.end_row + 1):
        for c in range(1, ws.max_column + 1):
            src = ws.cell(row=r, column=c)
            cells[(r, c)] = {
                "value": src.value,
                "font": copy(src.font),
                "fill": copy(src.fill),
                "border": copy(src.border),
                "alignment": copy(src.alignment),
                "number_format": src.number_format,
                "protection": copy(src.protection),
            }
    heights = {r: _row_height(ws, r) for r in range(table.start_row, table.end_row + 1)}
    merges = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= table.start_row and rng.max_row <= table.end_row:
            merges.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
    return cells, heights, merges


def _clear_region(ws, start: int, end: int):
    for rng in list(ws.merged_cells.ranges):
        if not (rng.max_row < start or rng.min_row > end):
            try:
                ws.unmerge_cells(str(rng))
            except KeyError:
                ws.merged_cells.ranges.discard(rng)
    for r in range(start, end + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell._style = copy(ws.cell(row=1, column=1)._style) if False else copy(cell._style)
            # Explicitly clear visual style without changing column-level formatting.
            cell.font = copy(ws.parent._named_styles[0].font)
            cell.fill = PatternFill(fill_type=None)
            cell.border = copy(ws.parent._named_styles[0].border)
            cell.alignment = copy(ws.parent._named_styles[0].alignment)
            cell.number_format = "General"
            cell.protection = copy(ws.parent._named_styles[0].protection)
        ws.row_dimensions[r].height = None


def _copy_snapshot_row(ws, cells, heights, source_row: int, dest_row: int):
    for c in range(1, ws.max_column + 1):
        data = cells[(source_row, c)]
        dst = ws.cell(row=dest_row, column=c)
        value = data["value"]
        if isinstance(value, str) and value.startswith("="):
            try:
                value = Translator(value, origin=f"{get_column_letter(c)}{source_row}").translate_formula(
                    f"{get_column_letter(c)}{dest_row}"
                )
            except Exception:
                pass
        dst.value = value
        dst.font = copy(data["font"])
        dst.fill = copy(data["fill"])
        dst.border = copy(data["border"])
        dst.alignment = copy(data["alignment"])
        dst.number_format = data["number_format"]
        dst.protection = copy(data["protection"])
    ws.row_dimensions[dest_row].height = heights[source_row]


def _replicate_merges(ws, merges, source_rows: list[int], dest_rows: list[int]):
    mapping = dict(zip(source_rows, dest_rows))
    selected = set(source_rows)
    for min_r, min_c, max_r, max_c in merges:
        source_span = set(range(min_r, max_r + 1))
        if not source_span.issubset(selected):
            continue
        mapped = [mapping[r] for r in range(min_r, max_r + 1)]
        if mapped != list(range(mapped[0], mapped[0] + len(mapped))):
            continue
        ws.merge_cells(
            start_row=mapped[0], start_column=min_c,
            end_row=mapped[-1], end_column=max_c,
        )


def _add_break(ws, row_id: int):
    if row_id >= 1 and not any(int(b.id) == row_id for b in ws.row_breaks.brk if b.id is not None):
        ws.row_breaks.append(Break(id=row_id))



def _last_content_row(ws) -> int:
    for r in range(ws.max_row, 0, -1):
        if _row_has_content(ws, r):
            return r
    return 1


def _refresh_print_area(ws):
    if not ws.print_area:
        return
    try:
        text = str(ws.print_area)
        first = text.split(",")[0].replace("'", "")
        if "!" in first:
            first = first.split("!", 1)[1]
        min_col, min_row, max_col, old_max_row = range_boundaries(first.replace("$", ""))
        new_max_row = max(old_max_row, _last_content_row(ws))
        ws.print_area = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{new_max_row}"
        )
    except Exception:
        pass

def _extra_rows_for_split(table: TableInfo, sizes: list[int]) -> int:
    """Rows that must be reserved below a source table before rebuilding it."""
    prefix_len = table.label_start - table.start_row
    suffix_len = table.end_row - table.base_row + 1
    original_len = table.end_row - table.start_row + 1
    new_len = sum(prefix_len + size + suffix_len for size in sizes) + GAP_ROWS * (len(sizes) - 1)
    return max(0, new_len - original_len)


def _shift_table(table: TableInfo, offset: int) -> TableInfo:
    if offset == 0:
        return table
    return TableInfo(
        sheet=table.sheet,
        start_row=table.start_row + offset,
        title_row=table.title_row + offset,
        header_start=table.header_start + offset,
        label_start=table.label_start + offset,
        label_end=table.label_end + offset,
        base_row=table.base_row + offset,
        footer_end=table.footer_end + offset,
        legend_start=(table.legend_start + offset) if table.legend_start is not None else None,
        legend_end=(table.legend_end + offset) if table.legend_end is not None else None,
        title=table.title,
    )


def _batch_insert_rows_seguro(ws, insertions: list[tuple[int, int]]) -> None:
    """Insert multiple row blocks while remapping merges/dimensions/breaks once.

    This is materially faster on long reports with thousands of custom row
    dimensions and more than a thousand merged ranges than calling the generic
    safe insertion helper once for every split table.
    """
    normalized = [(int(pos), int(amount)) for pos, amount in insertions if amount > 0]
    if not normalized:
        return
    normalized.sort()

    merges = [(rng.min_row, rng.min_col, rng.max_row, rng.max_col) for rng in list(ws.merged_cells.ranges)]
    for rng in list(ws.merged_cells.ranges):
        try:
            ws.unmerge_cells(str(rng))
        except KeyError:
            ws.merged_cells.ranges.discard(rng)

    row_dims = {idx: ws.row_dimensions.pop(idx) for idx in list(ws.row_dimensions.keys())}
    breaks = list(ws.row_breaks.brk)
    ws.row_breaks.brk = []

    # Positions are based on the original sheet. Descending insertion keeps all
    # still-pending positions valid.
    for pos, amount in sorted(normalized, reverse=True):
        ws.insert_rows(pos, amount=amount)

    def map_row(row: int) -> int:
        return row + sum(amount for pos, amount in normalized if pos <= row)

    for idx, dim in row_dims.items():
        new_idx = map_row(int(idx))
        dim.index = new_idx
        ws.row_dimensions[new_idx] = dim

    for min_r, min_c, max_r, max_c in merges:
        new_min, new_max = min_r, max_r
        for pos, amount in normalized:
            if pos <= new_min:
                new_min += amount
                new_max += amount
            elif new_min < pos <= new_max:
                new_max += amount
        ws.merge_cells(start_row=new_min, start_column=min_c, end_row=new_max, end_column=max_c)

    for br in breaks:
        if br.id is not None:
            br.id = map_row(int(br.id))
        ws.row_breaks.brk.append(br)


def _split_one_table(ws, table: TableInfo, sizes: list[int], *, preallocated: bool = False) -> SplitResult:
    cells, heights, merges = _capture_block(ws, table)
    prefix = list(range(table.start_row, table.label_start))
    labels = list(range(table.label_start, table.label_end + 1))
    # Repeat the complete fixed tail in every part: Base + Pergunta + spacer + LEGENDA.
    suffix = list(range(table.base_row, table.end_row + 1))

    original_len = table.end_row - table.start_row + 1
    new_len = sum(len(prefix) + size + len(suffix) for size in sizes) + GAP_ROWS * (len(sizes) - 1)
    extra = new_len - original_len
    if extra > 0 and not preallocated:
        inserir_linhas_seguro(ws, table.end_row + 1, extra)

    # Any pre-existing break inside the source table is superseded by the new breaks.
    ws.row_breaks.brk = [
        b for b in ws.row_breaks.brk
        if b.id is None or not (table.start_row <= int(b.id) < table.end_row)
    ]

    _clear_region(ws, table.start_row, table.start_row + new_len - 1)

    cursor = table.start_row
    label_cursor = 0
    for part_index, size in enumerate(sizes):
        selected_labels = labels[label_cursor:label_cursor + size]
        label_cursor += size
        source_rows = prefix + selected_labels + suffix
        dest_rows = []
        for source_row in source_rows:
            _copy_snapshot_row(ws, cells, heights, source_row, cursor)
            dest_rows.append(cursor)
            cursor += 1
        # Write the continuation marker before restoring merges. In production
        # reports column A of the two-line header is vertically merged; after
        # merge restoration a non-anchor coordinate can become a read-only
        # MergedCell depending on the source geometry. Writing first lets the
        # subsequent merge preserve the value in the top-left anchor.
        if part_index > 0:
            header_offset = prefix.index(table.header_start)
            continuation_row = dest_rows[header_offset]
            continuation_cell = ws.cell(row=continuation_row, column=1)
            continuation_cell.value = CONTINUATION_TEXT.strip()
            continuation_cell.font = Font(name="DIN", size=9, bold=True)
            continuation_cell.alignment = Alignment(horizontal="left", vertical="center")

        _replicate_merges(ws, merges, source_rows, dest_rows)

        if part_index < len(sizes) - 1:
            gap_start = cursor
            for rr in range(gap_start, gap_start + GAP_ROWS):
                ws.row_dimensions[rr].height = DEFAULT_ROW_HEIGHT
            _add_break(ws, gap_start + BREAK_AFTER_GAP_ROWS - 1)
            cursor += GAP_ROWS

    ws.row_breaks.brk = sorted(ws.row_breaks.brk, key=lambda b: int(b.id or 0))
    _refresh_print_area(ws)
    return SplitResult(
        sheet=ws.title,
        title=table.title,
        original_start=table.start_row,
        original_end=table.end_row,
        labels=table.label_count,
        parts=len(sizes),
        part_sizes=sizes,
        reason="quebra de pagina dentro da tabela",
    )


def _breaks_inside_table(table: TableInfo, breaks: Iterable[int]) -> list[int]:
    """Return page breaks that split the physical footprint of a table.

    Excel stores a horizontal break id as the row *after which* the page ends.
    A break is therefore internal whenever it occurs after any row from the
    table start up to the penultimate row of the table. This deliberately
    includes breaks after Base, Pergunta, blank spacer rows and LEGENDA rows.
    """
    return sorted({int(b) for b in breaks if table.start_row <= int(b) < table.end_row})


def analisar_planilha(ws) -> dict:
    # This module only acts on legend-bearing regional/capital tables. Using
    # the legend-anchored detector makes long reports substantially faster.
    tables = detectar_tabelas_regionais(ws)
    manual = sorted({int(b.id) for b in ws.row_breaks.brk if b.id is not None})
    automatic = estimar_quebras_automaticas(ws)
    effective = sorted(set(manual + automatic))
    flagged = []
    for table in tables:
        if not table.has_legend:
            continue
        crossing = _breaks_inside_table(table, effective)
        if crossing:
            flagged.append({
                "table": table,
                "breaks": crossing,
                "manual_breaks": [b for b in crossing if b in manual],
                "estimated_breaks": [b for b in crossing if b in automatic],
            })
    return {
        "sheet": ws.title,
        "tables": tables,
        "manual_breaks": manual,
        "estimated_breaks": automatic,
        "flagged": flagged,
    }


def processar_workbook(wb) -> dict:
    """Analyze report sheets and split only legend-bearing tables that cross pages."""
    analyses = []
    results: list[SplitResult] = []

    # Analyze first so the UI can report what existed before modifications.
    for ws in wb.worksheets:
        if ws.title.lower() == "sumario":
            continue
        analyses.append(analisar_planilha(ws))

    # Prepare all splits on the original geometry first. On long reports,
    # reserving all additional rows in one batch avoids repeatedly remapping
    # thousands of merges, row dimensions and page breaks.
    for analysis in analyses:
        ws = wb[analysis["sheet"]]
        effective = sorted(set(analysis["manual_breaks"] + analysis["estimated_breaks"]))
        prepared: list[tuple[TableInfo, list[int], int]] = []
        for item in sorted(analysis["flagged"], key=lambda item: item["table"].start_row):
            table = item["table"]
            sizes = _choose_part_sizes(ws, table, effective)
            if len(sizes) <= 1:
                continue
            extra = _extra_rows_for_split(table, sizes)
            prepared.append((table, sizes, extra))

        insertions = [(table.end_row + 1, extra) for table, _, extra in prepared if extra > 0]
        _batch_insert_rows_seguro(ws, insertions)

        # After batch reservation, tables below earlier insertions have moved.
        # No more row insertions happen while rebuilding, so these adjusted
        # coordinates remain stable for the entire pass.
        for table, sizes, _extra in sorted(prepared, key=lambda item: item[0].start_row):
            offset = sum(amount for pos, amount in insertions if pos <= table.start_row)
            adjusted = _shift_table(table, offset)
            results.append(_split_one_table(ws, adjusted, sizes, preallocated=True))

    results.sort(key=lambda x: (x.sheet, x.original_start))
    return {
        "sheets_analyzed": len(analyses),
        "tables_detected": sum(len(a["tables"]) for a in analyses),
        "tables_with_legend": sum(
            1 for a in analyses for table in a["tables"] if table.has_legend
        ),
        "tables_flagged": sum(len(a["flagged"]) for a in analyses),
        "tables_split": len(results),
        "parts_created": sum(r.parts for r in results),
        "results": results,
        "analyses": analyses,
    }
