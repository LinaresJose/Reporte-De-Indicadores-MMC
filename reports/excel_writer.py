"""reports/excel_writer.py — LYNX Report System.

Generates the consolidated Excel workbook (one ``General`` sheet plus one
sheet per region) that replicates the visual design of the original
``Reporte_de_indicadores_de_ventas_*.xlsm``.

Design conventions implemented
-------------------------------
- **Header row:** dark-blue fill (#1F3864), bold white text, centered.
- **Data rows:** alternating white/light-grey with colored alcance cells:
    - Green  (#C6EFCE) if alcance ≥ 100 %
    - Yellow (#FFEB9C) if 80 % ≤ alcance < 100 %
    - Red    (#FFC7CE) if alcance < 80 %
- **Totals row:** black fill (#000000), bold white text (the "black pill").
- **Working-days box:** 4-row × 2-col table below the totals row.
- **Regional sheets:** ``Región`` column hidden; tab colour green (#00B050).
- **Number formats:** currencies → ``#,##0.00``; percentages → ``0.00%``;
  integers → ``#,##0``; ranking → plain integer.

Typical usage::

    from reports.excel_writer import generate_excel_report

    output_path, regional_paths = generate_excel_report(
        records, app_config, dias_info, Path("output")
    )
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import config
from core.engine import calculate_rankings, filter_by_region, get_regional_totals
from core.models import AppConfig, VendorRecord, WorkingDaysInfo

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Colour palette  (matches the xlsm design)
# ---------------------------------------------------------------------------

_CLR_HEADER_BG = "1F3864"   # dark navy blue
_CLR_HEADER_FG = "FFFFFF"   # white text
_CLR_TOTALS_BG = "000000"   # black
_CLR_TOTALS_FG = "FFFFFF"   # white text
_CLR_ALT_ROW   = "F2F2F2"   # very light grey (alternating rows)
_CLR_GREEN_BG  = "C6EFCE"   # alcance ≥ 100 %
_CLR_GREEN_FG  = "276221"
_CLR_YELLOW_BG = "FFEB9C"   # 80 % ≤ alcance < 100 %
_CLR_YELLOW_FG = "9C5700"
_CLR_RED_BG    = "FFC7CE"   # alcance < 80 %
_CLR_RED_FG    = "9C0006"
_CLR_TAB_GREEN = "00B050"   # regional sheet tab colour

# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------


def _make_fill(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _make_font(bold: bool = False, color: str = "000000", size: int = 10) -> Font:
    return Font(bold=bold, color=color, size=size, name="Calibri")


def _make_border() -> Border:
    thin = Side(border_style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _make_center_align(wrap: bool = False) -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


# Pre-built styles
_FILL_HEADER  = _make_fill(_CLR_HEADER_BG)
_FILL_TOTALS  = _make_fill(_CLR_TOTALS_BG)
_FILL_ALT     = _make_fill(_CLR_ALT_ROW)
_FILL_WHITE   = _make_fill("FFFFFF")
_FILL_GREEN   = _make_fill(_CLR_GREEN_BG)
_FILL_YELLOW  = _make_fill(_CLR_YELLOW_BG)
_FILL_RED     = _make_fill(_CLR_RED_BG)

_FONT_HEADER  = _make_font(bold=True, color=_CLR_HEADER_FG, size=10)
_FONT_TOTALS  = _make_font(bold=True, color=_CLR_TOTALS_FG, size=10)
_FONT_BODY    = _make_font(size=10)
_FONT_BODY_B  = _make_font(bold=True, size=10)
_FONT_GREEN   = _make_font(bold=True, color=_CLR_GREEN_FG, size=10)
_FONT_YELLOW  = _make_font(bold=True, color=_CLR_YELLOW_FG, size=10)
_FONT_RED     = _make_font(bold=True, color=_CLR_RED_FG, size=10)

_BORDER       = _make_border()
_ALIGN_CENTER = _make_center_align()
_ALIGN_LEFT   = Alignment(horizontal="left", vertical="center")
_ALIGN_RIGHT  = Alignment(horizontal="right", vertical="center")

# ---------------------------------------------------------------------------
# Number format strings
# ---------------------------------------------------------------------------

_FMT_CURRENCY = '#,##0.00'
_FMT_INTEGER  = '#,##0'
_FMT_PERCENT  = '0.00%'
_FMT_PLAIN    = 'General'


def _col_format(col_index: int) -> str:
    """Return the Excel number format string for a given 1-based column index.

    The column index corresponds to the position in ``config.GENERAL_HEADERS``.

    Args:
        col_index: 1-based column position.

    Returns:
        An Excel number format string.
    """
    # Map header names to format codes
    header = config.GENERAL_HEADERS[col_index - 1] if col_index <= len(config.GENERAL_HEADERS) else ""
    currency_cols = {
        "Monto Pedidos", "Devoluciones", "Total Facturado", "Ventas Netas",
        "Monto por Procesar", "BackOrder", "Presup. Ventas", "Diferencia Ventas",
    }
    pct_cols = {"Alcance Ventas", "Alcance Activación", "Alcance SKUs", "% Sin Vencer"}
    int_cols = {
        "Cant. Pedidos", "Clientes Activados", "Presup. Activación",
        "Diferencia Activación", "SKUs Vendidos", "Presup. SKUs",
        "Diferencia SKUs", "Presup. SKUs",
        "Lugar Ventas", "Lugar Activación", "Lugar SKUs", "Lugar Indicadores",
    }
    pts_cols = {
        "Puntos Ventas (30)", "Puntos Activación (10)", "Puntos SKUs (10)",
        "Puntos Cobranza (50)", "Total Indicadores",
    }

    if header in currency_cols:
        return _FMT_CURRENCY
    if header in pct_cols:
        return _FMT_PERCENT
    if header in int_cols:
        return _FMT_INTEGER
    if header in pts_cols:
        return '#,##0.00'
    return _FMT_PLAIN


def _alcance_fill_font(alcance: float) -> Tuple[PatternFill, Font]:
    """Return (fill, font) for an alcance value cell.

    Args:
        alcance: The achievement ratio (0.0 – any positive float, where 1.0 = 100 %).

    Returns:
        Tuple of ``(PatternFill, Font)`` reflecting the KPI colour convention.
    """
    if alcance >= 1.0:
        return _FILL_GREEN, _FONT_GREEN
    if alcance >= 0.80:
        return _FILL_YELLOW, _FONT_YELLOW
    return _FILL_RED, _FONT_RED


# ---------------------------------------------------------------------------
# Low-level cell writers
# ---------------------------------------------------------------------------


def _apply_header_style(ws: Worksheet, row: int, n_cols: int) -> None:
    """Style an entire row as a dark-blue header.

    Args:
        ws: Target worksheet.
        row: 1-based row number.
        n_cols: Number of columns to style.
    """
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = _FILL_HEADER
        cell.font = _FONT_HEADER
        cell.alignment = _ALIGN_CENTER
        cell.border = _BORDER


def _apply_totals_style(ws: Worksheet, row: int, n_cols: int) -> None:
    """Style an entire row as the black totals pill.

    Args:
        ws: Target worksheet.
        row: 1-based row number.
        n_cols: Number of columns to style (should match the data columns).
    """
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = _FILL_TOTALS
        cell.font = _FONT_TOTALS
        cell.alignment = _ALIGN_CENTER
        cell.border = _BORDER


def _write_data_rows(
    ws: Worksheet,
    records: List[VendorRecord],
    start_row: int,
    hide_region_col: bool = False,
) -> int:
    """Write vendor data rows into *ws* starting at *start_row*.

    Alcance columns (Alcance Ventas, Alcance Activación, Alcance SKUs,
    % Sin Vencer) receive traffic-light colour coding.

    Args:
        ws: Target worksheet.
        records: Ordered list of :class:`~core.models.VendorRecord` objects.
        start_row: 1-based row number for the first data row.
        hide_region_col: If ``True``, the ``Región`` column (col 3) value is
            written but the column remains hidden (set by the caller).

    Returns:
        The 1-based row number of the row *after* the last data row
        (ready for the totals row).
    """
    # Column indices that hold alcance values (1-based, matching GENERAL_HEADERS)
    alcance_col_indices = {
        config.GENERAL_HEADERS.index("Alcance Ventas") + 1,
        config.GENERAL_HEADERS.index("Alcance Activación") + 1,
        config.GENERAL_HEADERS.index("Alcance SKUs") + 1,
        config.GENERAL_HEADERS.index("% Sin Vencer") + 1,
    }

    for i, rec in enumerate(records):
        row = start_row + i
        is_alt = (i % 2 == 1)
        row_data = rec.to_row()

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.border = _BORDER
            cell.alignment = _ALIGN_CENTER

            if col_idx in alcance_col_indices and isinstance(value, (int, float)):
                fill, font = _alcance_fill_font(float(value))
                cell.fill = fill
                cell.font = font
            else:
                cell.fill = _FILL_ALT if is_alt else _FILL_WHITE
                cell.font = _FONT_BODY

            # Apply number format
            cell.number_format = _col_format(col_idx)

        # First column (name) left-aligned
        ws.cell(row=row, column=1).alignment = _ALIGN_LEFT

    return start_row + len(records)


def _write_totals_row(
    ws: Worksheet,
    row: int,
    totals: Dict[str, float],
    n_cols: int,
) -> None:
    """Write the black totals pill row.

    Maps the engine's internal keys in *totals* to the 34 columns defined in
    ``config.GENERAL_HEADERS``.

    Args:
        ws: Target worksheet.
        row: 1-based row number.
        totals: Dictionary of ``{metric_name: value}`` as returned by
            :func:`~core.engine.get_regional_totals`.
        n_cols: Total number of data columns (34).
    """
    _apply_totals_style(ws, row, n_cols)

    # Col 1: Label "TOTALES"
    cell_tot = ws.cell(row=row, column=1, value="TOTALES")
    cell_tot.alignment = _ALIGN_LEFT
    cell_tot.font = _FONT_TOTALS

    # Mapping from column index (1-based) to key in totals dict
    col_mapping: Dict[int, str] = {
        2:  "zona_count",
        4:  "cantidad_pedidos",
        5:  "monto_pedidos",
        6:  "devoluciones",
        7:  "total_facturado",
        8:  "ventas_netas",
        9:  "monto_por_procesar",
        10: "backorder",
        11: "presup_ventas",
        12: "diferencia_ventas",
        13: "alcance_ventas",
        14: "puntos_ventas",
        15: "clientes_activados",
        16: "presup_activacion",
        17: "diferencia_activacion",
        18: "alcance_activacion",
        19: "puntos_activacion",
        20: "skus_vendidos",
        21: "presup_skus",
        22: "diferencia_skus",
        23: "alcance_skus",
        24: "puntos_skus",
        25: "pct_sin_vencer",
        26: "puntos_cobranza",
        27: "total_indicadores",
        28: "incentivo_count",
        33: "rifa_count",
        34: "participacion",
    }

    # Format mapping for alcance in totals
    alcance_cols = {13, 18, 23, 25}

    for col_idx in range(2, n_cols + 1):
        cell = ws.cell(row=row, column=col_idx)
        key = col_mapping.get(col_idx)

        if key and key in totals:
            val = totals[key]
            cell.value = val

            # Check if this column is an alcance to colorize or style appropriately
            if col_idx in alcance_cols and isinstance(val, (int, float)):
                fill, font = _alcance_fill_font(float(val))
                cell.fill = fill
                cell.font = font
            else:
                cell.fill = _FILL_TOTALS
                cell.font = _FONT_TOTALS

            cell.alignment = _ALIGN_CENTER
            cell.number_format = _col_format(col_idx)
        else:
            # Empty / blank cell in totals (e.g. Region, Rankings)
            cell.value = None
            cell.fill = _FILL_TOTALS
            cell.font = _FONT_TOTALS


def _write_working_days_box(
    ws: Worksheet,
    start_row: int,
    start_col: int,
    dias_info: WorkingDaysInfo,
) -> None:
    """Write the 4-row × 2-col working-days information box.

    Layout::

        | Días hábiles mes    | <total_dias_habiles> |
        | Días hábiles elaps. | <dias_habiles_trans>  |
        | Días hábiles rest.  | <dias_habiles_rest>   |
        | Avance del mes      | <pct_avance>%         |

    Args:
        ws: Target worksheet.
        start_row: 1-based row for the first box row.
        start_col: 1-based column for the left label column.
        dias_info: :class:`~core.models.WorkingDaysInfo` with working day counts.
    """
    dias_restantes = max(0, dias_info.dias_habiles_mes - dias_info.dias_habiles_transcurridos)
    pct_avance = (
        dias_info.dias_habiles_transcurridos / dias_info.dias_habiles_mes
        if dias_info.dias_habiles_mes > 0
        else 0.0
    )
    labels = [
        ("Días hábiles del mes", dias_info.dias_habiles_mes),
        ("Días hábiles transcurridos", dias_info.dias_habiles_transcurridos),
        ("Días hábiles restantes", dias_restantes),
        ("Avance del mes", pct_avance),
    ]

    box_fill   = _make_fill("DCE6F1")  # light blue
    box_font_l = _make_font(bold=True, size=9)
    box_font_v = _make_font(size=9)
    box_border = _make_border()

    for i, (label, value) in enumerate(labels):
        label_cell = ws.cell(row=start_row + i, column=start_col, value=label)
        label_cell.fill   = box_fill
        label_cell.font   = box_font_l
        label_cell.border = box_border
        label_cell.alignment = _ALIGN_LEFT

        value_cell = ws.cell(row=start_row + i, column=start_col + 1, value=value)
        value_cell.fill   = box_fill
        value_cell.font   = box_font_v
        value_cell.border = box_border
        value_cell.alignment = _ALIGN_CENTER

        # Format the last row as percentage
        if i == 3:
            value_cell.number_format = _FMT_PERCENT


def _set_column_widths(ws: Worksheet, headers: List[str]) -> None:
    """Set reasonable column widths based on header names.

    Args:
        ws: Target worksheet.
        headers: List of header strings in column order.
    """
    wide_cols = {
        "Nombre": 28,
        "Zona": 22,
        "Región": 16,
        "Incentivo": 12,
        "Rifa": 8,
        "Participación": 14,
    }
    for i, h in enumerate(headers, start=1):
        col_letter = get_column_letter(i)
        width = wide_cols.get(h, 14)
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = 30  # header row taller


# ---------------------------------------------------------------------------
# Sheet writers
# ---------------------------------------------------------------------------


def write_general_sheet(
    wb: Workbook,
    records: List[VendorRecord],
    dias_info: WorkingDaysInfo,
    mes_label: str,
) -> Worksheet:
    """Create and populate the ``General`` worksheet.

    Args:
        wb: :class:`openpyxl.Workbook` to add the sheet to.
        records: All vendor records sorted by ranking.
        dias_info: Working-days metadata for the working-days box.
        mes_label: Month label string (e.g. ``"Noviembre 2026"``).

    Returns:
        The newly created :class:`~openpyxl.worksheet.worksheet.Worksheet`.
    """
    ws = wb.create_sheet("General")
    headers = config.GENERAL_HEADERS
    n_cols = len(headers)

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1, value=f"REPORTE DE INDICADORES DE VENTAS — {mes_label.upper()}")
    title_cell.fill   = _FILL_HEADER
    title_cell.font   = _make_font(bold=True, color=_CLR_HEADER_FG, size=12)
    title_cell.alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 22

    # Header row (row 2)
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=col_idx, value=header)
    _apply_header_style(ws, 2, n_cols)

    # Data rows
    next_row = _write_data_rows(ws, records, start_row=3)

    # Totals row
    totals = get_regional_totals(records)
    _write_totals_row(ws, next_row, totals, n_cols)
    next_row += 1

    # Working-days box (2 columns after the data, below totals)
    _write_working_days_box(ws, next_row + 1, 1, dias_info)

    # Freeze header rows
    ws.freeze_panes = "A3"

    # Auto-filter
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"

    # Column widths
    _set_column_widths(ws, headers)

    logger.debug("General sheet written: %d vendors", len(records))
    return ws


def write_regional_sheet(
    wb: Workbook,
    records: List[VendorRecord],
    region: str,
    dias_info: WorkingDaysInfo,
    mes_label: str,
) -> Worksheet:
    """Create and populate a regional worksheet.

    Similar to :func:`write_general_sheet` but:
    - Only includes vendors belonging to *region*.
    - Hides the ``Región`` column (column 3).
    - Sets tab colour to green.

    Args:
        wb: :class:`openpyxl.Workbook` to add the sheet to.
        records: All vendor records (will be filtered to *region* internally).
        region: Region name string (e.g. ``"CAPITAL"``).
        dias_info: Working-days metadata.
        mes_label: Month label string.

    Returns:
        The newly created regional :class:`~openpyxl.worksheet.worksheet.Worksheet`.
    """
    regional_records = filter_by_region(records, region)
    ws = wb.create_sheet(region[:31])  # sheet names max 31 chars
    ws.sheet_properties.tabColor = _CLR_TAB_GREEN

    headers = config.GENERAL_HEADERS
    n_cols = len(headers)

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(
        row=1, column=1,
        value=f"INDICADORES DE VENTAS — {region.upper()} — {mes_label.upper()}"
    )
    title_cell.fill      = _FILL_HEADER
    title_cell.font      = _make_font(bold=True, color=_CLR_HEADER_FG, size=12)
    title_cell.alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 22

    # Header row
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=col_idx, value=header)
    _apply_header_style(ws, 2, n_cols)

    # Data rows
    next_row = _write_data_rows(ws, regional_records, start_row=3)

    # Totals row
    totals = get_regional_totals(regional_records)
    _write_totals_row(ws, next_row, totals, n_cols)
    next_row += 1

    # Working-days box
    _write_working_days_box(ws, next_row + 1, 1, dias_info)

    # Hide the Región column (col 3)
    try:
        region_col_idx = config.GENERAL_HEADERS.index("Región") + 1
        ws.column_dimensions[get_column_letter(region_col_idx)].hidden = True
    except ValueError:
        pass  # Column not found — harmless

    # Freeze & filter
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"
    _set_column_widths(ws, headers)

    logger.debug("Regional sheet '%s': %d vendors", region, len(regional_records))
    return ws


# ---------------------------------------------------------------------------
# Master generator
# ---------------------------------------------------------------------------


def generate_excel_report(
    records: List[VendorRecord],
    app_config: AppConfig,
    dias_info: WorkingDaysInfo,
    output_dir: Path,
) -> Tuple[Path, Dict[str, Path]]:
    """Generate the full consolidated Excel workbook and individual regional files.

    Workflow:

    1. Apply rankings to *records* via :func:`~core.engine.calculate_rankings`.
    2. Create a consolidated ``Reporte_General_<mes>.xlsx`` with:
       - ``General`` sheet (all vendors).
       - One sheet per region.
    3. Create individual regional ``.xlsx`` files in ``output_dir/REGIONALES/``.
    4. Return paths for manifest generation.

    Args:
        records: Fully computed vendor records (output of
            :func:`~core.engine.calculate_general`).
        app_config: Application configuration (contains ``mes_label``,
            ``regions``, etc.).
        dias_info: Working-days information.
        output_dir: Root output directory.

    Returns:
        A tuple ``(consolidated_path, regional_paths)`` where:

        - ``consolidated_path``: Path to the consolidated ``.xlsx`` file.
        - ``regional_paths``: ``{region_name: path_to_regional_xlsx}`` dict.

    Raises:
        :class:`~core.exceptions.ReportError`: If the workbook cannot be saved.

    Example::

        consolidated, regionals = generate_excel_report(
            records, app_config, dias_info, Path("output")
        )
        print(f"Saved: {consolidated}")
    """
    from core.exceptions import ReportError

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    regionales_dir = output_dir / "REGIONALES"
    regionales_dir.mkdir(parents=True, exist_ok=True)

    mes_label = app_config.mes_label
    safe_mes  = mes_label.replace(" ", "_")

    # Apply rankings
    records = calculate_rankings(records)

    # ----------------------------------------------------------------
    # 1. Consolidated workbook
    # ----------------------------------------------------------------
    wb = Workbook()
    # Remove default empty sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    write_general_sheet(wb, records, dias_info, mes_label)

    for region in app_config.regions:
        write_regional_sheet(wb, records, region, dias_info, mes_label)

    consolidated_path = output_dir / f"Reporte_General_{safe_mes}.xlsx"
    try:
        wb.save(consolidated_path)
        logger.info("Consolidated workbook saved: %s", consolidated_path)
    except Exception as exc:
        raise ReportError(
            f"No se pudo guardar el reporte consolidado: {exc}",
            report_type="excel_consolidado",
        ) from exc

    # ----------------------------------------------------------------
    # 2. Individual regional workbooks
    # ----------------------------------------------------------------
    regional_paths: Dict[str, Path] = {}

    for region in app_config.regions:
        regional_records = filter_by_region(records, region)
        if not regional_records:
            logger.warning("No vendors found for region '%s'; skipping.", region)
            continue

        wb_reg = Workbook()
        if "Sheet" in wb_reg.sheetnames:
            del wb_reg["Sheet"]

        write_regional_sheet(wb_reg, records, region, dias_info, mes_label)

        safe_region = region.replace(" ", "_")
        reg_path = regionales_dir / f"{safe_region}_Indicadores_{safe_mes}.xlsx"
        try:
            wb_reg.save(reg_path)
            regional_paths[region] = reg_path
            logger.info("Regional workbook saved: %s", reg_path)
        except Exception as exc:
            raise ReportError(
                f"No se pudo guardar el reporte regional '{region}': {exc}",
                report_type="excel_regional",
            ) from exc

    return consolidated_path, regional_paths
