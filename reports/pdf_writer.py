"""reports/pdf_writer.py — LYNX Report System.

Generates individual PDF performance reports for each vendor, replicating
exactly the visual design, colors, geometry, and logo of the
"Individual Ventas" sheet from the original Excel macro workbook.

Layout features:
- Sheet Header: MES, ZONA, COMISIONISTA (with yellow highlight) + MMC black box logo.
- 5 main Indicator blocks:
  1. VENTAS (PRESUPUESTO, VENTAS NETAS, ALCANCE, PUNTOS (30), POSICION)
  2. ACTIVACION (PRESUPUESTO, CLIENTES ACTIVOS, ALCANCE, PUNTOS (10), POSICION)
  3. SKUS (PRESUPUESTO, SKUS NETOS, ALCANCE, PUNTOS (10), POSICION)
  4. COBRANZA (% SIN VENCER, % 0 A 40 DIAS, PUNTOS (50))
  5. TOTAL INDICADORES (TOTAL PUNTOS, POSICION)
- Bonus blocks:
  - INCENTIVO (0% o 1%) + descriptive rule box
  - CONCURSO (SI o NO) + descriptive rule box
- Outer border enclosing the whole layout (just like the Excel print area).
"""

from __future__ import annotations

import logging
import os
import re
import unicodedata
from pathlib import Path
from typing import List, Optional

from fpdf import FPDF

from core.exceptions import ReportError
from core.models import AppConfig, VendorRecord

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants & Colors (RGB)
# ---------------------------------------------------------------------------

_PDF_SUBDIR = "PDF_INDICADORES"

# Colors matching the original Excel "Individual Ventas"
_CLR_BLACK       = (0, 0, 0)
_CLR_WHITE       = (255, 255, 255)
_CLR_HEADER_GRAY = (220, 220, 220)       # Light gray for sub-headers
_CLR_YELLOW      = (255, 255, 0)         # Yellow #FFFF00 for COMISIONISTA
_CLR_PINK_ALERT  = (255, 199, 206)       # Soft pink/light red for < threshold alcance
_CLR_GREEN_ALERT = (198, 239, 206)       # Soft green for >= threshold alcance
_CLR_LINE        = (0, 0, 0)             # Black grid lines

# Default path for the extracted MMC logo
_LOGO_PATH = Path(__file__).parent.parent / "data" / "image1.png"


def _normalize_filename(text: str) -> str:
    """Convert text to a filesystem-safe ASCII string."""
    nfd = unicodedata.normalize("NFD", text)
    ascii_text = nfd.encode("ascii", "ignore").decode("ascii")
    safe = re.sub(r'[\\/:*?"<>|]', "", ascii_text)
    safe = re.sub(r"[\s\-]+", "_", safe.strip())
    safe = re.sub(r"[^\w]", "", safe)
    return safe or "SIN_NOMBRE"


def _sanitize(text: str) -> str:
    """Sanitize strings for latin-1 / standard PDF Helvetica."""
    if not isinstance(text, str):
        text = str(text)
    nfd = unicodedata.normalize("NFKD", text)
    clean = "".join(c for c in nfd if not unicodedata.combining(c))
    return (
        clean.replace("—", "-")
        .replace("–", "-")
        .replace("✓", "[X]")
        .replace("✗", "[-]")
        .encode("ascii", "replace")
        .decode("ascii")
    )


def _fmt_money(val: float) -> str:
    """Format money value: hyphen if 0 or None, else comma-separated 2 decimals."""
    if val is None or abs(val) < 0.001:
        return "-"
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_int(val: float | int) -> str:
    """Format integer value: hyphen if 0, else integer string."""
    if val is None or int(val) == 0:
        return "-"
    return f"{int(val):,}".replace(",", ".")


def _fmt_pct(val: float) -> str:
    """Format percentage: e.g. 0.8945 -> 89,45%."""
    if val is None:
        return "0,00%"
    pct = val * 100.0
    return f"{pct:.2f}%".replace(".", ",")


def _fmt_pts(val: float) -> str:
    """Format points: e.g. 30.0 -> 30,00."""
    if val is None:
        return "0,00"
    return f"{val:.2f}".replace(".", ",")


# ---------------------------------------------------------------------------
# PDF Generator Class
# ---------------------------------------------------------------------------


class IndividualVentasPDF(FPDF):
    """Generates the exact 1-page PDF replica of Individual Ventas."""

    def __init__(self) -> None:
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_auto_page_break(auto=False)
        self.set_margins(15, 15, 15)

    def generate_page(
        self,
        record: VendorRecord,
        mes_label: str,
        logo_path: Optional[Path] = None,
    ) -> None:
        self.add_page()

        # Geometry offsets
        start_x = 22.0
        start_y = 35.0

        # Column widths
        w_a = 34.0   # Col A: Left category title (VENTAS, etc.)
        w_b = 40.0   # Col B: Sub-item title (PRESUPUESTO, ALCANCE, etc.)
        w_c = 28.0   # Col C: Values (numeric / percentage)
        w_de = 64.0  # Cols D+E: Right side banner / rule text
        total_w = w_a + w_b + w_c + w_de  # 166.0 mm

        h_row = 6.0  # Standard row height

        self.set_line_width(0.3)
        self.set_draw_color(*_CLR_LINE)

        # ===================================================================
        # HEADER BLOCK (Rows 1 to 3)
        # ===================================================================
        curr_y = start_y

        # Row 1: MES
        # A1: MES
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, h_row, _sanitize("MES"), border=1, fill=True, align="L")

        # B1: Mes Value
        self.set_fill_color(*_CLR_WHITE)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_b, h_row, _sanitize(mes_label), border=1, fill=True, align="C")

        # Logo Box (C1:E2 merged -> w_c + w_de = 92 mm, height = 2*h_row = 12 mm)
        logo_box_w = w_c + w_de
        logo_box_h = h_row * 2
        self.set_xy(start_x + w_a + w_b, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.rect(start_x + w_a + w_b, curr_y, logo_box_w, logo_box_h, style="DF")

        # Insert MMC Logo image inside the box if available
        use_logo = logo_path if (logo_path and logo_path.exists()) else _LOGO_PATH
        if use_logo and use_logo.exists():
            img_w = 48.0
            img_h = 13.0
            img_x = (start_x + w_a + w_b) + (logo_box_w - img_w) / 2.0
            img_y = curr_y - 0.5
            self.image(str(use_logo), x=img_x, y=img_y, w=img_w, h=img_h)
        else:
            # Fallback text logo
            self.set_xy(start_x + w_a + w_b, curr_y)
            self.set_font("Helvetica", "B", 14)
            self.set_text_color(*_CLR_WHITE)
            self.cell(logo_box_w, logo_box_h, _sanitize("MMC"), align="C")

        # Row 2: ZONA
        curr_y += h_row
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, h_row, _sanitize("ZONA"), border=1, fill=True, align="L")

        self.set_fill_color(*_CLR_WHITE)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "", 9)
        self.cell(w_b, h_row, _sanitize(str(record.zona)), border=1, fill=True, align="C")

        # Row 3: COMISIONISTA
        curr_y += h_row
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, h_row, _sanitize("COMISIONISTA"), border=1, fill=True, align="L")

        # Yellow Comisionista Name (merged B3:E3 -> total_w - w_a = 132 mm)
        self.set_fill_color(*_CLR_YELLOW)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 9)
        self.cell(total_w - w_a, h_row, _sanitize(record.vendor_name), border=1, fill=True, align="C")

        # Row 4 is blank in Excel, jump down
        curr_y += h_row + 4.0

        # Row 5: Header INDICADOR
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_WHITE)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, h_row, _sanitize("INDICADOR"), border=1, fill=False, align="C")

        curr_y += h_row

        # ===================================================================
        # HELPER FOR 5-ROW INDICATOR BLOCKS (VENTAS, ACTIVACION, SKUS)
        # ===================================================================
        def draw_5row_block(
            cat_name: str,
            label1: str, val1: str,
            label2: str, val2: str,
            label3: str, val3: str, is_alcance: bool, alcance_ratio: float,
            label4: str, val4: str,
            label5: str, val5: str,
            y_pos: float,
        ) -> float:
            block_h = h_row * 5

            # Left Category Cell (Col A, merged 5 rows, Black fill, White text)
            self.set_xy(start_x, y_pos)
            self.set_fill_color(*_CLR_BLACK)
            self.set_text_color(*_CLR_WHITE)
            self.set_font("Helvetica", "B", 9)
            self.cell(w_a, block_h, _sanitize(cat_name), border=1, fill=True, align="C")

            # 5 Sub-rows
            sub_rows = [
                (label1, val1, False),
                (label2, val2, False),
                (label3, val3, is_alcance),
                (label4, val4, False),
                (label5, val5, False),
            ]

            for i, (lbl, val, alc) in enumerate(sub_rows):
                ry = y_pos + (i * h_row)

                # Col B: Sub-item title (Light gray fill)
                self.set_xy(start_x + w_a, ry)
                self.set_fill_color(*_CLR_HEADER_GRAY)
                self.set_text_color(*_CLR_BLACK)
                self.set_font("Helvetica", "B", 8)
                self.cell(w_b, h_row, _sanitize(lbl), border=1, fill=True, align="L")

                # Col C: Value
                self.set_xy(start_x + w_a + w_b, ry)
                if alc:
                    # Pink highlight if below 90% (0.895), soft green if >= 90%
                    bg_color = _CLR_GREEN_ALERT if alcance_ratio >= 0.895 else _CLR_PINK_ALERT
                    self.set_fill_color(*bg_color)
                    self.cell(w_c, h_row, _sanitize(val), border=1, fill=True, align="R")
                else:
                    self.set_fill_color(*_CLR_WHITE)
                    self.cell(w_c, h_row, _sanitize(val), border=1, fill=False, align="R")

            return y_pos + block_h + 3.0

        # -------------------------------------------------------------------
        # 1. VENTAS BLOCK
        # -------------------------------------------------------------------
        curr_y = draw_5row_block(
            cat_name="VENTAS",
            label1="PRESUPUESTO", val1=_fmt_money(record.presup_ventas),
            label2="VENTAS NETAS", val2=_fmt_money(record.ventas_netas),
            label3="ALCANCE", val3=_fmt_pct(record.alcance_ventas), is_alcance=True, alcance_ratio=record.alcance_ventas,
            label4="PUNTOS (30)", val4=_fmt_pts(record.puntos_ventas),
            label5="POSICION", val5=str(record.lugar_ventas) if record.lugar_ventas > 0 else "0",
            y_pos=curr_y,
        )

        # -------------------------------------------------------------------
        # 2. ACTIVACION BLOCK
        # -------------------------------------------------------------------
        curr_y = draw_5row_block(
            cat_name="ACTIVACION",
            label1="PRESUPUESTO", val1=_fmt_int(record.presup_activacion),
            label2="CLIENTES ACTIVOS", val2=_fmt_int(record.clientes_activados),
            label3="ALCANCE", val3=_fmt_pct(record.alcance_activacion), is_alcance=True, alcance_ratio=record.alcance_activacion,
            label4="PUNTOS (10)", val4=_fmt_pts(record.puntos_activacion),
            label5="POSICION", val5=str(record.lugar_activacion) if record.lugar_activacion > 0 else "0",
            y_pos=curr_y,
        )

        # -------------------------------------------------------------------
        # 3. SKUS BLOCK
        # -------------------------------------------------------------------
        curr_y = draw_5row_block(
            cat_name="SKUS",
            label1="PRESUPUESTO", val1=_fmt_int(record.presup_skus),
            label2="SKUS NETOS", val2=_fmt_int(record.skus_vendidos),
            label3="ALCANCE", val3=_fmt_pct(record.alcance_skus), is_alcance=True, alcance_ratio=record.alcance_skus,
            label4="PUNTOS (10)", val4=_fmt_pts(record.puntos_skus),
            label5="POSICION", val5=str(record.lugar_skus) if record.lugar_skus > 0 else "0",
            y_pos=curr_y,
        )

        # -------------------------------------------------------------------
        # 4. COBRANZA BLOCK (3 rows: % SIN VENCER, % 0 A 40 DIAS, PUNTOS (50))
        # -------------------------------------------------------------------
        cob_h = h_row * 3
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, cob_h, _sanitize("COBRANZA"), border=1, fill=True, align="C")

        # Row 1: % SIN VENCER
        self.set_xy(start_x + w_a, curr_y)
        self.set_fill_color(*_CLR_HEADER_GRAY)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 8)
        self.cell(w_b, h_row, _sanitize("% SIN VENCER"), border=1, fill=True, align="L")

        self.set_xy(start_x + w_a + w_b, curr_y)
        pct_color = _CLR_GREEN_ALERT if record.pct_sin_vencer >= 0.895 else _CLR_PINK_ALERT
        self.set_fill_color(*pct_color)
        self.cell(w_c, h_row, _fmt_pct(record.pct_sin_vencer), border=1, fill=True, align="R")

        # Row 2: % 0 A 40 DIAS (Blank in master sheet)
        self.set_xy(start_x + w_a, curr_y + h_row)
        self.set_fill_color(*_CLR_HEADER_GRAY)
        self.set_font("Helvetica", "B", 8)
        self.cell(w_b, h_row, _sanitize("% 0 A 40 DIAS"), border=1, fill=True, align="L")
        self.set_xy(start_x + w_a + w_b, curr_y + h_row)
        self.cell(w_c, h_row, "", border=1, fill=False)

        # Row 3: PUNTOS (50)
        self.set_xy(start_x + w_a, curr_y + h_row * 2)
        self.set_fill_color(*_CLR_HEADER_GRAY)
        self.set_font("Helvetica", "B", 8)
        self.cell(w_b, h_row, _sanitize("PUNTOS (50)"), border=1, fill=True, align="L")
        self.set_xy(start_x + w_a + w_b, curr_y + h_row * 2)
        self.cell(w_c, h_row, _fmt_pts(record.puntos_cobranza), border=1, fill=False, align="R")

        curr_y += cob_h + 3.0

        # -------------------------------------------------------------------
        # 5. TOTAL INDICADORES BLOCK (2 rows)
        # -------------------------------------------------------------------
        tot_h = h_row * 2
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 8)
        # Multi-line title
        self.cell(w_a, tot_h, "", border=1, fill=True)
        self.set_xy(start_x, curr_y + 1.5)
        self.cell(w_a, 4, _sanitize("TOTAL"), align="C")
        self.set_xy(start_x, curr_y + 5.5)
        self.cell(w_a, 4, _sanitize("INDICADORES"), align="C")

        # Row 1: TOTAL PUNTOS
        self.set_xy(start_x + w_a, curr_y)
        self.set_fill_color(*_CLR_HEADER_GRAY)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 8)
        self.cell(w_b, h_row, _sanitize("TOTAL PUNTOS"), border=1, fill=True, align="L")

        self.set_xy(start_x + w_a + w_b, curr_y)
        tot_bg = _CLR_GREEN_ALERT if record.total_indicadores >= 80.0 else _CLR_PINK_ALERT
        self.set_fill_color(*tot_bg)
        self.cell(w_c, h_row, _fmt_pts(record.total_indicadores), border=1, fill=True, align="R")

        # Row 2: POSICION
        self.set_xy(start_x + w_a, curr_y + h_row)
        self.set_fill_color(*_CLR_HEADER_GRAY)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 8)
        self.cell(w_b, h_row, _sanitize("POSICION"), border=1, fill=True, align="L")
        self.set_xy(start_x + w_a + w_b, curr_y + h_row)
        self.cell(w_c, h_row, str(record.lugar_indicadores) if record.lugar_indicadores > 0 else "0", border=1, fill=False, align="R")

        curr_y += tot_h + 4.0

        # ===================================================================
        # BONUSES (INCENTIVO & CONCURSO)
        # ===================================================================
        # Block INCENTIVO (Height = 12 mm = 2 rows)
        inc_h = 12.0
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, inc_h, _sanitize("INCENTIVO"), border=1, fill=True, align="C")

        # Percentage earned (1% if SI, else 0%)
        self.set_xy(start_x + w_a, curr_y)
        self.set_fill_color(*_CLR_WHITE)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 10)
        inc_str = "1%" if str(record.incentivo).upper() in ("SI", "0.01", "1") else "0%"
        self.cell(w_b, inc_h, _sanitize(inc_str), border=1, fill=False, align="C")

        # Right explanatory rule (merged C54:E55 -> w_c + w_de = 92 mm)
        self.set_xy(start_x + w_a + w_b, curr_y)
        self.rect(start_x + w_a + w_b, curr_y, w_c + w_de, inc_h)
        self.set_font("Helvetica", "B", 7)
        self.set_xy(start_x + w_a + w_b + 2.0, curr_y + 1.5)
        self.cell(w_c + w_de - 4.0, 4.0, _sanitize("PARA OBTENEDER EL 1% SE DEBE CUMPLIR MINIMO 90% SIN VENCER"))
        self.set_xy(start_x + w_a + w_b + 2.0, curr_y + 5.5)
        self.cell(w_c + w_de - 4.0, 4.0, _sanitize("Y 90% DE ALCANCE DE VENTAS (AMBOS)"))

        curr_y += inc_h + 3.0

        # Block CONCURSO / RIFA (Height = 12 mm = 2 rows)
        self.set_xy(start_x, curr_y)
        self.set_fill_color(*_CLR_BLACK)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", "B", 9)
        self.cell(w_a, inc_h, _sanitize("CONCURSO"), border=1, fill=True, align="C")

        self.set_xy(start_x + w_a, curr_y)
        self.set_fill_color(*_CLR_WHITE)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", "B", 10)
        concurso_str = "SI" if str(record.rifa).upper() in ("SI", "TRUE") else "NO"
        self.cell(w_b, inc_h, _sanitize(concurso_str), border=1, fill=False, align="C")

        self.set_xy(start_x + w_a + w_b, curr_y)
        self.rect(start_x + w_a + w_b, curr_y, w_c + w_de, inc_h)
        self.set_font("Helvetica", "B", 7)
        self.set_xy(start_x + w_a + w_b + 2.0, curr_y + 1.5)
        self.cell(w_c + w_de - 4.0, 4.0, _sanitize("PARA PARTICIPAR SE DEBE CUMPLIR MINIMO 90% ACTIVACION Y 90%"))
        self.set_xy(start_x + w_a + w_b + 2.0, curr_y + 5.5)
        self.cell(w_c + w_de - 4.0, 4.0, _sanitize("SKU (AMBOS)"))

        # ===================================================================
        # OUTER BORDER (Excel print area border)
        # ===================================================================
        border_top = start_y
        border_bottom = curr_y + inc_h
        border_h = border_bottom - border_top

        self.set_line_width(0.6)
        self.rect(start_x, border_top, total_w, border_h)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def generate_pdf_reports(
    records: List[VendorRecord],
    app_config: AppConfig,
    output_dir: Path,
) -> List[Path]:
    """Generate individual PDF reports identical to the Excel sheet.

    Args:
        records: Computed vendor records.
        app_config: Application configuration.
        output_dir: Output root folder.

    Returns:
        List of generated PDF file paths.
    """
    pdf_dir = Path(output_dir) / _PDF_SUBDIR
    pdf_dir.mkdir(parents=True, exist_ok=True)

    mes_label = app_config.mes_label or "Noviembre 2026"
    saved_paths: List[Path] = []

    for rec in records:
        # Skip vendors without budget
        if rec.presup_ventas == 0 and rec.presup_activacion == 0 and rec.presup_skus == 0:
            continue

        safe_name = _normalize_filename(rec.vendor_name)
        pdf_path = pdf_dir / f"Indicadores_{safe_name}.pdf"

        try:
            pdf = IndividualVentasPDF()
            pdf.generate_page(record=rec, mes_label=mes_label)
            pdf.output(str(pdf_path))
            saved_paths.append(pdf_path)
        except Exception as exc:
            logger.error("Error generating PDF for %s: %s", rec.vendor_name, exc)
            raise ReportError(f"No se pudo generar el PDF para '{rec.vendor_name}': {exc}") from exc

    logger.info("PDF generation complete: %d PDFs saved to %s", len(saved_paths), pdf_dir)
    return saved_paths
