"""reports/pdf_writer.py — LYNX Report System.

Generates individual PDF performance reports for each vendor, replicating
the visual layout of the "Individual Ventas" sheet in the original
``Reporte_de_indicadores_de_ventas_*.xlsm``.

Each PDF is a single A4 page containing:

- **Header banner:** dark-blue with LYNX logo text, month, and vendor name.
- **Identity block:** Zona, Región, Comisionista.
- **KPI sections (4):** VENTAS, ACTIVACIÓN, SKUs, COBRANZA — each showing
  the budget target, achieved value, alcance %, points earned, and a
  colour-coded progress bar.
- **Total Indicadores box:** summed points + overall ranking position.
- **Bonos section:** INCENTIVO (1% extra commission flag) and RIFA eligibility.
- **Footer:** generation timestamp and LYNX version.

Typical usage::

    from reports.pdf_writer import generate_pdf_reports

    pdf_paths = generate_pdf_reports(records, app_config, Path("output"))
    print(f"{len(pdf_paths)} PDFs generated.")
"""

from __future__ import annotations

import logging
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import List

from fpdf import FPDF

import config
from core.models import AppConfig, VendorRecord

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_PDF_SUBDIR = "PDF_INDICADORES"

# Colour palette (R, G, B)
_CLR_NAVY     = (31,  56, 100)   # #1F3864 — header/section backgrounds
_CLR_WHITE    = (255, 255, 255)
_CLR_BLACK    = (0,   0,   0)
_CLR_GREY_BG  = (242, 242, 242)  # alternating row bg
_CLR_GREEN    = (39,  98,  33)   # #276221 — good alcance text
_CLR_GREEN_BG = (198, 239, 206)  # #C6EFCE
_CLR_YELLOW   = (156, 87,   0)   # #9C5700
_CLR_YELLOW_BG= (255, 235, 156)  # #FFEB9C
_CLR_RED      = (156,  0,   6)   # #9C0006
_CLR_RED_BG   = (255, 199, 206)  # #FFC7CE
_CLR_BAR_BG   = (220, 230, 241)  # progress bar track
_CLR_SECTION  = (68, 114, 196)   # #4472C4 — section accent blue

# Page margins (mm)
_MARGIN_LEFT  = 15
_MARGIN_RIGHT = 15
_MARGIN_TOP   = 10
_PAGE_W       = 210  # A4 width mm
_CONTENT_W    = _PAGE_W - _MARGIN_LEFT - _MARGIN_RIGHT  # 180 mm


def _normalize_filename(text: str) -> str:
    """Convert *text* to a filesystem-safe ASCII string for use in file names.

    Args:
        text: Arbitrary string possibly containing accents and special chars.

    Returns:
        Clean ASCII string safe for Windows file names.
    """
    nfd = unicodedata.normalize("NFD", text)
    ascii_text = nfd.encode("ascii", "ignore").decode("ascii")
    safe = re.sub(r'[\\/:*?"<>|]', "", ascii_text)
    safe = re.sub(r"[\s\-]+", "_", safe.strip())
    safe = re.sub(r"[^\w]", "", safe)
    return safe or "SIN_NOMBRE"


def _alcance_colors(alcance: float):
    """Return ``(bg_rgb, fg_rgb)`` colour tuples for a given alcance ratio.

    Args:
        alcance: Achievement ratio (1.0 = 100 %).

    Returns:
        Tuple of ``((R,G,B), (R,G,B))`` for background and foreground.
    """
    if alcance >= 1.0:
        return _CLR_GREEN_BG, _CLR_GREEN
    if alcance >= 0.80:
        return _CLR_YELLOW_BG, _CLR_YELLOW
    return _CLR_RED_BG, _CLR_RED


# ---------------------------------------------------------------------------
# FPDF subclass
# ---------------------------------------------------------------------------


class LynxPDF(FPDF):
    """Custom FPDF subclass that adds LYNX-branded header and footer.

    Args:
        mes_label: Month label string shown in the header (e.g. "Noviembre 2026").
        vendor_nombre: Vendor name shown in the header subtitle.
    """

    def __init__(self, mes_label: str, vendor_nombre: str) -> None:
        super().__init__(orientation="P", unit="mm", format="A4")
        self.mes_label     = mes_label
        self.vendor_nombre = vendor_nombre
        self.set_margins(_MARGIN_LEFT, _MARGIN_TOP, _MARGIN_RIGHT)
        self.set_auto_page_break(auto=False)

    # ------------------------------------------------------------------
    # Header & footer (called automatically by FPDF on add_page / close)
    # ------------------------------------------------------------------

    def header(self) -> None:  # type: ignore[override]
        """Render the dark-blue header banner with LYNX title and vendor name."""
        # Banner background
        self.set_fill_color(*_CLR_NAVY)
        self.rect(0, 0, _PAGE_W, 28, style="F")

        # LYNX title
        self.set_xy(_MARGIN_LEFT, 4)
        self.set_font("Helvetica", style="B", size=18)
        self.set_text_color(*_CLR_WHITE)
        self.cell(100, 10, "LYNX", ln=0)

        # Subtitle: Reporte de Indicadores
        self.set_xy(_MARGIN_LEFT, 14)
        self.set_font("Helvetica", style="", size=9)
        self.cell(100, 6, "Sistema de Indicadores de Ventas", ln=0)

        # Month label (right-aligned)
        self.set_xy(_PAGE_W - _MARGIN_RIGHT - 80, 4)
        self.set_font("Helvetica", style="B", size=11)
        self.cell(80, 10, self.mes_label.upper(), ln=0, align="R")

        # Vendor name (right-aligned)
        self.set_xy(_PAGE_W - _MARGIN_RIGHT - 80, 15)
        self.set_font("Helvetica", style="", size=9)
        self.cell(80, 6, self.vendor_nombre, ln=0, align="R")

        # Thin white separator line
        self.set_draw_color(*_CLR_WHITE)
        self.line(_MARGIN_LEFT, 28, _PAGE_W - _MARGIN_RIGHT, 28)

        self.set_text_color(*_CLR_BLACK)
        self.ln(22)  # advance past the banner

    def footer(self) -> None:  # type: ignore[override]
        """Render the page footer with generation timestamp."""
        self.set_y(-12)
        self.set_font("Helvetica", style="I", size=7)
        self.set_text_color(128, 128, 128)
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.cell(0, 5, f"Generado: {timestamp}  |  LYNX v1.0", align="C")
        self.set_text_color(*_CLR_BLACK)

    # ------------------------------------------------------------------
    # Building-block helpers
    # ------------------------------------------------------------------

    def _section_title(self, title: str) -> None:
        """Render a section title bar (accent blue background, white text).

        Args:
            title: Section title text in uppercase.
        """
        self.set_fill_color(*_CLR_SECTION)
        self.set_text_color(*_CLR_WHITE)
        self.set_font("Helvetica", style="B", size=10)
        self.cell(_CONTENT_W, 7, f"  {title}", ln=True, fill=True)
        self.set_text_color(*_CLR_BLACK)
        self.ln(1)

    def _label_value_row(
        self,
        label: str,
        value: str,
        bg: tuple = _CLR_WHITE,
        fg: tuple = _CLR_BLACK,
        bold_value: bool = False,
        h: float = 7,
    ) -> None:
        """Render a two-column label-value row.

        Args:
            label: Left-side descriptive text.
            value: Right-side value text.
            bg: Background RGB tuple.
            fg: Foreground (text) RGB tuple for value.
            bold_value: If ``True``, render value in bold.
            h: Row height in mm.
        """
        col_w = _CONTENT_W / 2

        # Label cell (left)
        self.set_fill_color(*_CLR_GREY_BG)
        self.set_text_color(*_CLR_BLACK)
        self.set_font("Helvetica", style="", size=9)
        self.cell(col_w, h, f"  {label}", border=1, fill=True)

        # Value cell (right)
        self.set_fill_color(*bg)
        self.set_text_color(*fg)
        self.set_font("Helvetica", style="B" if bold_value else "", size=9)
        self.cell(col_w, h, value, border=1, fill=True, align="C", ln=True)

        self.set_text_color(*_CLR_BLACK)

    def _progress_bar(self, alcance: float, width: float = _CONTENT_W) -> None:
        """Render a horizontal progress bar for an alcance value.

        The bar is filled to ``min(alcance, 1.0)`` of its width, capped at
        100 % visually.  Colour follows the traffic-light convention.

        Args:
            alcance: Achievement ratio (0.0 – any positive float).
            width: Total bar width in mm.
        """
        h       = 4
        fill_w  = min(alcance, 1.0) * width
        _, fg   = _alcance_colors(alcance)
        bar_fg  = fg  # filled portion colour

        # Track (background)
        self.set_fill_color(*_CLR_BAR_BG)
        self.rect(self.get_x(), self.get_y(), width, h, style="F")

        # Filled portion
        if fill_w > 0:
            self.set_fill_color(*bar_fg)
            self.rect(self.get_x(), self.get_y(), fill_w, h, style="F")

        # Thin border around the full bar
        self.set_draw_color(180, 180, 180)
        self.rect(self.get_x(), self.get_y(), width, h)

        self.ln(h + 1)

    def _kpi_section(
        self,
        title: str,
        presupuesto: float,
        logrado: float,
        alcance: float,
        puntos: float,
        max_puntos: int,
        fmt_currency: bool = True,
    ) -> None:
        """Render a complete KPI section (title + 4 rows + progress bar).

        Args:
            title: Section header (e.g. ``"VENTAS"``).
            presupuesto: Budget / target value.
            logrado: Actually achieved value.
            alcance: Achievement ratio (logrado / presupuesto).
            puntos: Points earned.
            max_puntos: Maximum possible points for this KPI.
            fmt_currency: If ``True``, format numbers as currency; otherwise integer.
        """
        self._section_title(title)

        def _fmt(v: float) -> str:
            if fmt_currency:
                return f"Bs. {v:,.2f}"
            return f"{int(v):,}"

        pct_str = f"{alcance:.2%}"
        pts_str = f"{puntos:.2f} / {max_puntos}"

        bg, fg = _alcance_colors(alcance)

        self._label_value_row("Presupuesto", _fmt(presupuesto))
        self._label_value_row("Logrado",     _fmt(logrado))
        self._label_value_row("Alcance",     pct_str,  bg=bg, fg=fg, bold_value=True)
        self._label_value_row("Puntos",      pts_str,  bg=bg, fg=fg, bold_value=True)

        self._progress_bar(alcance)
        self.ln(3)

    # ------------------------------------------------------------------
    # Page composer
    # ------------------------------------------------------------------

    def add_vendor_page(self, record: VendorRecord) -> None:
        """Add a full A4 page for a single vendor's performance report.

        Args:
            record: Fully computed :class:`~core.models.VendorRecord`.
        """
        self.add_page()

        # ---- Identity block ----
        self._section_title("DATOS DEL COMISIONISTA")
        self._label_value_row("Nombre / Comisionista", record.nombre.upper())
        self._label_value_row("Zona",   record.zona)
        self._label_value_row("Región", record.region)
        self.ln(4)

        # ---- KPI: VENTAS (30 pts) ----
        self._kpi_section(
            title        = "VENTAS  (30 puntos)",
            presupuesto  = record.presup_ventas,
            logrado      = record.ventas_netas,
            alcance      = record.alcance_ventas,
            puntos       = record.puntos_ventas,
            max_puntos   = 30,
            fmt_currency = True,
        )

        # ---- KPI: COBRANZA (50 pts) ----
        self._kpi_section(
            title        = "COBRANZA  (50 puntos)",
            presupuesto  = 1.0,  # target is 100 % sin vencer
            logrado      = record.pct_sin_vencer,
            alcance      = record.pct_sin_vencer,
            puntos       = record.puntos_cobranza,
            max_puntos   = 50,
            fmt_currency = False,  # shown as percentage below
        )
        # Override the last logrado/alcance display since it's already a ratio
        # (The _kpi_section renders a progress bar based on alcance — correct)

        # ---- KPI: ACTIVACIÓN (10 pts) ----
        self._kpi_section(
            title        = "ACTIVACIÓN DE CLIENTES  (10 puntos)",
            presupuesto  = record.presup_activacion,
            logrado      = record.clientes_activados,
            alcance      = record.alcance_activacion,
            puntos       = record.puntos_activacion,
            max_puntos   = 10,
            fmt_currency = False,
        )

        # ---- KPI: SKUs (10 pts) ----
        self._kpi_section(
            title        = "SKUs VENDIDOS  (10 puntos)",
            presupuesto  = record.presup_skus,
            logrado      = record.skus_vendidos,
            alcance      = record.alcance_skus,
            puntos       = record.puntos_skus,
            max_puntos   = 10,
            fmt_currency = False,
        )

        # ---- Total Indicadores ----
        self._section_title("TOTAL INDICADORES")
        total_bg, total_fg = _alcance_colors(
            record.total_indicadores / 100.0
        )
        self._label_value_row(
            "Total de Puntos",
            f"{record.total_indicadores:.2f} / 100",
            bg=total_bg, fg=total_fg, bold_value=True,
        )
        self._label_value_row(
            "Posición en el ranking",
            f"#{record.lugar_indicadores}" if record.lugar_indicadores > 0 else "—",
        )
        self.ln(4)

        # ---- Bonos ----
        self._section_title("BONOS")

        incentivo_text = "✓ APLICA (1 % comisión extra)" if record.incentivo == "SI" else "✗ No aplica"
        inc_bg = _CLR_GREEN_BG if record.incentivo == "SI" else _CLR_RED_BG
        inc_fg = _CLR_GREEN    if record.incentivo == "SI" else _CLR_RED
        self._label_value_row(
            "Incentivo (Alcance Ventas ≥ 89.45 % y Cobranza ≥ 89.45 %)",
            incentivo_text,
            bg=inc_bg, fg=inc_fg, bold_value=True,
        )

        rifa_text = "✓ PARTICIPA" if record.rifa == "SI" else "✗ No participa"
        rifa_bg = _CLR_GREEN_BG if record.rifa == "SI" else _CLR_RED_BG
        rifa_fg = _CLR_GREEN    if record.rifa == "SI" else _CLR_RED
        self._label_value_row(
            "Rifa (Activación ≥ 90 % y SKUs ≥ 90 %)",
            rifa_text,
            bg=rifa_bg, fg=rifa_fg, bold_value=True,
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def generate_pdf_reports(
    records: List[VendorRecord],
    app_config: AppConfig,
    output_dir: Path,
) -> List[Path]:
    """Generate one PDF performance report per vendor.

    Only vendors with non-zero budget (``presup_ventas > 0`` OR
    ``presup_activacion > 0``) receive a PDF — vendors without budget targets
    are skipped (they cannot have meaningful alcance metrics).

    PDFs are saved to ``output_dir / PDF_INDICADORES /``.

    Args:
        records: Fully computed and ranked vendor records.
        app_config: Application configuration (used for ``mes_label``).
        output_dir: Root output directory.

    Returns:
        List of :class:`~pathlib.Path` objects pointing to the generated PDFs,
        in the same order as the filtered *records*.

    Raises:
        :class:`~core.exceptions.ReportError`: If a PDF file cannot be saved.

    Example::

        paths = generate_pdf_reports(records, app_config, Path("output"))
        print(f"Generated {len(paths)} PDFs")
    """
    from core.exceptions import ReportError

    pdf_dir = Path(output_dir) / _PDF_SUBDIR
    pdf_dir.mkdir(parents=True, exist_ok=True)

    mes_label = app_config.mes_label
    saved_paths: List[Path] = []
    skipped = 0

    for rec in records:
        # Skip vendors without budget (no meaningful report possible)
        if rec.presup_ventas == 0 and rec.presup_activacion == 0 and rec.presup_skus == 0:
            logger.debug("Skipping PDF for '%s' — no budget.", rec.nombre)
            skipped += 1
            continue

        safe_name = _normalize_filename(rec.nombre)
        pdf_path  = pdf_dir / f"Indicadores_{safe_name}.pdf"

        try:
            pdf = LynxPDF(mes_label=mes_label, vendor_nombre=rec.nombre)
            pdf.add_vendor_page(rec)
            pdf.output(str(pdf_path))
            saved_paths.append(pdf_path)
            logger.debug("PDF saved: %s", pdf_path.name)

        except Exception as exc:
            raise ReportError(
                f"No se pudo generar el PDF para '{rec.nombre}': {exc}",
                report_type="pdf_individual",
            ) from exc

    logger.info(
        "PDF generation complete: %d PDFs saved, %d vendors skipped (no budget).",
        len(saved_paths),
        skipped,
    )
    return saved_paths
