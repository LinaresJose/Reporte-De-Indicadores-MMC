"""
Data models for the LYNX sales KPI system.

All monetary values are in USD. Percentages are expressed as decimals (0.0–1.0).
Rankings are 1-based integers; 0 means not ranked (vendor has no budget).
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import Optional

from core.exceptions import ETLError


# ---------------------------------------------------------------------------
# Reference / configuration models
# ---------------------------------------------------------------------------


@dataclass
class ZonaInfo:
    """Mapping of a vendor to their geographic zone and contact details.

    Attributes:
        vendor_name: Full name of the sales representative as it appears
            in Odoo source files.
        zona: Zone identifier (e.g. "Z1", "Z2").
        region: Broad regional grouping from :data:`config.REGIONES`.
        correo: Primary e-mail address of the vendor.
        codigo: Internal vendor code used for cross-referencing.
    """

    vendor_name: str
    zona: str
    region: str
    correo: str
    codigo: str


@dataclass
class BudgetRecord:
    """Monthly budget targets for a single vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        presup_activacion: Target number of activated clients (float for
            consistency with weighted partial budgets).
        presup_skus: Target number of distinct SKUs sold.
        presup_ventas: Target net sales amount in USD.
        region: Regional grouping; used for regional sub-report aggregation.
    """

    vendor_name: str
    presup_activacion: float
    presup_skus: float
    presup_ventas: float
    region: str


# ---------------------------------------------------------------------------
# Aggregation models  (outputs of the ETL phase, one per vendor)
# ---------------------------------------------------------------------------


@dataclass
class FacturadoAgg:
    """Aggregated invoiced-sales metrics for one vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        total_facturado: Sum of all invoice line subtotals in USD.
        clientes_activados: Count of distinct client NIF/tax-IDs with at
            least one invoice line during the period.
        skus_vendidos: Count of distinct product internal references sold.
    """

    vendor_name: str
    total_facturado: float = 0.0
    clientes_activados: int = 0
    skus_vendidos: int = 0


@dataclass
class PedidosAgg:
    """Aggregated received-orders metrics for one vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        cantidad_pedidos: Total number of distinct order references.
        monto_pedidos: Sum of base-imponible (pre-tax) amounts in USD.
    """

    vendor_name: str
    cantidad_pedidos: int = 0
    monto_pedidos: float = 0.0


@dataclass
class DevolucionesAgg:
    """Aggregated returns / credit notes for one vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        total_devoluciones: Sum of return subtotals in USD (positive value).
    """

    vendor_name: str
    total_devoluciones: float = 0.0


@dataclass
class CobranzaAgg:
    """Aggregated accounts-receivable aging for one vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        total_sin_vencer: Total outstanding amount not yet past due (USD).
        total_vencido: Total outstanding amount past due (USD).
    """

    vendor_name: str
    total_sin_vencer: float = 0.0
    total_vencido: float = 0.0

    @property
    def pct_sin_vencer(self) -> float:
        """Fraction of total receivables that are not yet past due.

        Returns:
            ``sin_vencer / (sin_vencer + vencido)`` as a decimal in
            [0.0, 1.0].  Returns ``0.0`` when the denominator is zero to
            avoid :class:`ZeroDivisionError`.
        """
        denominator: float = self.total_sin_vencer + self.total_vencido
        if denominator == 0.0:
            return 0.0
        return self.total_sin_vencer / denominator


@dataclass
class PorFacturarAgg:
    """Aggregated orders pending invoicing (por facturar) for one vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        pedidos_unicos: Count of distinct order references awaiting invoicing.
        total_base_imponible: Sum of pre-tax base amounts yet to be invoiced
            (USD).
    """

    vendor_name: str
    pedidos_unicos: int = 0
    total_base_imponible: float = 0.0


@dataclass
class BackorderAgg:
    """Aggregated backorder metrics for one vendor.

    Attributes:
        vendor_name: Full name of the sales representative.
        pedidos_unicos: Count of distinct order references with backordered
            items.
        backorder_value: Total monetary value of backordered items (USD).
    """

    vendor_name: str
    pedidos_unicos: int = 0
    backorder_value: float = 0.0


# ---------------------------------------------------------------------------
# Calendar helper
# ---------------------------------------------------------------------------


@dataclass
class WorkingDaysInfo:
    """Working-day calendar data for the current reporting period.

    Attributes:
        dias_habiles_mes: Total working days in the reporting month
            (after excluding weekends and public holidays).
        dias_habiles_transcurridos: Working days elapsed since the first day
            of the month up to and including the closing date.
    """

    dias_habiles_mes: int
    dias_habiles_transcurridos: int

    @property
    def total_dias_habiles(self) -> int:
        """Alias for dias_habiles_mes."""
        return self.dias_habiles_mes

    @property
    def dias_habiles_restantes(self) -> int:
        """Working days remaining in the month."""
        return max(0, self.dias_habiles_mes - self.dias_habiles_transcurridos)

    @property
    def pct_transcurrido(self) -> float:
        """Fraction of the month's working days that have elapsed.

        Returns:
            ``dias_habiles_transcurridos / dias_habiles_mes`` as a decimal
            in [0.0, 1.0].  Returns ``0.0`` when ``dias_habiles_mes`` is
            zero to avoid :class:`ZeroDivisionError`.
        """
        if self.dias_habiles_mes == 0:
            return 0.0
        return self.dias_habiles_transcurridos / self.dias_habiles_mes

    @property
    def pct_avance(self) -> float:
        """Alias for pct_transcurrido."""
        return self.pct_transcurrido


# ---------------------------------------------------------------------------
# Primary output model
# ---------------------------------------------------------------------------


@dataclass
class VendorRecord:
    """Complete KPI record for a single vendor — one row in the General sheet.

    This dataclass captures all 34 output columns produced by the LYNX
    reporting engine for each sales representative.  Monetary values are
    in USD; percentages are decimals in [0.0, 1.0]; rankings are 1-based
    integers (0 = not ranked because no budget was assigned).

    Attributes:
        vendor_name: Full name of the sales representative.
        zona: Zone identifier from the Zonas configuration file.
        region: Regional grouping from :data:`config.REGIONES`.

        cantidad_pedidos: Total number of received orders (col 4).
        monto_pedidos: Pre-tax monetary sum of received orders in USD (col 5).

        devoluciones: Total value of returns / credit notes in USD (col 6).
        total_facturado: Gross invoiced amount in USD (col 7).
        ventas_netas: Net sales = ``total_facturado - devoluciones`` (col 8).

        monto_por_procesar: Value of orders pending invoicing in USD (col 9).
        backorder: Value of backordered items in USD (col 10).

        presup_ventas: Monthly sales budget target in USD (col 11).
        diferencia_ventas: ``|ventas_netas - presup_ventas|`` (col 12).
        alcance_ventas: ``ventas_netas / presup_ventas`` as a decimal (col 13).
        puntos_ventas: ``MIN(30, 30 × alcance_ventas)`` (col 14).

        clientes_activados: Distinct clients with at least one invoice (col 15).
        presup_activacion: Target number of activated clients (col 16).
        diferencia_activacion: ``|clientes_activados - presup_activacion|``
            (col 17).
        alcance_activacion: ``clientes_activados / presup_activacion`` (col 18).
        puntos_activacion: ``MIN(10, 10 × alcance_activacion)`` (col 19).

        skus_vendidos: Distinct SKUs sold during the period (col 20).
        presup_skus: Target SKU count (col 21).
        diferencia_skus: ``|skus_vendidos - presup_skus|`` (col 22).
        alcance_skus: ``skus_vendidos / presup_skus`` as a decimal (col 23).
        puntos_skus: ``MIN(10, 10 × alcance_skus)`` (col 24).

        pct_sin_vencer: Fraction of receivables not yet past due (col 25).
        puntos_cobranza: ``50 × pct_sin_vencer`` (col 26).

        total_indicadores: Sum of all *puntos* fields (col 27).
        incentivo: ``0.01`` when ``alcance_ventas >= 89.45 %`` **and**
            ``pct_sin_vencer >= 89.45 %``; otherwise ``0.0`` (col 28).

        lugar_ventas: Rank by net sales among budgeted vendors (col 29).
        lugar_activacion: Rank by activated clients (col 30).
        lugar_skus: Rank by SKUs sold (col 31).
        lugar_indicadores: Rank by total_indicadores (col 32).

        rifa: ``"SI"`` when ``alcance_activacion >= 90 %`` **and**
            ``alcance_skus >= 90 %``; otherwise ``"NO"`` (col 33).
        participacion: ``ventas_netas / total_ventas_global`` (col 34).
    """

    # ── Identity ────────────────────────────────────────────────────────────
    vendor_name: str
    zona: str
    region: str

    # ── Pedidos (cols 4-5) ───────────────────────────────────────────────────
    cantidad_pedidos: int = 0
    monto_pedidos: float = 0.0

    # ── Devoluciones y Facturado (cols 6-8) ─────────────────────────────────
    devoluciones: float = 0.0
    total_facturado: float = 0.0
    ventas_netas: float = 0.0  # = facturado - devoluciones

    # ── Logística (cols 9-10) ────────────────────────────────────────────────
    monto_por_procesar: float = 0.0  # from PorFacturar
    backorder: float = 0.0

    # ── Ventas KPI (cols 11-14) ──────────────────────────────────────────────
    presup_ventas: float = 0.0
    diferencia_ventas: float = 0.0   # ABS(ventas_netas - presup_ventas)
    alcance_ventas: float = 0.0      # ventas_netas / presup_ventas
    puntos_ventas: float = 0.0       # MIN(30, 30 * alcance_ventas)

    # ── Activación KPI (cols 15-19) ──────────────────────────────────────────
    clientes_activados: int = 0
    presup_activacion: float = 0.0
    diferencia_activacion: float = 0.0
    alcance_activacion: float = 0.0
    puntos_activacion: float = 0.0   # MIN(10, 10 * alcance_activacion)

    # ── SKUs KPI (cols 20-24) ────────────────────────────────────────────────
    skus_vendidos: int = 0
    presup_skus: float = 0.0
    diferencia_skus: float = 0.0
    alcance_skus: float = 0.0
    puntos_skus: float = 0.0         # MIN(10, 10 * alcance_skus)

    # ── Cobranza KPI (cols 25-26) ────────────────────────────────────────────
    pct_sin_vencer: float = 0.0
    puntos_cobranza: float = 0.0     # 50 * pct_sin_vencer

    # ── Totals and bonuses (cols 27-28) ─────────────────────────────────────
    total_indicadores: float = 0.0   # sum of all puntos
    incentivo: float = 0.0           # 0.01 if ventas>=89.45% AND cobranza>=89.45%, else 0

    # ── Rankings (cols 29-32) — 0 = not ranked ──────────────────────────────
    lugar_ventas: int = 0
    lugar_activacion: int = 0
    lugar_skus: int = 0
    lugar_indicadores: int = 0

    # ── Bonus (col 33) ───────────────────────────────────────────────────────
    rifa: str = "NO"  # "SI" if activacion>=90% AND skus>=90%, else "NO"

    # ── Participation (col 34) ───────────────────────────────────────────────
    participacion: float = 0.0  # ventas_netas / total_ventas_global

    def to_row(self) -> list:
        """Return all 34 column values in the order defined by GENERAL_HEADERS.

        The returned list is intended to be written directly as a row to the
        *General* worksheet via ``openpyxl``.  The order exactly matches
        :data:`config.GENERAL_HEADERS`.

        Returns:
            A list of 34 values covering identity, operational metrics, KPIs,
            rankings, and participation — ready for spreadsheet insertion.
        """
        return [
            # Identity
            self.vendor_name,       # col 1  – Nombre
            self.zona,              # col 2  – Zona
            self.region,            # col 3  – Región
            # Pedidos
            self.cantidad_pedidos,  # col 4  – Cantidad Pedidos
            self.monto_pedidos,     # col 5  – Monto Pedidos
            # Devoluciones y Facturado
            self.devoluciones,      # col 6  – Devoluciones
            self.total_facturado,   # col 7  – Total Facturado
            self.ventas_netas,      # col 8  – Ventas Netas
            # Logística
            self.monto_por_procesar,  # col 9  – Monto pedidos por Procesar
            self.backorder,           # col 10 – BackOrder
            # Ventas KPI
            self.presup_ventas,       # col 11 – Presup. Ventas
            self.diferencia_ventas,   # col 12 – Diferencia de Ventas
            self.alcance_ventas,      # col 13 – Alcance de Ventas
            self.puntos_ventas,       # col 14 – Puntos Ventas (30)
            # Activación KPI
            self.clientes_activados,  # col 15 – Clientes Activados
            self.presup_activacion,   # col 16 – Presup. Activación
            self.diferencia_activacion,  # col 17 – Diferencia Activación
            self.alcance_activacion,  # col 18 – Alcance Activación
            self.puntos_activacion,   # col 19 – Puntos Activación
            # SKUs KPI
            self.skus_vendidos,       # col 20 – SKUs Vendidos
            self.presup_skus,         # col 21 – Presup. SKUs
            self.diferencia_skus,     # col 22 – Diferencia SKUs
            self.alcance_skus,        # col 23 – Alcance SKUs
            self.puntos_skus,         # col 24 – Puntos SKUs
            # Cobranza KPI
            self.pct_sin_vencer,      # col 25 – % Sin Vencer
            self.puntos_cobranza,     # col 26 – Puntos Cobranza
            # Totals and bonuses
            self.total_indicadores,   # col 27 – Total Reporte Indicadores
            self.incentivo,           # col 28 – Incentivo
            # Rankings
            self.lugar_ventas,        # col 29 – Lugar Ventas
            self.lugar_activacion,    # col 30 – Lugar Activación
            self.lugar_skus,          # col 31 – Lugar SKUs
            self.lugar_indicadores,   # col 32 – Lugar Indicadores
            # Bonus
            self.rifa,                # col 33 – Rifa
            # Participation
            self.participacion,       # col 34 – Participación
        ]


# ---------------------------------------------------------------------------
# Configuration / runtime models
# ---------------------------------------------------------------------------


@dataclass
class SourceFiles:
    """Paths to all source files required by the LYNX ETL pipeline.

    Each attribute corresponds to one Odoo export file.  ``None`` indicates
    that the file has not yet been selected by the user.

    Attributes:
        facturado: Path to ``facturado_cierre_.xlsx``.
        pedidos: Path to ``Pedido_recibido_cierre.xlsx``.
        devoluciones: Path to ``devoluciones_cierre.xlsx``.
        por_cobrar: Path to ``cuentas_por_cobrar_cierre.xlsx``.
        por_facturar: Path to ``Pedido_por_facturar_cierre.xlsx``.
        backorder: Path to ``backorder_cierre.xlsx``.
        zonas: Path to the Zonas configuration workbook.
        presupuestos: Path to the monthly budgets workbook.
        feriados: Path to the public-holidays calendar file.
        cobranza_cierre: Path to the ``INDICADOR_COBRANZA RESUMEN`` workbook
            (alternative summary-level cobranza source).
    """

    facturado: Optional[str] = None
    pedidos: Optional[str] = None
    devoluciones: Optional[str] = None
    por_cobrar: Optional[str] = None
    por_facturar: Optional[str] = None
    backorder: Optional[str] = None
    zonas: Optional[str] = None
    presupuestos: Optional[str] = None
    feriados: Optional[str] = None
    cobranza_cierre: Optional[str] = None

    def validate(self) -> None:
        """Assert that all mandatory source paths have been provided.

        The four files below are required for the core pipeline to run.
        The remaining files (``por_facturar``, ``backorder``, etc.) are
        optional and may be omitted without raising an error.

        Raises:
            ETLError: If any of ``facturado``, ``pedidos``, ``devoluciones``,
                or ``por_cobrar`` is ``None``.
        """
        required: dict[str, Optional[str]] = {
            "facturado": self.facturado,
            "pedidos": self.pedidos,
            "devoluciones": self.devoluciones,
            "por_cobrar": self.por_cobrar,
        }
        missing: list[str] = [
            name for name, path in required.items() if path is None
        ]
        if missing:
            raise ETLError(
                f"Missing required source file(s): {', '.join(missing)}",
                detail=(
                    "All four mandatory files must be provided before "
                    "running the ETL pipeline."
                ),
            )


@dataclass
class AppConfig:
    """Runtime configuration for a single LYNX report run."""

    output_dir: str = "output"
    mes_label: str = ""
    anio: int = field(default_factory=lambda: datetime.date.today().year)
    regions: list[str] = field(default_factory=lambda: [
        "CAPITAL", "CENTRO", "OCCIDENTE", "ANDES",
        "ORIENTE NORTE", "ORIENTE SUR", "LLANOS", "CENTRO COSTA"
    ])
    threshold_incentivo: float = 0.8945
    threshold_rifa: float = 0.90
