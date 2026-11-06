"""LYNX global configuration: column name constants, thresholds, and defaults.

All column-name constants map to the actual header strings found in the
Odoo-exported Excel files.  Centralising them here means only this file
needs updating if the ERP export format changes.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Facturado / Devoluciones column names
# (shared schema between facturado_cierre_.xlsx and devoluciones.xlsx)
# ---------------------------------------------------------------------------
COL_FAC_COMERCIAL: str = "Comercial"
"""Vendor / salesperson name column in the invoiced-sales export."""

COL_FAC_SUBTOTAL: str = "Subtotal"
"""Line subtotal (before taxes) column in the invoiced-sales export."""

COL_FAC_NIF: str = "NIF"
"""Customer tax-ID column used to count unique activated clients."""

COL_FAC_REF_INTERNA: str = "Referencia interna"
"""Internal product reference column used to count unique SKUs sold."""

# ---------------------------------------------------------------------------
# Pedidos / Por Facturar / Backorder column names
# ---------------------------------------------------------------------------
COL_PED_COMERCIAL: str = "Comercial"
"""Vendor / salesperson name column in the orders export."""

COL_PED_REFERENCIA: str = "Número"
"""Order reference / number column used to count unique orders."""

COL_PED_BASE_IMPONIBLE: str = "Base imponible"
"""Taxable base amount column in the orders export."""

COL_PED_BACKORDER_VALUE: str = "Importe pendiente"
"""Pending / backorder monetary value column."""

# ---------------------------------------------------------------------------
# Cobranza (accounts-receivable) column names
# ---------------------------------------------------------------------------
COL_COB_COMERCIAL: str = "Comercial"
"""Vendor name column in the accounts-receivable export."""

COL_COB_FECHA_ENTREGA: str = "Fecha de vencimiento"
"""Due date column used to classify amounts as current vs. overdue."""

COL_COB_IMPORTE_ADEUDADO: str = "Importe adeudado"
"""Outstanding balance column in the accounts-receivable export."""

# ---------------------------------------------------------------------------
# Zonas column names
# ---------------------------------------------------------------------------
COL_ZONA_NOMBRE: str = "Comercial"
"""Vendor name column in the zones / territories master table."""

COL_ZONA_ZONA: str = "Zona"
"""Territory / zone code or name column."""

COL_ZONA_CORREO: str = "Correo"
"""Vendor e-mail address column."""

COL_ZONA_CODIGO: str = "Código"
"""Vendor internal code column."""

COL_ZONA_REGION: str = "Región"
"""Region name column used to filter and group vendors."""

# ---------------------------------------------------------------------------
# Presupuestos (budgets) column names
# ---------------------------------------------------------------------------
COL_PRES_NOMBRE: str = "Comercial"
"""Vendor name column in the monthly budgets table."""

COL_PRES_ACTIVACION: str = "Presupuesto Activación"
"""Budget target for client activations."""

COL_PRES_SKU: str = "Presupuesto SKU"
"""Budget target for unique SKUs sold."""

COL_PRES_VENTAS: str = "Presupuesto Ventas"
"""Budget target for net sales amount."""

COL_PRES_REGION: str = "Región"
"""Region column in the budgets table (mirrors COL_ZONA_REGION)."""

# ---------------------------------------------------------------------------
# KPI scoring weights
# ---------------------------------------------------------------------------
MAX_PUNTOS_VENTAS: float = 30.0
"""Maximum points awarded for the Ventas (sales) KPI component."""

MAX_PUNTOS_COBRANZA: float = 50.0
"""Maximum points awarded for the Cobranza (collections) KPI component."""

MAX_PUNTOS_ACTIVACION: float = 10.0
"""Maximum points awarded for the Activación (client activation) KPI component."""

MAX_PUNTOS_SKUS: float = 10.0
"""Maximum points awarded for the SKUs (product breadth) KPI component."""

# ---------------------------------------------------------------------------
# Business-rule thresholds
# ---------------------------------------------------------------------------
THRESHOLD_INCENTIVO: float = 0.8945
"""Minimum attainment rate required (for both ventas and cobranza) to earn
the 1 % sales incentive commission."""

THRESHOLD_RIFA: float = 0.90
"""Minimum attainment rate required (for both activación and SKUs) to qualify
for the monthly raffle ("rifa")."""

INCENTIVO_RATE: float = 0.01
"""Commission rate applied when the incentivo threshold is met (1 %)."""

# ---------------------------------------------------------------------------
# Misc
# ---------------------------------------------------------------------------
DEFAULT_REGION: str = "Sin Región"
"""Fallback region label for vendors whose zone record has no region value."""

CUTOFF_HOUR: int = 16
CUTOFF_MINUTE: int = 30
"""Hour and minute at which the working-day reference date rolls forward.
Before this time the previous business day is used as the reference."""

# ---------------------------------------------------------------------------
# 34-column General sheet headers (order matches VendorRecord.to_row())
# ---------------------------------------------------------------------------
GENERAL_HEADERS: list[str] = [
    "Nombre",
    "Zona",
    "Región",
    "Cant. Pedidos",
    "Monto Pedidos",
    "Devoluciones",
    "Total Facturado",
    "Ventas Netas",
    "Monto por Procesar",
    "BackOrder",
    "Presup. Ventas",
    "Diferencia Ventas",
    "Alcance Ventas",
    "Puntos Ventas (30)",
    "Clientes Activados",
    "Presup. Activación",
    "Diferencia Activación",
    "Alcance Activación",
    "Puntos Activación (10)",
    "SKUs Vendidos",
    "Presup. SKUs",
    "Diferencia SKUs",
    "Alcance SKUs",
    "Puntos SKUs (10)",
    "% Sin Vencer",
    "Puntos Cobranza (50)",
    "Total Indicadores",
    "Incentivo",
    "Lugar Ventas",
    "Lugar Activación",
    "Lugar SKUs",
    "Lugar Indicadores",
    "Rifa",
    "Participación",
]
"""Ordered list of column headers for the General sheet (34 columns)."""

# ---------------------------------------------------------------------------
# Regions
# ---------------------------------------------------------------------------
REGIONES: list[str] = [
    "CAPITAL",
    "CENTRO",
    "OCCIDENTE",
    "ANDES",
    "ORIENTE NORTE",
    "ORIENTE SUR",
    "LLANOS",
    "CENTRO COSTA",
]
"""The 8 sales regions — each gets its own regional sheet in the Excel report."""

# ---------------------------------------------------------------------------
# Spanish month names (for GUI month selector)
# ---------------------------------------------------------------------------
MESES_ES: list[str] = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
"""Spanish month names, January-indexed at position 0."""
