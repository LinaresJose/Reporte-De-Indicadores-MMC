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
COL_FAC_COMERCIAL: str = "Líneas de factura/Empresa/Comercial"
"""Vendor / salesperson name column in the invoiced-sales export."""

COL_FAC_SUBTOTAL: str = "Líneas de factura/Subtotal"
"""Line subtotal (before taxes) column in the invoiced-sales export."""

COL_FAC_NIF: str = "Líneas de factura/Empresa/NIF/Identificación fiscal"
"""Customer tax-ID column used to count unique activated clients."""

COL_FAC_REF_INTERNA: str = "Líneas de factura/Producto/Referencia interna"
"""Internal product reference column used to count unique SKUs sold."""

# ---------------------------------------------------------------------------
# Pedidos / Por Facturar / Backorder column names
# ---------------------------------------------------------------------------
COL_PED_COMERCIAL: str = "Comercial"
"""Vendor / salesperson name column in the orders export."""

COL_PED_REFERENCIA: str = "Referencia del pedido"
"""Order reference / number column used to count unique orders."""

COL_PED_BASE_IMPONIBLE: str = "Base imponible"
"""Taxable base amount column in the orders export."""

COL_PED_BACKORDER_VALUE: str = "Backorder Value"
"""Pending / backorder monetary value column."""

# ---------------------------------------------------------------------------
# Cobranza (accounts-receivable) column names and credit limits
# ---------------------------------------------------------------------------
COL_COB_COMERCIAL: str = "Comercial/Nombre"
"""Vendor name column in the accounts-receivable export."""

COL_COB_FECHA_ENTREGA: str = "Fecha de entrega"
"""Delivery date column used to calculate days elapsed since delivery."""

COL_COB_FECHA_VENCIMIENTO: str = "Fecha de vencimiento"
"""Due date column used as secondary/fallback classification."""

COL_COB_CLIENTE: str = "Nombre del socio a mostrar en la factura."
"""Customer name column used to determine VIP status."""

COL_COB_IMPORTE_ADEUDADO: str = "Importe adeudado con signo"
"""Outstanding balance column in the accounts-receivable export."""

CREDIT_DAYS_STANDARD: int = 32
"""Maximum days since delivery for standard accounts to be considered 'Sin Vencer' (30 days credit + 2 grace days)."""

CREDIT_DAYS_VIP: int = 45
"""Maximum days since delivery for VIP accounts to be considered 'Sin Vencer'."""


# ---------------------------------------------------------------------------
# Zonas column names
# ---------------------------------------------------------------------------
COL_ZONA_NOMBRE: str = "Nombre"
"""Vendor name column in the zones / territories master table."""

COL_ZONA_ZONA: str = "Zona"
"""Territory / zone code or name column."""

COL_ZONA_CORREO: str = "Correo"
"""Vendor e-mail address column."""

COL_ZONA_CODIGO: str = "Codigo"
"""Vendor internal code column."""

COL_ZONA_REGION: str = "Región"
"""Region name column used to filter and group vendors."""

# ---------------------------------------------------------------------------
# Presupuestos (budgets) column names
# ---------------------------------------------------------------------------
COL_PRES_NOMBRE: str = "Nombre"
"""Vendor name column in the monthly budgets table."""

COL_PRES_ACTIVACION: str = "ACTIVACIÓN"
"""Budget target for client activations."""

COL_PRES_SKU: str = "SKU"
"""Budget target for unique SKUs sold."""

COL_PRES_VENTAS: str = "VENTAS"
"""Budget target for net sales amount."""

COL_PRES_REGION: str = "Region"
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

# ---------------------------------------------------------------------------
# Vendor name alias map  (Odoo name  →  canonical name in zonas.xlsx)
#
# When Odoo exports a vendor with a different name string than the one stored
# in zonas.xlsx, add a mapping here.  The engine applies this map before
# building its lookup dictionaries so that the data always gets linked to the
# correct vendor record regardless of how Odoo labels it.
#
# Keys are the EXACT strings that appear in the Odoo Excel exports.
# Values are the EXACT strings as they appear in data/zonas.xlsx  (Nombre col).
# Matching is done case-insensitively, but prefer to keep the same casing here
# for readability.
# ---------------------------------------------------------------------------
VENDOR_ALIAS: dict[str, str] = {
    # Odoo name                              →  Canonical name in zonas.xlsx
    "DOUGLAS GONZALEZ(FALCON-DABJ ZULIA)": "DOUGLAS GONZALEZ(DABAJURO)",
    "MARIA NOGUERA(ZULIA-MCBO)":           "MARIA NOGUERA(MARACAIBO)",
}
"""Map from Odoo export vendor names → canonical vendor names in zonas.xlsx.

Add a new entry here whenever an Odoo export uses a name that differs from
the one registered in ``data/zonas.xlsx``.  The engine resolves aliases before
building its lookup dictionaries.
"""

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
