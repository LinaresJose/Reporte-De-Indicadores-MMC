"""
ETL module for the LYNX sales KPI system.

Responsible for reading and transforming the six Odoo source files into
in-memory aggregated structures that the engine layer can consume.

Each public function follows the pattern:
  read_<source>(path) -> Dict[str, <AggModel>]

The dictionary key is always the stripped vendor name as it appears in
the Odoo export, allowing O(1) lookup in the engine.
"""
from __future__ import annotations

import logging
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List

import pandas as pd

import config
from core.exceptions import ETLError, MissingColumnError, EmptySourceError
from core.models import (
    BackorderAgg,
    BudgetRecord,
    CobranzaAgg,
    DevolucionesAgg,
    FacturadoAgg,
    PedidosAgg,
    PorFacturarAgg,
    WorkingDaysInfo,
    ZonaInfo,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _load_workbook(path: str | Path, sheet_index: int = 0) -> pd.DataFrame:
    """Load an Excel workbook and return the specified sheet as a DataFrame.

    Args:
        path: Absolute or relative path to the ``.xlsx`` / ``.xlsm`` file.
        sheet_index: Zero-based index of the sheet to read (default: 0).

    Returns:
        A :class:`pandas.DataFrame` with the sheet contents.

    Raises:
        ETLError: If the file does not exist, cannot be opened, or the sheet
            index is out of range.
    """
    path = Path(path)
    if not path.exists():
        raise ETLError(
            f"Source file not found: '{path}'",
            detail=str(path),
        )
    try:
        df = pd.read_excel(path, sheet_name=sheet_index, engine="openpyxl")
    except Exception as exc:
        raise ETLError(
            f"Cannot read workbook '{path}': {exc}",
            detail=str(path),
        ) from exc

    logger.info("Loaded '%s' — %d rows, %d columns.", path.name, len(df), len(df.columns))
    return df


def _validate_columns(
    df: pd.DataFrame,
    required: list[str],
    source_name: str,
) -> None:
    """Assert that all required columns exist in *df* and that it has data.

    Args:
        df: The DataFrame to validate.
        required: List of column names that must be present.
        source_name: Human-readable label used in error messages (e.g. file name).

    Raises:
        MissingColumnError: If one or more required columns are absent.
            The *column_name* attribute contains all missing names joined by
            ``', '``.
        EmptySourceError: If *df* has zero rows.
    """
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise MissingColumnError(
            column_name=", ".join(missing),
            detail=f"Source: {source_name}",
        )
    if df.empty:
        raise EmptySourceError(
            f"Source file '{source_name}' has no data rows.",
            detail=source_name,
        )


# ---------------------------------------------------------------------------
# Public ETL functions
# ---------------------------------------------------------------------------


def read_facturado(path: str | Path) -> Dict[str, FacturadoAgg]:
    """Read and aggregate the invoiced-sales file (``facturado_cierre_.xlsx``).

    Groups invoice lines by sales representative and computes:
    - **total_facturado**: sum of all subtotals.
    - **clientes_activados**: distinct client NIF/tax-IDs with at least one line.
    - **skus_vendidos**: distinct product internal references sold.

    Args:
        path: Path to ``facturado_cierre_.xlsx`` exported from Odoo.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`FacturadoAgg`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_FAC_COMERCIAL,
        config.COL_FAC_SUBTOTAL,
        config.COL_FAC_NIF,
        config.COL_FAC_REF_INTERNA,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_FAC_COMERCIAL] = df[config.COL_FAC_COMERCIAL].astype(str).str.strip()
    df[config.COL_FAC_SUBTOTAL] = pd.to_numeric(df[config.COL_FAC_SUBTOTAL], errors="coerce").fillna(0.0)

    grouped = df.groupby(config.COL_FAC_COMERCIAL).agg(
        total_facturado=(config.COL_FAC_SUBTOTAL, "sum"),
        clientes_activados=(config.COL_FAC_NIF, "nunique"),
        skus_vendidos=(config.COL_FAC_REF_INTERNA, "nunique"),
    ).reset_index()

    result: Dict[str, FacturadoAgg] = {}
    for _, row in grouped.iterrows():
        name = str(row[config.COL_FAC_COMERCIAL]).strip()
        if name and name.lower() != "nan":
            result[name] = FacturadoAgg(
                vendor_name=name,
                total_facturado=float(row["total_facturado"]),
                clientes_activados=int(row["clientes_activados"]),
                skus_vendidos=int(row["skus_vendidos"]),
            )

    total = sum(r.total_facturado for r in result.values())
    logger.info(
        "read_facturado: %d vendors | total_facturado=$%.2f",
        len(result),
        total,
    )
    return result


def read_pedidos(path: str | Path) -> Dict[str, PedidosAgg]:
    """Read and aggregate the received-orders file (``Pedido_recibido_cierre.xlsx``).

    Groups orders by sales representative and computes:
    - **cantidad_pedidos**: count of distinct order references.
    - **monto_pedidos**: sum of base taxable amounts.

    Args:
        path: Path to the orders export from Odoo.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`PedidosAgg`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_PED_COMERCIAL,
        config.COL_PED_REFERENCIA,
        config.COL_PED_BASE_IMPONIBLE,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_PED_COMERCIAL] = df[config.COL_PED_COMERCIAL].astype(str).str.strip()
    df = df[df[config.COL_PED_COMERCIAL].notna() & (df[config.COL_PED_COMERCIAL] != "") & (df[config.COL_PED_COMERCIAL].str.lower() != "nan")]
    df[config.COL_PED_BASE_IMPONIBLE] = pd.to_numeric(df[config.COL_PED_BASE_IMPONIBLE], errors="coerce").fillna(0.0)

    grouped = df.groupby(config.COL_PED_COMERCIAL).agg(
        cantidad_pedidos=(config.COL_PED_REFERENCIA, "nunique"),
        monto_pedidos=(config.COL_PED_BASE_IMPONIBLE, "sum"),
    ).reset_index()

    result: Dict[str, PedidosAgg] = {}
    for _, row in grouped.iterrows():
        name = str(row[config.COL_PED_COMERCIAL]).strip()
        result[name] = PedidosAgg(
            vendor_name=name,
            cantidad_pedidos=int(row["cantidad_pedidos"]),
            monto_pedidos=float(row["monto_pedidos"]),
        )

    logger.info("read_pedidos: %d vendors", len(result))
    return result


def read_devoluciones(path: str | Path) -> Dict[str, DevolucionesAgg]:
    """Read and aggregate the returns/credit-notes file (``devoluciones_cierre.xlsx``).

    Groups return lines by sales representative and sums subtotals.

    Args:
        path: Path to the credit-notes export from Odoo. The column structure
            matches ``facturado_cierre_.xlsx`` (``Líneas de factura/...``).

    Returns:
        Dictionary keyed by stripped vendor name → :class:`DevolucionesAgg`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [config.COL_FAC_COMERCIAL, config.COL_FAC_SUBTOTAL]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_FAC_COMERCIAL] = df[config.COL_FAC_COMERCIAL].astype(str).str.strip()
    df[config.COL_FAC_SUBTOTAL] = pd.to_numeric(df[config.COL_FAC_SUBTOTAL], errors="coerce").fillna(0.0)

    grouped = df.groupby(config.COL_FAC_COMERCIAL).agg(
        total_devoluciones=(config.COL_FAC_SUBTOTAL, "sum"),
    ).reset_index()

    result: Dict[str, DevolucionesAgg] = {}
    for _, row in grouped.iterrows():
        name = str(row[config.COL_FAC_COMERCIAL]).strip()
        if name and name.lower() != "nan":
            result[name] = DevolucionesAgg(
                vendor_name=name,
                total_devoluciones=float(row["total_devoluciones"]),
            )

    logger.info("read_devoluciones: %d vendors", len(result))
    return result


def read_por_cobrar(path: str | Path) -> Dict[str, CobranzaAgg]:
    """Read and aggregate the accounts-receivable file (``cuentas_por_cobrar_cierre.xlsx``).

    Classifies each invoice as *sin vencer* (not yet due) or *vencido* (past due)
    by comparing ``Fecha de entrega`` against today's date, then groups by vendor.

    Args:
        path: Path to the accounts-receivable export from Odoo.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`CobranzaAgg`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_COB_COMERCIAL,
        config.COL_COB_FECHA_ENTREGA,
        config.COL_COB_IMPORTE_ADEUDADO,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_COB_COMERCIAL] = df[config.COL_COB_COMERCIAL].astype(str).str.strip()
    df[config.COL_COB_IMPORTE_ADEUDADO] = pd.to_numeric(
        df[config.COL_COB_IMPORTE_ADEUDADO], errors="coerce"
    ).fillna(0.0)
    df[config.COL_COB_FECHA_ENTREGA] = pd.to_datetime(
        df[config.COL_COB_FECHA_ENTREGA], errors="coerce"
    )

    today = date.today()
    df["_sin_vencer"] = df.apply(
        lambda r: float(r[config.COL_COB_IMPORTE_ADEUDADO])
        if pd.notna(r[config.COL_COB_FECHA_ENTREGA])
        and r[config.COL_COB_FECHA_ENTREGA].date() >= today
        else 0.0,
        axis=1,
    )
    df["_vencido"] = df.apply(
        lambda r: float(r[config.COL_COB_IMPORTE_ADEUDADO])
        if pd.isna(r[config.COL_COB_FECHA_ENTREGA])
        or r[config.COL_COB_FECHA_ENTREGA].date() < today
        else 0.0,
        axis=1,
    )

    grouped = df.groupby(config.COL_COB_COMERCIAL).agg(
        total_sin_vencer=("_sin_vencer", "sum"),
        total_vencido=("_vencido", "sum"),
    ).reset_index()

    result: Dict[str, CobranzaAgg] = {}
    for _, row in grouped.iterrows():
        name = str(row[config.COL_COB_COMERCIAL]).strip()
        if name and name.lower() != "nan":
            result[name] = CobranzaAgg(
                vendor_name=name,
                total_sin_vencer=float(row["total_sin_vencer"]),
                total_vencido=float(row["total_vencido"]),
            )

    logger.info("read_por_cobrar: %d vendors", len(result))
    return result


def read_por_facturar(path: str | Path) -> Dict[str, PorFacturarAgg]:
    """Read and aggregate the pending-invoicing file (``Pedido_por_facturar_cierre.xlsx``).

    Groups orders pending invoicing by sales representative.

    Args:
        path: Path to the pending-invoicing export from Odoo.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`PorFacturarAgg`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_PED_COMERCIAL,
        config.COL_PED_REFERENCIA,
        config.COL_PED_BASE_IMPONIBLE,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_PED_COMERCIAL] = df[config.COL_PED_COMERCIAL].astype(str).str.strip()
    df = df[df[config.COL_PED_COMERCIAL].notna() & (df[config.COL_PED_COMERCIAL] != "") & (df[config.COL_PED_COMERCIAL].str.lower() != "nan")]
    df[config.COL_PED_BASE_IMPONIBLE] = pd.to_numeric(df[config.COL_PED_BASE_IMPONIBLE], errors="coerce").fillna(0.0)

    grouped = df.groupby(config.COL_PED_COMERCIAL).agg(
        pedidos_unicos=(config.COL_PED_REFERENCIA, "nunique"),
        total_base_imponible=(config.COL_PED_BASE_IMPONIBLE, "sum"),
    ).reset_index()

    result: Dict[str, PorFacturarAgg] = {}
    for _, row in grouped.iterrows():
        name = str(row[config.COL_PED_COMERCIAL]).strip()
        result[name] = PorFacturarAgg(
            vendor_name=name,
            pedidos_unicos=int(row["pedidos_unicos"]),
            total_base_imponible=float(row["total_base_imponible"]),
        )

    logger.info("read_por_facturar: %d vendors", len(result))
    return result


def read_backorder(path: str | Path) -> Dict[str, BackorderAgg]:
    """Read and aggregate the backorder file (``backorder_cierre.xlsx``).

    Groups partially-delivered orders by sales representative and sums
    the pending backorder value.

    Args:
        path: Path to the backorder export from Odoo.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`BackorderAgg`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_PED_COMERCIAL,
        config.COL_PED_REFERENCIA,
        config.COL_PED_BACKORDER_VALUE,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_PED_COMERCIAL] = df[config.COL_PED_COMERCIAL].astype(str).str.strip()
    df = df[df[config.COL_PED_COMERCIAL].notna() & (df[config.COL_PED_COMERCIAL] != "") & (df[config.COL_PED_COMERCIAL].str.lower() != "nan")]
    df[config.COL_PED_BACKORDER_VALUE] = pd.to_numeric(df[config.COL_PED_BACKORDER_VALUE], errors="coerce").fillna(0.0)

    grouped = df.groupby(config.COL_PED_COMERCIAL).agg(
        pedidos_unicos=(config.COL_PED_REFERENCIA, "nunique"),
        backorder_value=(config.COL_PED_BACKORDER_VALUE, "sum"),
    ).reset_index()

    result: Dict[str, BackorderAgg] = {}
    for _, row in grouped.iterrows():
        name = str(row[config.COL_PED_COMERCIAL]).strip()
        result[name] = BackorderAgg(
            vendor_name=name,
            pedidos_unicos=int(row["pedidos_unicos"]),
            backorder_value=float(row["backorder_value"]),
        )

    logger.info("read_backorder: %d vendors", len(result))
    return result


def read_zonas(path: str | Path) -> Dict[str, ZonaInfo]:
    """Read the vendor-to-zone mapping file (``zonas.xlsx``).

    This file is the master catalogue of sales representatives. Only vendors
    listed here will appear in the final report.

    Args:
        path: Path to ``zonas.xlsx``. Expected columns match ``config.COL_ZONA_*``.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`ZonaInfo`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_ZONA_NOMBRE,
        config.COL_ZONA_ZONA,
        config.COL_ZONA_CORREO,
        config.COL_ZONA_CODIGO,
        config.COL_ZONA_REGION,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_ZONA_NOMBRE] = df[config.COL_ZONA_NOMBRE].astype(str).str.strip()
    df = df[df[config.COL_ZONA_NOMBRE].notna() & (df[config.COL_ZONA_NOMBRE] != "") & (df[config.COL_ZONA_NOMBRE].str.lower() != "nan")]

    result: Dict[str, ZonaInfo] = {}
    for _, row in df.iterrows():
        name = str(row[config.COL_ZONA_NOMBRE]).strip()
        result[name] = ZonaInfo(
            vendor_name=name,
            zona=str(row.get(config.COL_ZONA_ZONA, "")).strip(),
            region=str(row.get(config.COL_ZONA_REGION, "")).strip().upper(),
            correo=str(row.get(config.COL_ZONA_CORREO, "")).strip(),
            codigo=str(row.get(config.COL_ZONA_CODIGO, "")).strip(),
        )

    logger.info("read_zonas: %d vendors loaded.", len(result))
    return result


def read_presupuestos(path: str | Path) -> Dict[str, BudgetRecord]:
    """Read the monthly budget targets file (``presupuestos.xlsx``).

    Args:
        path: Path to the budget Excel file provided by the manager.
            Expected columns match ``config.COL_PRES_*``.

    Returns:
        Dictionary keyed by stripped vendor name → :class:`BudgetRecord`.

    Raises:
        ETLError: On file access or format errors.
        MissingColumnError: If a required column is absent.
        EmptySourceError: If the file has no data rows.
    """
    df = _load_workbook(path)
    required = [
        config.COL_PRES_NOMBRE,
        config.COL_PRES_ACTIVACION,
        config.COL_PRES_SKU,
        config.COL_PRES_VENTAS,
        config.COL_PRES_REGION,
    ]
    _validate_columns(df, required, Path(path).name)

    df[config.COL_PRES_NOMBRE] = df[config.COL_PRES_NOMBRE].astype(str).str.strip()
    df = df[df[config.COL_PRES_NOMBRE].notna() & (df[config.COL_PRES_NOMBRE] != "") & (df[config.COL_PRES_NOMBRE].str.lower() != "nan")]

    for col in [config.COL_PRES_ACTIVACION, config.COL_PRES_SKU, config.COL_PRES_VENTAS]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    result: Dict[str, BudgetRecord] = {}
    for _, row in df.iterrows():
        name = str(row[config.COL_PRES_NOMBRE]).strip()
        result[name] = BudgetRecord(
            vendor_name=name,
            presup_activacion=float(row[config.COL_PRES_ACTIVACION]),
            presup_skus=float(row[config.COL_PRES_SKU]),
            presup_ventas=float(row[config.COL_PRES_VENTAS]),
            region=str(row.get(config.COL_PRES_REGION, "")).strip().upper(),
        )

    logger.info("read_presupuestos: %d budget records loaded.", len(result))
    return result


def read_feriados(path: str | Path) -> List[date]:
    """Read the public-holidays file (``feriados.xlsx``).

    This file is non-critical; if it is missing the function logs a warning
    and returns an empty list (no holidays excluded from working-day counts).

    Args:
        path: Path to an Excel file with dates in column A. The first row
            may optionally contain a header such as ``"Fecha"``.

    Returns:
        List of :class:`datetime.date` objects representing public holidays.
    """
    path = Path(path)
    if not path.exists():
        logger.warning("Feriados file not found at '%s'. Proceeding with no holidays.", path)
        return []
    try:
        df = pd.read_excel(path, header=None, engine="openpyxl")
        dates: List[date] = []
        for val in df.iloc[:, 0]:
            if pd.isna(val):
                continue
            if isinstance(val, (datetime,)):
                dates.append(val.date())
            elif isinstance(val, date):
                dates.append(val)
            else:
                try:
                    dates.append(pd.to_datetime(val).date())
                except Exception:
                    pass  # skip non-parseable values
        logger.info("read_feriados: %d holidays loaded from '%s'.", len(dates), path.name)
        return dates
    except Exception as exc:
        logger.warning("Could not read feriados from '%s': %s. Continuing without holidays.", path, exc)
        return []


def calculate_working_days(
    year: Optional[int] = None,
    month: Optional[int] = None,
    feriados: Optional[List[date]] = None,
) -> WorkingDaysInfo:
    """Calculate working-day metrics for a given month.

    Working days are Monday–Friday excluding dates listed in *feriados*.
    If *year* or *month* are omitted, they default to current local date.
    If *feriados* is omitted, it defaults to empty list.
    """
    now = datetime.now()
    if year is None:
        year = now.year
    if month is None:
        month = now.month
    if feriados is None:
        feriados = []
    from calendar import monthrange

    holiday_set = set(feriados)
    _, last_day = monthrange(year, month)
    month_start = date(year, month, 1)
    month_end = date(year, month, last_day)

    # Total working days in the month
    dias_mes = sum(
        1
        for d in (month_start + timedelta(n) for n in range((month_end - month_start).days + 1))
        if d.weekday() < 5 and d not in holiday_set
    )

    # Reference date for elapsed calculation
    now = datetime.now()
    cutoff = now.replace(hour=16, minute=30, second=0, microsecond=0)
    reference = now.date() if now >= cutoff else now.date() - timedelta(days=1)

    # Clamp reference to month boundaries
    reference = min(reference, month_end)
    reference = max(reference, month_start)

    dias_transcurridos = sum(
        1
        for d in (month_start + timedelta(n) for n in range((reference - month_start).days + 1))
        if d.weekday() < 5 and d not in holiday_set
    )

    logger.info(
        "Working days — month: %d | elapsed (ref %s): %d",
        dias_mes,
        reference,
        dias_transcurridos,
    )
    return WorkingDaysInfo(
        dias_habiles_mes=dias_mes,
        dias_habiles_transcurridos=dias_transcurridos,
    )
