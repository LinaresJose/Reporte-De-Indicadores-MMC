"""maintenance/budget_loader.py — LYNX Report System.

Handles the monthly Presupuestos (budget) maintenance workflow.

The sales manager delivers a new Excel file each month with updated budget
targets per vendor (VENTAS, ACTIVACIÓN, SKUs).  This module loads that file
and persists the data to ``data/presupuestos.xlsx`` so the main application
can re-load it on subsequent runs without requiring the manager's file again.

The manager's file format (e.g. ``NOVIEMBRE_2026.xlsx``) contains two sheets:

* **Hoja1**: Historical vs. current comparison (not used for budget loading).
* **Hoja2**: Clean budget table for the current month with columns:
    ``Nombre``, ``Zona``, ``Región``, ``<MES>`` (columna D, meta de ventas del vendedor),
    ``NETO <MES>``, ``Presup. Activación``, ``Presup. SKUs``.

La columna del mes (Columna D, e.g. ``NOVIEMBRE``) es la meta de ventas que debe
cumplir el vendedor (``presup_ventas``).

Typical usage::

    from maintenance.budget_loader import load_presupuestos, save_presupuestos_config

    presupuestos = load_presupuestos(manager_file_path)
    save_presupuestos_config(presupuestos, Path("data/presupuestos.xlsx"))
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from core.exceptions import LynxBaseError, MaintenanceError
from core.models import BudgetRecord

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column detection aliases
# ---------------------------------------------------------------------------

_COL_NOMBRE = "Nombre"

# "NETO NOV", "NETO ENE", "NETO DIC", etc.  — the 60% net ventas target
_NETO_MES_PATTERN = re.compile(r"^NETO\s+\w+$", re.IGNORECASE)

# Gross month column: "NOVIEMBRE", "ENERO", "DICIEMBRE", etc.
_MES_NAMES = {
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
}

# Canonical column aliases (from data/presupuestos.xlsx saved by this module)
_COL_VENTAS_ALIASES: List[str] = [
    "VENTAS", "Ventas", "ventas", "Meta Ventas", "PRESUPUESTO VENTAS",
    "NETO NOV", "NETO ENE", "NETO FEB", "NETO MAR", "NETO ABR",
    "NETO MAY", "NETO JUN", "NETO JUL", "NETO AGO", "NETO SEP",
    "NETO OCT", "NETO DIC",
]
_COL_ACTIVACION_ALIASES: List[str] = [
    "ACTIVACIÓN", "ACTIVACION", "Activación", "activacion",
    "Meta Activación", "Presup. Activación", "Presup. Activacion",
]
_COL_SKUS_ALIASES: List[str] = [
    "SKU", "SKUs", "SKUS", "sku", "Meta SKUs",
    "PRESUPUESTO SKU", "Presup. SKUs", "Presup. SKU",
]
_COL_REGION_ALIASES: List[str] = [
    "Region", "REGION", "Región", "REGIÓN", "region", "Regi\u00f3n",
]


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _find_column(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    """Return the first column name from *aliases* that exists in *df*.

    Args:
        df: DataFrame whose columns are searched.
        aliases: Candidate column name strings, tried in order.

    Returns:
        Matching column name, or ``None`` if none found.
    """
    for alias in aliases:
        if alias in df.columns:
            return alias
    return None


def _find_neto_column(df: pd.DataFrame) -> Optional[str]:
    """Detect the ``NETO <MES>`` column automatically from column names.

    Looks for a column whose name matches the pattern ``NETO <word>``
    (case-insensitive), which the manager uses for the 60% net target.

    Args:
        df: DataFrame to search.

    Returns:
        Column name if found, else ``None``.
    """
    for col in df.columns:
        if _NETO_MES_PATTERN.match(str(col)):
            return col
    return None


def _find_mes_column(df: pd.DataFrame) -> Optional[str]:
    """Detect the gross-month column (e.g. ``NOVIEMBRE``) from column names.

    Args:
        df: DataFrame to search.

    Returns:
        Column name if found, else ``None``.
    """
    for col in df.columns:
        if str(col).strip().upper() in _MES_NAMES:
            return col
    return None


def _detect_manager_format(xl: pd.ExcelFile) -> Optional[str]:
    """Check if the Excel file is in the manager's monthly format.

    The manager's file has a ``Hoja2`` sheet with the clean budget table.
    Returns the sheet name to use, or ``None`` if the format is not recognised.

    Args:
        xl: An open :class:`pandas.ExcelFile` object.

    Returns:
        ``"Hoja2"`` if detected, otherwise ``None``.
    """
    if "Hoja2" in xl.sheet_names:
        return "Hoja2"
    return None


def _parse_manager_sheet(df: pd.DataFrame) -> Dict[str, BudgetRecord]:
    """Parse a DataFrame from the manager's monthly budget sheet (Hoja2).

    Expects columns: ``Nombre``, ``Zona``, ``Región``, a gross-month column
    (e.g. ``NOVIEMBRE``), a ``NETO <MES>`` column, ``Presup. Activación``,
    ``Presup. SKUs``.

    The ``NETO <MES>`` value is used as ``presup_ventas`` (the 60% net target
    used in all KPI computations).  If that column is missing, falls back to
    the gross-month column.

    Args:
        df: DataFrame read from the manager's budget sheet.

    Returns:
        Dictionary mapping upper-cased stripped vendor name → BudgetRecord.

    Raises:
        MaintenanceError: If neither a ventas nor region column can be found.
    """
    # --- Nombre column --------------------------------------------------
    col_nombre: Optional[str] = _COL_NOMBRE if _COL_NOMBRE in df.columns else df.columns[0]

    # --- Ventas: prefer columna del mes (columna D, e.g. NOVIEMBRE) ------
    # Prioridad: 1) Nombre del mes (e.g. 'NOVIEMBRE') o columna D (índice 3)
    col_ventas: Optional[str] = _find_mes_column(df)
    if col_ventas is None and len(df.columns) > 3:
        # Columna D directa
        col_ventas = df.columns[3]
    if col_ventas is None:
        col_ventas = _find_neto_column(df)
    if col_ventas is None:
        col_ventas = _find_column(df, _COL_VENTAS_ALIASES)
    if col_ventas is None:
        raise MaintenanceError(
            "No se encontró columna de ventas/presupuesto en la hoja del gerente. "
            "Se esperaba un nombre de mes (e.g. 'NOVIEMBRE') en la columna D.",
            maintenance_task="load_presupuestos",
        )

    # --- Activación -----------------------------------------------------
    col_activacion: Optional[str] = _find_column(df, _COL_ACTIVACION_ALIASES)
    # --- SKUs -----------------------------------------------------------
    col_skus: Optional[str] = _find_column(df, _COL_SKUS_ALIASES)
    # --- Region ---------------------------------------------------------
    col_region: Optional[str] = _find_column(df, _COL_REGION_ALIASES)

    logger.info(
        "Manager budget sheet columns detected → ventas='%s', activacion='%s', "
        "skus='%s', region='%s'",
        col_ventas, col_activacion, col_skus, col_region,
    )

    # Drop rows where Nombre is blank / NaN (includes the totals footer row)
    df = df.copy()
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    df = df[df[col_nombre].notna() & (df[col_nombre] != "") & (df[col_nombre].str.lower() != "nan")]

    # Coerce numerics
    for col in filter(None, [col_ventas, col_activacion, col_skus]):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    result: Dict[str, BudgetRecord] = {}
    for _, row in df.iterrows():
        name = str(row[col_nombre]).strip()
        if not name or name.lower() == "nan":
            continue

        # Normalización de alias conocidos entre presupuesto y Odoo ERP
        if name.upper().strip() == "AMILCAR LEON":
            name = "AMILCAR LEON(BOLIVAR)"

        key = name.upper().strip()
        region_val = str(row[col_region]).strip().upper() if col_region and pd.notna(row.get(col_region)) else ""
        result[key] = BudgetRecord(
            vendor_name=name,
            presup_ventas=float(row[col_ventas]) if col_ventas else 0.0,
            presup_activacion=float(row[col_activacion]) if col_activacion else 0.0,
            presup_skus=float(row[col_skus]) if col_skus else 0.0,
            region=region_val,
        )

    logger.info("Manager budget sheet parsed: %d vendors.", len(result))
    return result


def _parse_canonical_sheet(df: pd.DataFrame) -> Dict[str, BudgetRecord]:
    """Parse a DataFrame in the canonical presupuestos format.

    The canonical format is the one saved by :func:`save_presupuestos_config`
    with columns: ``Nombre``, ``VENTAS``, ``ACTIVACIÓN``, ``SKU``, ``Region``.

    Args:
        df: DataFrame from the canonical budget file.

    Returns:
        Dictionary mapping upper-cased stripped vendor name → BudgetRecord.

    Raises:
        MaintenanceError: If required columns are missing.
    """
    col_nombre: Optional[str] = _COL_NOMBRE if _COL_NOMBRE in df.columns else None
    if col_nombre is None:
        raise MaintenanceError(
            "Columna 'Nombre' no encontrada en el archivo canónico de presupuestos.",
            maintenance_task="load_presupuestos",
        )

    col_ventas = _find_column(df, _COL_VENTAS_ALIASES)
    col_activacion = _find_column(df, _COL_ACTIVACION_ALIASES)
    col_skus = _find_column(df, _COL_SKUS_ALIASES)
    col_region = _find_column(df, _COL_REGION_ALIASES)

    missing = [n for n, c in [("VENTAS", col_ventas), ("ACTIVACION", col_activacion), ("SKU", col_skus)] if c is None]
    if missing:
        raise MaintenanceError(
            f"Columnas no encontradas en presupuestos canónico: {missing}",
            maintenance_task="load_presupuestos",
        )

    df = df.copy()
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    df = df[df[col_nombre].notna() & (df[col_nombre] != "") & (df[col_nombre].str.lower() != "nan")]

    for col in filter(None, [col_ventas, col_activacion, col_skus]):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    result: Dict[str, BudgetRecord] = {}
    for _, row in df.iterrows():
        name = str(row[col_nombre]).strip()
        if not name or name.lower() == "nan":
            continue
        key = name.upper()
        region_val = str(row[col_region]).strip().upper() if col_region and pd.notna(row.get(col_region)) else ""
        result[key] = BudgetRecord(
            vendor_name=name,
            presup_ventas=float(row[col_ventas]),
            presup_activacion=float(row[col_activacion]),
            presup_skus=float(row[col_skus]),
            region=region_val,
        )

    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def load_presupuestos(path: str | Path) -> Dict[str, BudgetRecord]:
    """Load vendor budget targets from the manager's monthly Excel file.

    Automatically detects the file format:

    * **Manager format** (e.g. ``NOVIEMBRE_2026.xlsx``): Contains a ``Hoja2``
      sheet with columns ``Nombre``, ``Zona``, ``Región``, ``NOVIEMBRE``,
      ``NETO NOV``, ``Presup. Activación``, ``Presup. SKUs``.  The
      ``NETO NOV`` (60 % net) value is used as ``presup_ventas``.

    * **Canonical format** (``data/presupuestos.xlsx`` saved by this module):
      Columns ``Nombre``, ``VENTAS``, ``ACTIVACIÓN``, ``SKU``, ``Region``.

    Args:
        path: Path to the Excel file (any ``.xlsx``).

    Returns:
        Dictionary mapping normalised vendor name (stripped, upper-cased) to a
        :class:`~core.models.BudgetRecord`.

    Raises:
        MaintenanceError: If the file cannot be parsed or mandatory columns are
            missing after format detection.

    Example::

        budgets = load_presupuestos("/mnt/shares/NOVIEMBRE_2026.xlsx")
        print(budgets["MARIA GARCIA"].presup_ventas)
    """
    path = Path(path)
    logger.info("Loading Presupuestos from: %s", path)

    if not path.exists():
        raise MaintenanceError(
            f"El archivo de presupuestos no existe: {path}",
            maintenance_task="load_presupuestos",
        )

    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception as exc:
        raise MaintenanceError(
            f"No se pudo abrir el archivo '{path.name}': {exc}",
            maintenance_task="load_presupuestos",
        ) from exc

    manager_sheet = _detect_manager_format(xl)

    if manager_sheet:
        # Manager's monthly file → use Hoja2
        logger.info("Detected manager format; reading sheet '%s'.", manager_sheet)
        try:
            df = xl.parse(manager_sheet, header=0)
        except Exception as exc:
            raise MaintenanceError(
                f"Error leyendo hoja '{manager_sheet}' de '{path.name}': {exc}",
                maintenance_task="load_presupuestos",
            ) from exc
        return _parse_manager_sheet(df)
    else:
        # Canonical presupuestos.xlsx format (single sheet)
        logger.info("Detected canonical format; reading first sheet.")
        try:
            df = xl.parse(0, header=0)
        except Exception as exc:
            raise MaintenanceError(
                f"Error leyendo archivo canónico '{path.name}': {exc}",
                maintenance_task="load_presupuestos",
            ) from exc
        return _parse_canonical_sheet(df)


def save_presupuestos_config(
    presupuestos: Dict[str, BudgetRecord],
    config_path: Path,
) -> None:
    """Persist the loaded budget dictionary to an Excel file.

    The saved file uses the canonical column names expected by
    :func:`core.etl.read_presupuestos`, so it can be re-loaded on the next
    run without requiring the manager's file.

    Args:
        presupuestos: Mapping from vendor key (upper-cased nombre) to
            :class:`~core.models.BudgetRecord`, as returned by
            :func:`load_presupuestos`.
        config_path: Destination ``.xlsx`` file path.  Parent directory is
            created if it does not exist.

    Raises:
        MaintenanceError: If the file cannot be written.

    Example::

        save_presupuestos_config(budgets, Path("data/presupuestos.xlsx"))
    """
    config_path = Path(config_path)
    config_path.parent.mkdir(parents=True, exist_ok=True)

    rows = [
        {
            "Nombre": rec.vendor_name,
            "VENTAS": rec.presup_ventas,
            "ACTIVACIÓN": rec.presup_activacion,
            "SKU": rec.presup_skus,
            "Region": rec.region,
        }
        for rec in presupuestos.values()
    ]

    df = pd.DataFrame(rows)
    try:
        df.to_excel(config_path, index=False, engine="openpyxl")
        logger.info(
            "Presupuestos config saved: %s (%d vendors)", config_path, len(rows)
        )
    except Exception as exc:
        raise MaintenanceError(
            f"No se pudo guardar la configuración de presupuestos: {exc}",
            maintenance_task="save_presupuestos_config",
        ) from exc


def load_presupuestos_from_config(config_path: Path) -> Dict[str, BudgetRecord]:
    """Load vendor budgets from the application's persisted configuration file.

    This is the fast-path for subsequent runs: instead of asking the user to
    provide the manager's file again, the application reads the canonically
    formatted file saved by :func:`save_presupuestos_config`.

    Args:
        config_path: Path to ``data/presupuestos.xlsx`` (or equivalent).

    Returns:
        Dictionary mapping upper-cased vendor name to
        :class:`~core.models.BudgetRecord`, or an empty dict if the file does
        not exist.

    Raises:
        MaintenanceError: If the file exists but cannot be parsed.

    Example::

        budgets = load_presupuestos_from_config(Path("data/presupuestos.xlsx"))
        if not budgets:
            budgets = load_presupuestos(manager_file)
    """
    config_path = Path(config_path)

    if not config_path.exists():
        logger.info(
            "Presupuestos config not found at '%s'; returning empty dict.", config_path
        )
        return {}

    try:
        result = load_presupuestos(str(config_path))
        logger.info(
            "Presupuestos loaded from config: %d vendors", len(result)
        )
        return result
    except MaintenanceError as exc:
        raise MaintenanceError(
            f"Error leyendo presupuestos de configuración '{config_path.name}': {exc}",
            maintenance_task="load_presupuestos_from_config",
        ) from exc
