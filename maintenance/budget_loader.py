"""maintenance/budget_loader.py — LYNX Report System.

Handles the monthly Presupuestos (budget) maintenance workflow.

The sales manager delivers a new Excel file each month with updated budget
targets per vendor (VENTAS, ACTIVACIÓN, SKUs).  This module loads that file
and persists the data to ``data/presupuestos.xlsx`` so the main application
can re-load it on subsequent runs without requiring the manager's file again.

Typical usage::

    from maintenance.budget_loader import load_presupuestos, save_presupuestos_config

    presupuestos = load_presupuestos(manager_file_path)
    save_presupuestos_config(presupuestos, Path("data/presupuestos.xlsx"))
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict

import openpyxl
import pandas as pd

from core.etl import read_presupuestos
from core.exceptions import LynxBaseError, MaintenanceError
from core.models import BudgetRecord

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column name aliases accepted from the manager's file
# ---------------------------------------------------------------------------

_COL_NOMBRE = "Nombre"
_COL_VENTAS_ALIASES = ["VENTAS", "Ventas", "ventas", "Meta Ventas", "PRESUPUESTO VENTAS"]
_COL_ACTIVACION_ALIASES = [
    "ACTIVACIÓN",
    "ACTIVACION",
    "Activación",
    "activacion",
    "Meta Activación",
]
_COL_SKUS_ALIASES = ["SKU", "SKUs", "SKUS", "sku", "Meta SKUs", "PRESUPUESTO SKU"]
_COL_REGION_ALIASES = ["Region", "REGION", "Región", "REGIÓN", "region"]


def _find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
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


def load_presupuestos(path: str | Path) -> Dict[str, BudgetRecord]:
    """Load vendor budget targets from the manager's monthly Excel file.

    Delegates primary parsing to :func:`core.etl.read_presupuestos`.  If that
    function raises due to unexpected column names, this function attempts a
    fuzzy column-alias fallback before giving up.

    Args:
        path: Path to the manager's Excel file (any ``.xlsx`` or ``.xls``).

    Returns:
        Dictionary mapping normalised vendor name (stripped, upper-cased) to a
        :class:`~core.models.BudgetRecord`.

    Raises:
        MaintenanceError: If the file cannot be parsed or mandatory columns are
            missing after alias resolution.

    Example::

        budgets = load_presupuestos("/mnt/shares/presupuestos_nov_2026.xlsx")
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
        # Primary path: delegate to canonical ETL reader
        result = read_presupuestos(str(path))
        logger.info(
            "Presupuestos loaded via ETL: %d vendors", len(result)
        )
        return result

    except LynxBaseError as exc:
        logger.warning(
            "ETL reader failed (%s), attempting alias fallback.", exc
        )

    # -------------------------------------------------------------------
    # Alias fallback: try to read with flexible column names
    # -------------------------------------------------------------------
    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as exc:
        raise MaintenanceError(
            f"No se pudo leer el archivo de presupuestos '{path.name}': {exc}",
            maintenance_task="load_presupuestos",
        ) from exc

    # Resolve columns
    col_nombre = _COL_NOMBRE if _COL_NOMBRE in df.columns else None
    if col_nombre is None:
        # Try first column as nombre
        col_nombre = df.columns[0]
        logger.warning(
            "Columna '%s' no encontrada; usando primera columna '%s' como Nombre.",
            _COL_NOMBRE,
            col_nombre,
        )

    col_ventas = _find_column(df, _COL_VENTAS_ALIASES)
    col_activacion = _find_column(df, _COL_ACTIVACION_ALIASES)
    col_skus = _find_column(df, _COL_SKUS_ALIASES)
    col_region = _find_column(df, _COL_REGION_ALIASES)

    missing = [
        name
        for name, col in [
            ("VENTAS", col_ventas),
            ("ACTIVACION", col_activacion),
            ("SKU", col_skus),
        ]
        if col is None
    ]
    if missing:
        raise MaintenanceError(
            f"Columnas no encontradas en presupuestos tras alias fallback: {missing}",
            maintenance_task="load_presupuestos",
        )

    df = df.dropna(subset=[col_nombre])
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    df = df[df[col_nombre] != ""]

    result: Dict[str, BudgetRecord] = {}
    for _, row in df.iterrows():
        nombre = str(row[col_nombre]).strip()
        key = nombre.upper()
        result[key] = BudgetRecord(
            nombre=nombre,
            presup_ventas=float(row.get(col_ventas, 0) or 0),
            presup_activacion=float(row.get(col_activacion, 0) or 0),
            presup_skus=int(row.get(col_skus, 0) or 0),
            region=str(row.get(col_region, "")) if col_region else "",
        )

    logger.info(
        "Presupuestos loaded via alias fallback: %d vendors", len(result)
    )
    return result


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
            "Nombre": rec.nombre,
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
        result = read_presupuestos(str(config_path))
        logger.info(
            "Presupuestos loaded from config: %d vendors", len(result)
        )
        return result
    except LynxBaseError as exc:
        raise MaintenanceError(
            f"Error leyendo presupuestos de configuración '{config_path.name}': {exc}",
            maintenance_task="load_presupuestos_from_config",
        ) from exc
