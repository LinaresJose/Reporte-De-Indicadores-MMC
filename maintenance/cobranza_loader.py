"""maintenance/cobranza_loader.py — LYNX Report System.

Loads the monthly "cobranza cierre" data from the manager's
``INDICADOR_COBRANZA_MES_AÑO.xlsx`` file, specifically the "RESUMEN" sheet.

At month-end the sales manager sends this file.  The "RESUMEN" sheet contains
one row per vendor with their final collections performance (amounts overdue
vs. current).  This data **overrides** the daily CxC calculation for
``pct_sin_vencer`` in the monthly closing report.

Typical usage::

    from maintenance.cobranza_loader import load_cobranza_cierre

    cobranza = load_cobranza_cierre("INDICADOR_COBRANZA_OCTUBRE_2026.xlsx")
    # returns: {"MARIA GARCIA": CobranzaAgg(sin_vencer=..., vencido=...), ...}
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import Dict, Optional, Set

import pandas as pd

from core.exceptions import MaintenanceError, MissingColumnError
from core.models import CobranzaAgg

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column name constants (from INDICADOR_COBRANZA RESUMEN sheet)
# ---------------------------------------------------------------------------

_SHEET_RESUMEN = "RESUMEN"

# Primary column names as observed in the manager's file
_COL_VENDEDOR = "VENDEDOR"
_COL_SIN_VENCER = "Total Sin Vencer"
_COL_VENCIDO = "Total Vencido"

# Fallback aliases in case the manager's file has slightly different headers
_COL_VENDEDOR_ALIASES = ["VENDEDOR", "Vendedor", "COMISIONISTA", "Nombre", "NOMBRE"]
_COL_SIN_VENCER_ALIASES = [
    "Total Sin Vencer",
    "TOTAL SIN VENCER",
    "Sin Vencer",
    "SIN VENCER",
    "Importe Sin Vencer",
]
_COL_VENCIDO_ALIASES = [
    "Total Vencido",
    "TOTAL VENCIDO",
    "Vencido",
    "VENCIDO",
    "Importe Vencido",
]


def _find_column(df: pd.DataFrame, aliases: list[str]) -> Optional[str]:
    """Return the first alias that matches a column name in *df*.

    Args:
        df: DataFrame to search.
        aliases: Candidate column names, tried in order.

    Returns:
        First matching column name, or ``None`` if none found.
    """
    for alias in aliases:
        if alias in df.columns:
            return alias
    return None


def load_cobranza_cierre(path: str | Path) -> Dict[str, CobranzaAgg]:
    """Load month-end collections data from the manager's indicator file.

    Reads the **"RESUMEN"** sheet from an ``INDICADOR_COBRANZA_*.xlsx`` file and
    returns a dictionary of :class:`~core.models.CobranzaAgg` objects, one per
    vendor, keyed by upper-cased vendor name.

    The returned data is intended to **replace** the daily CxC calculation in
    :func:`core.engine.build_vendor_record` when performing month-end closing.

    Args:
        path: Path to the ``INDICADOR_COBRANZA_MES_AÑO.xlsx`` file.

    Returns:
        Dictionary mapping ``VENDEDOR.upper().strip()`` →
        :class:`~core.models.CobranzaAgg` with:

        - ``sin_vencer``: Total amount not yet due (``Total Sin Vencer``).
        - ``vencido``: Total past-due amount (``Total Vencido``).
        - ``pct_sin_vencer``: Computed as ``sin_vencer / (sin_vencer + vencido)``
          (0.0 if the denominator is zero).

    Raises:
        MaintenanceError: If the file does not exist, the "RESUMEN" sheet is
            missing, or mandatory columns cannot be resolved.
        MissingColumnError: If the VENDEDOR column cannot be identified even
            after alias resolution.

    Example::

        cobranza = load_cobranza_cierre("INDICADOR_COBRANZA_OCTUBRE_2026.xlsx")
        entry = cobranza.get("PEDRO LOPEZ")
        if entry:
            print(f"Sin vencer: {entry.sin_vencer:,.2f}")
    """
    path = Path(path)
    logger.info("Loading Cobranza Cierre from: %s", path)

    if not path.exists():
        raise MaintenanceError(
            f"El archivo de cobranza cierre no existe: {path}",
            maintenance_task="load_cobranza_cierre",
        )

    # ------------------------------------------------------------------
    # Read the RESUMEN sheet
    # ------------------------------------------------------------------
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
        available_sheets = xl.sheet_names
    except Exception as exc:
        raise MaintenanceError(
            f"No se pudo abrir '{path.name}': {exc}",
            maintenance_task="load_cobranza_cierre",
        ) from exc

    if _SHEET_RESUMEN not in available_sheets:
        raise MaintenanceError(
            f"La hoja '{_SHEET_RESUMEN}' no existe en '{path.name}'. "
            f"Hojas disponibles: {available_sheets}",
            maintenance_task="load_cobranza_cierre",
        )

    try:
        df = pd.read_excel(xl, sheet_name=_SHEET_RESUMEN)
    except Exception as exc:
        raise MaintenanceError(
            f"No se pudo leer la hoja '{_SHEET_RESUMEN}' de '{path.name}': {exc}",
            maintenance_task="load_cobranza_cierre",
        ) from exc

    logger.debug(
        "RESUMEN sheet read: %d rows, columns: %s", len(df), list(df.columns)
    )

    # ------------------------------------------------------------------
    # Resolve column names (with alias fallback)
    # ------------------------------------------------------------------
    col_vendedor = _find_column(df, _COL_VENDEDOR_ALIASES)
    col_sin_vencer = _find_column(df, _COL_SIN_VENCER_ALIASES)
    col_vencido = _find_column(df, _COL_VENCIDO_ALIASES)

    if col_vendedor is None:
        raise MissingColumnError(
            sheet_name=_SHEET_RESUMEN,
            missing_columns=["VENDEDOR"],
            source_file=path.name,
        )

    if col_sin_vencer is None or col_vencido is None:
        missing = []
        if col_sin_vencer is None:
            missing.append("Total Sin Vencer")
        if col_vencido is None:
            missing.append("Total Vencido")
        raise MissingColumnError(
            sheet_name=_SHEET_RESUMEN,
            missing_columns=missing,
            source_file=path.name,
        )

    # ------------------------------------------------------------------
    # Clean and parse data
    # ------------------------------------------------------------------
    df = df[[col_vendedor, col_sin_vencer, col_vencido]].copy()
    df.columns = ["vendedor", "sin_vencer", "vencido"]

    # Drop rows where vendedor is null/empty or looks like a header repeat
    df["vendedor"] = df["vendedor"].astype(str).str.strip()
    df = df[df["vendedor"].str.len() > 0]
    df = df[df["vendedor"].str.upper() != "VENDEDOR"]  # remove header rows

    # Coerce numeric columns — treat errors as 0
    df["sin_vencer"] = pd.to_numeric(df["sin_vencer"], errors="coerce").fillna(0.0)
    df["vencido"] = pd.to_numeric(df["vencido"], errors="coerce").fillna(0.0)

    # ------------------------------------------------------------------
    # Build result dictionary
    # ------------------------------------------------------------------
    result: Dict[str, CobranzaAgg] = {}

    for _, row in df.iterrows():
        nombre: str = row["vendedor"]
        sin_vencer: float = float(row["sin_vencer"])
        vencido: float = float(row["vencido"])

        key = nombre.upper()
        result[key] = CobranzaAgg(
            vendor_name=nombre,
            total_sin_vencer=sin_vencer,
            total_vencido=vencido,
        )

    logger.info(
        "Cobranza cierre loaded: %d vendors from '%s'", len(result), path.name
    )
    return result


def save_cobranza_cierre(
    source_path: str | Path,
    dest_path: Optional[str | Path] = None,
) -> Path:
    """Save or copy the manager's cobranza closing file to the persistent data folder.

    Args:
        source_path: Path to the manager's ``INDICADOR_COBRANZA_*.xlsx`` file.
        dest_path: Optional destination path. If not provided, defaults to
            ``data/cobranza_cierre.xlsx``.

    Returns:
        The destination :class:`Path` where the file was saved.

    Raises:
        MaintenanceError: If the source file cannot be read or written.
    """
    src = Path(source_path)
    if not src.exists():
        raise MaintenanceError(
            f"El archivo origen de cobranza no existe: {src}",
            maintenance_task="save_cobranza_cierre",
        )

    if dest_path is None:
        dest = Path(__file__).parent.parent / "data" / "cobranza_cierre.xlsx"
    else:
        dest = Path(dest_path)

    dest.parent.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copy2(src, dest)
        logger.info("Cobranza cierre persistent copy created at '%s'", dest)
        return dest
    except Exception as exc:
        raise MaintenanceError(
            f"No se pudo guardar la cobranza de cierre en '{dest}': {exc}",
            maintenance_task="save_cobranza_cierre",
        ) from exc


def load_vip_clients(path: Optional[str | Path] = None) -> Set[str]:
    """Load the list of VIP clients who have extended credit terms (45 days).

    Looks for a sheet named ``'VIP'`` in ``data/cobranza_cierre.xlsx`` (or the
    supplied *path*).

    Args:
        path: Path to the file containing a 'VIP' sheet. Defaults to
            ``data/cobranza_cierre.xlsx``.

    Returns:
        Set of uppercase, stripped client names that qualify as VIP.
    """
    if path is None:
        target = Path(__file__).parent.parent / "data" / "cobranza_cierre.xlsx"
    else:
        target = Path(path)

    if not target.exists():
        return set()

    try:
        xl = pd.ExcelFile(target, engine="openpyxl")
        if "VIP" not in xl.sheet_names:
            return set()
        df = pd.read_excel(xl, sheet_name="VIP")
        # Try finding client column
        client_col = None
        for col in ["CLIENTE", "Cliente", "Nombre", "Socio"]:
            if col in df.columns:
                client_col = col
                break
        if client_col is None and len(df.columns) > 1:
            client_col = df.columns[1]  # Typically second column is CLIENTE

        if client_col is None:
            return set()

        vip_set = {
            str(c).strip().upper()
            for c in df[client_col].dropna()
            if str(c).strip() and str(c).strip().upper() != "CLIENTE"
        }
        logger.debug("Loaded %d VIP clients from '%s'", len(vip_set), target.name)
        return vip_set
    except Exception as exc:
        logger.warning("Could not load VIP clients from '%s': %s", target, exc)
        return set()

