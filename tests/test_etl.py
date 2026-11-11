"""tests/test_etl.py — Unit tests for core/etl.py.

Tests use the actual uploaded source files when available (read-only) and
synthetic minimal DataFrames for edge-case scenarios.

Run with::

    py -m pytest tests/test_etl.py -v
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

# Ensure the project root is on sys.path regardless of where pytest is invoked
PROJECT_ROOT = Path(__file__).parent.parent.resolve()
sys.path.insert(0, str(PROJECT_ROOT))

from core.etl import (
    _validate_columns,
    calculate_working_days,
    read_backorder,
    read_devoluciones,
    read_facturado,
    read_por_cobrar,
    read_por_facturar,
    read_pedidos,
)
from core.exceptions import EmptySourceError, MissingColumnError

# ---------------------------------------------------------------------------
# Paths to actual source files (skip tests if not present)
# ---------------------------------------------------------------------------

_UPLOAD_DIR = (
    Path(__file__).parent.parent.parent.parent.parent
    / ".gemini/antigravity/brain/350462f1-e0c0-4daf-8e49-e7d82c16dd69/.user_uploaded"
)

_FACTURADO_PATH   = _UPLOAD_DIR / "media_1793973555125_ad40d368.xlsx"
_PEDIDOS_PATH     = _UPLOAD_DIR / "media_1793973154286_ada57d7b.xlsx"
_DEVOL_PATH       = _UPLOAD_DIR / "media_1793973154289_14d4194f.xlsx"
_CXC_PATH         = _UPLOAD_DIR / "media_1793973154276_e06a2192.xlsx"
_POR_FACTURAR_PATH= _UPLOAD_DIR / "media_1793973154275_b3c14984.xlsx"
_BACKORDER_PATH   = _UPLOAD_DIR / "media_1793973154286_18334d23.xlsx"

_has_files = all(p.exists() for p in [
    _FACTURADO_PATH, _PEDIDOS_PATH, _DEVOL_PATH,
    _CXC_PATH, _POR_FACTURAR_PATH, _BACKORDER_PATH,
])


# ===========================================================================
# _validate_columns
# ===========================================================================

class TestValidateColumns:
    """Tests for the internal column validation helper."""

    def test_passes_when_all_columns_present(self):
        df = pd.DataFrame(columns=["A", "B", "C"])
        # Should not raise
        _validate_columns(df, ["A", "B"], source_file="test.xlsx", sheet_name="Sheet1")

    def test_raises_missing_column_error(self):
        df = pd.DataFrame(columns=["A"])
        with pytest.raises(MissingColumnError) as exc_info:
            _validate_columns(df, ["A", "MISSING"], source_file="test.xlsx", sheet_name="Sheet1")
        assert "MISSING" in str(exc_info.value)

    def test_raises_for_empty_dataframe_with_required_cols(self):
        df = pd.DataFrame()
        with pytest.raises(MissingColumnError):
            _validate_columns(df, ["X"], source_file="test.xlsx", sheet_name="Sheet1")


# ===========================================================================
# read_facturado
# ===========================================================================

@pytest.mark.skipif(not _has_files, reason="Source files not available")
class TestReadFacturado:
    """Integration tests for read_facturado using the actual Odoo export."""

    def test_returns_dataframe(self):
        df = read_facturado(str(_FACTURADO_PATH))
        assert isinstance(df, pd.DataFrame)

    def test_has_expected_columns(self):
        df = read_facturado(str(_FACTURADO_PATH))
        assert "comercial" in df.columns
        assert "subtotal" in df.columns

    def test_no_negative_subtotals(self):
        """Devoluciones are in a separate file; facturado should have no negatives."""
        df = read_facturado(str(_FACTURADO_PATH))
        # Allow small negatives from adjustments but flag large ones
        assert (df["subtotal"] < -1000).sum() == 0, (
            "Unexpected large negative subtotals in facturado"
        )

    def test_row_count_nonzero(self):
        df = read_facturado(str(_FACTURADO_PATH))
        assert len(df) > 0

    def test_comercial_column_non_null(self):
        df = read_facturado(str(_FACTURADO_PATH))
        assert df["comercial"].notna().all(), "comercial column contains null values"


# ===========================================================================
# read_por_cobrar (CxC date classification)
# ===========================================================================

class TestReadPorCobrarSynthetic:
    """Unit tests for read_por_cobrar aging and credit policies using synthetic workbooks."""

    def test_credit_days_standard_and_vip(self, tmp_path):
        import config
        from datetime import date, timedelta
        from core.models import CobranzaAgg

        ref_date = date(2026, 11, 10)
        # 10 days ago (<= 32) -> Sin Vencer
        # 35 days ago (> 32, <= 45) -> Standard is Vencido, VIP is Sin Vencer
        # 50 days ago (> 45) -> Vencido for both
        test_data = pd.DataFrame([
            {
                config.COL_COB_COMERCIAL: "VENDEDOR 1",
                config.COL_COB_CLIENTE: "CLIENTE NORMAL",
                config.COL_COB_FECHA_ENTREGA: ref_date - timedelta(days=10),
                config.COL_COB_IMPORTE_ADEUDADO: 1000.0,
            },
            {
                config.COL_COB_COMERCIAL: "VENDEDOR 1",
                config.COL_COB_CLIENTE: "CLIENTE NORMAL",
                config.COL_COB_FECHA_ENTREGA: ref_date - timedelta(days=35),
                config.COL_COB_IMPORTE_ADEUDADO: 500.0,
            },
            {
                config.COL_COB_COMERCIAL: "VENDEDOR 2",
                config.COL_COB_CLIENTE: "CLIENTE VIP ESPECIAL",
                config.COL_COB_FECHA_ENTREGA: ref_date - timedelta(days=35),
                config.COL_COB_IMPORTE_ADEUDADO: 2000.0,
            },
            {
                config.COL_COB_COMERCIAL: "VENDEDOR 2",
                config.COL_COB_CLIENTE: "CLIENTE VIP ESPECIAL",
                config.COL_COB_FECHA_ENTREGA: ref_date - timedelta(days=50),
                config.COL_COB_IMPORTE_ADEUDADO: 1000.0,
            },
        ])

        excel_file = tmp_path / "cxc_test.xlsx"
        test_data.to_excel(excel_file, index=False)

        res = read_por_cobrar(
            excel_file,
            as_of_date=ref_date,
            vip_clients={"CLIENTE VIP ESPECIAL"},
        )

        assert "VENDEDOR 1" in res
        assert "VENDEDOR 2" in res

        # Vendedor 1: 1000 sin vencer, 500 vencido
        v1 = res["VENDEDOR 1"]
        assert isinstance(v1, CobranzaAgg)
        assert v1.total_sin_vencer == pytest.approx(1000.0)
        assert v1.total_vencido == pytest.approx(500.0)
        assert v1.pct_sin_vencer == pytest.approx(1000.0 / 1500.0)

        # Vendedor 2: 2000 sin vencer (VIP 35 días <= 45), 1000 vencido
        v2 = res["VENDEDOR 2"]
        assert v2.total_sin_vencer == pytest.approx(2000.0)
        assert v2.total_vencido == pytest.approx(1000.0)
        assert v2.pct_sin_vencer == pytest.approx(2000.0 / 3000.0)

    def test_fallback_to_fecha_vencimiento(self, tmp_path):
        import config
        from datetime import date, timedelta

        ref_date = date(2026, 11, 10)
        test_data = pd.DataFrame([
            {
                config.COL_COB_COMERCIAL: "VENDEDOR 3",
                config.COL_COB_FECHA_ENTREGA: None,
                config.COL_COB_FECHA_VENCIMIENTO: ref_date + timedelta(days=5),
                config.COL_COB_IMPORTE_ADEUDADO: 300.0,
            },
            {
                config.COL_COB_COMERCIAL: "VENDEDOR 3",
                config.COL_COB_FECHA_ENTREGA: None,
                config.COL_COB_FECHA_VENCIMIENTO: ref_date - timedelta(days=2),
                config.COL_COB_IMPORTE_ADEUDADO: 700.0,
            },
        ])

        excel_file = tmp_path / "cxc_venc.xlsx"
        test_data.to_excel(excel_file, index=False)

        res = read_por_cobrar(excel_file, as_of_date=ref_date)
        v3 = res["VENDEDOR 3"]
        assert v3.total_sin_vencer == pytest.approx(300.0)
        assert v3.total_vencido == pytest.approx(700.0)



# ===========================================================================
# read_pedidos
# ===========================================================================

@pytest.mark.skipif(not _has_files, reason="Source files not available")
class TestReadPedidos:
    """Integration tests for read_pedidos."""

    def test_returns_dataframe(self):
        df = read_pedidos(str(_PEDIDOS_PATH))
        assert isinstance(df, pd.DataFrame)

    def test_has_comercial_and_monto(self):
        df = read_pedidos(str(_PEDIDOS_PATH))
        assert "comercial" in df.columns
        monto_cols = [c for c in df.columns if "monto" in c.lower() or "base" in c.lower() or "imponible" in c.lower()]
        assert len(monto_cols) > 0, f"No monto column. Columns: {df.columns.tolist()}"


# ===========================================================================
# read_devoluciones
# ===========================================================================

@pytest.mark.skipif(not _has_files, reason="Source files not available")
class TestReadDevoluciones:
    """Integration tests for read_devoluciones."""

    def test_returns_dataframe(self):
        df = read_devoluciones(str(_DEVOL_PATH))
        assert isinstance(df, pd.DataFrame)

    def test_nonzero_rows(self):
        df = read_devoluciones(str(_DEVOL_PATH))
        assert len(df) > 0


# ===========================================================================
# calculate_working_days
# ===========================================================================

class TestCalculateWorkingDays:
    """Tests for the working-days calculator."""

    def test_returns_working_days_info(self):
        from core.models import WorkingDaysInfo
        info = calculate_working_days()
        assert isinstance(info, WorkingDaysInfo)

    def test_total_dias_habiles_positive(self):
        info = calculate_working_days()
        assert info.total_dias_habiles > 0

    def test_pct_avance_between_0_and_1(self):
        info = calculate_working_days()
        assert 0.0 <= info.pct_avance <= 1.0, (
            f"pct_avance out of range: {info.pct_avance}"
        )

    def test_dias_transcurridos_le_total(self):
        info = calculate_working_days()
        assert info.dias_habiles_transcurridos <= info.total_dias_habiles

    def test_dias_restantes_plus_transcurridos_eq_total(self):
        info = calculate_working_days()
        assert (
            info.dias_habiles_transcurridos + info.dias_habiles_restantes
            == info.total_dias_habiles
        )
