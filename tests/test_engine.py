"""tests/test_engine.py — Unit tests for core/engine.py."""
from __future__ import annotations
import sys
from pathlib import Path
import pytest

PROJECT_ROOT = Path(__file__).parent.parent.resolve()
sys.path.insert(0, str(PROJECT_ROOT))

from core.engine import (
    _safe_divide, _calc_alcance_puntos, _apply_vendor_aliases,
    build_vendor_record, calculate_rankings, filter_by_region, get_regional_totals,
    calculate_general,
)
from core.models import (
    BackorderAgg, BudgetRecord, CobranzaAgg, DevolucionesAgg,
    FacturadoAgg, PedidosAgg, PorFacturarAgg, VendorRecord, ZonaInfo,
)
import config

def _make_zona(nombre="TEST VENDEDOR", zona="ZONA-01", region="CAPITAL"):
    return ZonaInfo(vendor_name=nombre, zona=zona, region=region, correo="", codigo="")

def _make_budget(ventas=100_000.0, activacion=50, skus=20):
    return BudgetRecord(vendor_name="TEST", presup_ventas=ventas, presup_activacion=activacion, presup_skus=skus, region="CAPITAL")

def _make_facturado(total=90_000.0, clientes=45, skus=18):
    return FacturadoAgg(vendor_name="TEST", total_facturado=total, clientes_activados=clientes, skus_vendidos=skus)

def _make_pedidos(monto=10_000.0, cant=5):
    return PedidosAgg(vendor_name="TEST", monto_pedidos=monto, cantidad_pedidos=cant)

def _make_devoluciones(monto=5_000.0):
    return DevolucionesAgg(vendor_name="TEST", total_devoluciones=monto)

def _make_cobranza(sin_vencer=80_000.0, vencido=20_000.0):
    return CobranzaAgg(vendor_name="TEST", total_sin_vencer=sin_vencer, total_vencido=vencido)

def _make_por_facturar(monto=2_000.0):
    return PorFacturarAgg(vendor_name="TEST", pedidos_unicos=1, total_base_imponible=monto)

def _make_backorder(monto=1_500.0):
    return BackorderAgg(vendor_name="TEST", pedidos_unicos=1, backorder_value=monto)

def _build_record(**kwargs):
    defaults = dict(
        zona_info=_make_zona(), budget=_make_budget(),
        facturado=_make_facturado(), pedidos=_make_pedidos(),
        devoluciones=_make_devoluciones(), cobranza=_make_cobranza(),
        por_facturar=_make_por_facturar(), backorder=_make_backorder(),
    )
    defaults.update(kwargs)
    return build_vendor_record(**defaults)


class TestSafeDivide:
    def test_normal_division(self):
        assert _safe_divide(10.0, 4.0) == pytest.approx(2.5)
    def test_zero_denominator_returns_zero(self):
        assert _safe_divide(5.0, 0.0) == 0.0
    def test_zero_numerator(self):
        assert _safe_divide(0.0, 100.0) == 0.0
    def test_negative_values(self):
        assert _safe_divide(-50.0, 10.0) == pytest.approx(-5.0)
    def test_both_zero(self):
        assert _safe_divide(0.0, 0.0) == 0.0


class TestCalcAlcancePuntos:
    def test_full_achievement(self):
        alcance, puntos = _calc_alcance_puntos(100.0, 100.0, max_puntos=30)
        assert alcance == pytest.approx(1.0)
        assert puntos == pytest.approx(30.0)
    def test_over_achievement_capped_at_max(self):
        _, puntos = _calc_alcance_puntos(120.0, 100.0, max_puntos=30)
        assert puntos == pytest.approx(30.0)
    def test_zero_budget_returns_zero(self):
        alcance, puntos = _calc_alcance_puntos(50.0, 0.0, max_puntos=10)
        assert alcance == 0.0 and puntos == 0.0
    def test_partial_achievement(self):
        alcance, puntos = _calc_alcance_puntos(80.0, 100.0, max_puntos=30)
        assert alcance == pytest.approx(0.80)
        assert puntos == pytest.approx(24.0)
    def test_zero_logrado(self):
        alcance, puntos = _calc_alcance_puntos(0.0, 100.0, max_puntos=10)
        assert alcance == 0.0 and puntos == 0.0


class TestBuildVendorRecord:
    def test_returns_vendor_record(self):
        assert isinstance(_build_record(), VendorRecord)

    def test_ventas_netas_calculation(self):
        rec = _build_record(
            facturado=_make_facturado(total=100_000.0),
            devoluciones=_make_devoluciones(monto=10_000.0),
        )
        assert rec.ventas_netas == pytest.approx(90_000.0)

    def test_puntos_ventas_capped_at_30(self):
        rec = _build_record(
            budget=_make_budget(ventas=50_000.0),
            facturado=_make_facturado(total=200_000.0),
        )
        assert rec.puntos_ventas == pytest.approx(30.0)

    def test_puntos_cobranza_max_50(self):
        rec = _build_record(cobranza=_make_cobranza(sin_vencer=100_000.0, vencido=0.0))
        assert rec.puntos_cobranza == pytest.approx(50.0)

    def test_puntos_cobranza_proportional(self):
        rec = _build_record(cobranza=_make_cobranza(sin_vencer=80_000.0, vencido=20_000.0))
        assert rec.puntos_cobranza == pytest.approx(40.0)

    def test_total_indicadores_max_100(self):
        rec = _build_record(
            budget=_make_budget(ventas=50_000.0, activacion=10, skus=5),
            facturado=_make_facturado(total=50_000.0, clientes=10, skus=5),
            devoluciones=_make_devoluciones(monto=0.0),
            cobranza=_make_cobranza(sin_vencer=100_000.0, vencido=0.0),
        )
        assert rec.total_indicadores == pytest.approx(100.0)

    def test_incentivo_si_when_both_thresholds_met(self):
        rec = _build_record(
            budget=_make_budget(ventas=100_000.0),
            facturado=_make_facturado(total=90_000.0),
            devoluciones=_make_devoluciones(monto=0.0),
            cobranza=_make_cobranza(sin_vencer=90_000.0, vencido=10_000.0),
        )
        assert rec.incentivo == "SI"

    def test_incentivo_no_when_ventas_below_threshold(self):
        rec = _build_record(
            budget=_make_budget(ventas=100_000.0),
            facturado=_make_facturado(total=80_000.0),
            devoluciones=_make_devoluciones(monto=0.0),
            cobranza=_make_cobranza(sin_vencer=95_000.0, vencido=5_000.0),
        )
        assert rec.incentivo == "NO"

    def test_rifa_si_when_both_thresholds_met(self):
        rec = _build_record(
            budget=_make_budget(activacion=10, skus=10),
            facturado=_make_facturado(clientes=10, skus=10),
        )
        assert rec.rifa == "SI"

    def test_rifa_no_when_activacion_below_threshold(self):
        rec = _build_record(
            budget=_make_budget(activacion=10, skus=10),
            facturado=_make_facturado(clientes=8, skus=10),
        )
        assert rec.rifa == "NO"

    def test_to_row_length(self):
        row = _build_record().to_row()
        assert len(row) == len(config.GENERAL_HEADERS)

    def test_vendor_name_in_first_position(self):
        rec = _build_record(zona_info=_make_zona(nombre="JUAN PEREZ"))
        assert rec.to_row()[0] == "JUAN PEREZ"

    def test_zona_in_second_position(self):
        rec = _build_record(zona_info=_make_zona(zona="ZONA-99"))
        assert rec.to_row()[1] == "ZONA-99"


class TestCalculateRankings:
    def _make_records(self):
        r1 = _build_record(
            zona_info=_make_zona("MARIA"),
            budget=_make_budget(ventas=100_000),
            facturado=_make_facturado(total=100_000, clientes=50, skus=20),
            devoluciones=_make_devoluciones(0),
            cobranza=_make_cobranza(100_000, 0),
        )
        r2 = _build_record(
            zona_info=_make_zona("PEDRO"),
            budget=_make_budget(ventas=100_000),
            facturado=_make_facturado(total=70_000, clientes=30, skus=10),
            devoluciones=_make_devoluciones(0),
            cobranza=_make_cobranza(50_000, 50_000),
        )
        r3 = _build_record(
            zona_info=_make_zona("ANA"),
            budget=_make_budget(ventas=100_000),
            facturado=_make_facturado(total=50_000, clientes=20, skus=5),
            devoluciones=_make_devoluciones(0),
            cobranza=_make_cobranza(20_000, 80_000),
        )
        return [r1, r2, r3]

    def test_ranking_sorted_descending(self):
        records = calculate_rankings(self._make_records())
        totals = [r.total_indicadores for r in records]
        assert totals == sorted(totals, reverse=True)

    def test_lugar_indicadores_starts_at_1(self):
        records = calculate_rankings(self._make_records())
        assert records[0].lugar_indicadores == 1

    def test_lugar_indicadores_sequential(self):
        records = calculate_rankings(self._make_records())
        for i, rec in enumerate(records):
            assert rec.lugar_indicadores == i + 1

    def test_empty_list_returns_empty(self):
        assert calculate_rankings([]) == []


class TestFilterByRegion:
    def _make_mixed_records(self):
        cap = _build_record(zona_info=_make_zona(region="CAPITAL"))
        llanos = _build_record(zona_info=_make_zona(region="LLANOS"))
        return [cap, llanos]

    def test_filters_correctly(self):
        result = filter_by_region(self._make_mixed_records(), "CAPITAL")
        assert all(r.region == "CAPITAL" for r in result) and len(result) == 1

    def test_case_insensitive(self):
        result = filter_by_region(self._make_mixed_records(), "capital")
        assert len(result) == 1

    def test_nonexistent_region_returns_empty(self):
        assert filter_by_region(self._make_mixed_records(), "ORIENTE NORTE") == []


class TestGetRegionalTotals:
    def test_returns_dict(self):
        assert isinstance(get_regional_totals([_build_record()]), dict)

    def test_monto_pedidos_sums_correctly(self):
        r1 = _build_record(pedidos=_make_pedidos(monto=10_000))
        r2 = _build_record(pedidos=_make_pedidos(monto=20_000))
        totals = get_regional_totals([r1, r2])
        # Key may be "monto_pedidos" (internal) or "Monto Pedidos" (header name)
        total_val = totals.get("monto_pedidos") or totals.get("Monto Pedidos", 0)
        assert total_val == pytest.approx(30_000.0)

    def test_empty_list_returns_dict(self):
        assert isinstance(get_regional_totals([]), dict)


class TestApplyVendorAliases:
    def test_alias_replaces_odoo_name_with_canonical(self):
        raw_data = {
            "DOUGLAS GONZALEZ(FALCON-DABJ ZULIA)": "value_douglas",
            "MARIA NOGUERA(ZULIA-MCBO)": "value_maria",
            "VENDEDOR NORMAL": "value_normal",
        }
        res = _apply_vendor_aliases(raw_data)
        assert "DOUGLAS GONZALEZ(DABAJURO)" in res
        assert res["DOUGLAS GONZALEZ(DABAJURO)"] == "value_douglas"
        assert "MARIA NOGUERA(MARACAIBO)" in res
        assert res["MARIA NOGUERA(MARACAIBO)"] == "value_maria"
        assert "VENDEDOR NORMAL" in res
        assert res["VENDEDOR NORMAL"] == "value_normal"

    def test_alias_case_insensitive(self):
        raw_data = {
            "douglas gonzalez(falcon-dabj zulia)": 123,
        }
        res = _apply_vendor_aliases(raw_data)
        assert "DOUGLAS GONZALEZ(DABAJURO)" in res
        assert res["DOUGLAS GONZALEZ(DABAJURO)"] == 123

    def test_alias_oriana_to_jhonny_silva(self):
        raw_data = {
            "ORIANA DIAZ (CCS - LA GUAIRA)": 999.0,
        }
        res = _apply_vendor_aliases(raw_data)
        assert "JHONNY SILVA (CCS - LA GUAIRA)" in res
        assert res["JHONNY SILVA (CCS - LA GUAIRA)"] == 999.0


class TestCalculateGeneralNewVendors:
    def test_new_vendor_in_budget_is_included(self):
        zonas = {
            "EXISTING VENDOR": _make_zona(nombre="EXISTING VENDOR", zona="01", region="CAPITAL")
        }
        presupuestos = {
            "EXISTING VENDOR": _make_budget(ventas=10_000.0),
            "NEW VENDOR FUTURE": BudgetRecord(
                vendor_name="NEW VENDOR FUTURE",
                presup_ventas=15_000.0,
                presup_activacion=10.0,
                presup_skus=50.0,
                region="OCCIDENTE",
                zona="02",
            ),
        }
        records = calculate_general(
            zonas=zonas,
            facturado={},
            pedidos={},
            devoluciones={},
            cobranza={},
            por_facturar={},
            backorder={},
            presupuestos=presupuestos,
        )
        assert len(records) == 2
        names = [r.vendor_name for r in records]
        assert "EXISTING VENDOR" in names
        assert "NEW VENDOR FUTURE" in names

        new_rec = next(r for r in records if r.vendor_name == "NEW VENDOR FUTURE")
        assert new_rec.region == "OCCIDENTE"
        assert new_rec.presup_ventas == 15_000.0
        assert new_rec.zona == "02"


