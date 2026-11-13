"""
Business logic engine for the LYNX sales KPI system.

This module transforms the aggregated ETL data into the final 34-column
``VendorRecord`` objects, computes rankings, filters by region, and calculates
regional totals for the Â«totals rowÂ» in each sheet.
"""
from __future__ import annotations

import logging
from typing import Dict, List, Optional

import config
from core.exceptions import EngineError
from core.models import (
    BackorderAgg,
    BudgetRecord,
    CobranzaAgg,
    DevolucionesAgg,
    FacturadoAgg,
    PedidosAgg,
    PorFacturarAgg,
    VendorRecord,
    ZonaInfo,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _apply_vendor_aliases(data: dict) -> dict:
    """Return a new dict with Odoo vendor names replaced by their canonical
    aliases as defined in :data:`config.VENDOR_ALIAS`.

    The comparison is done case-insensitively so minor casing differences in
    the Odoo export do not break the mapping.  Keys that are not listed in the
    alias map are kept unchanged.

    Args:
        data: Any ETL result dict keyed by vendor name string.

    Returns:
        New dict with aliased keys where applicable; original entries untouched.
    """
    alias_upper: dict[str, str] = {
        k.upper().strip(): v for k, v in config.VENDOR_ALIAS.items()
    }
    result: dict = {}
    for key, value in data.items():
        canonical = alias_upper.get(key.upper().strip(), key)
        result[canonical] = value
    return result


def _safe_divide(numerator: float, denominator: float) -> float:
    """Safely divide two floats, returning 0.0 when the denominator is zero.

    Args:
        numerator: The dividend.
        denominator: The divisor.

    Returns:
        ``numerator / denominator`` or ``0.0`` if ``denominator == 0``.
    """
    if denominator == 0.0:
        return 0.0
    try:
        return numerator / denominator
    except ZeroDivisionError:
        return 0.0


def _calc_alcance_puntos(
    actual: float,
    presup: float,
    max_puntos: float,
) -> tuple[float, float]:
    """Compute the KPI achievement ratio and capped score.

    Args:
        actual: Actual value achieved (e.g. ventas netas).
        presup: Budget / target value.
        max_puntos: Maximum points that can be awarded for this KPI.

    Returns:
        A ``(alcance, puntos)`` tuple where ``alcance = actual / presup``
        (capped at 0 if presup is 0) and ``puntos = min(max_puntos, max_puntos * alcance)``.
    """
    alcance = _safe_divide(actual, presup)
    puntos = min(max_puntos, max_puntos * alcance)
    return alcance, puntos


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------


def build_vendor_record(
    zona_info: ZonaInfo,
    budget: Optional[BudgetRecord],
    facturado: Optional[FacturadoAgg],
    pedidos: Optional[PedidosAgg],
    devoluciones: Optional[DevolucionesAgg],
    cobranza: Optional[CobranzaAgg],
    por_facturar: Optional[PorFacturarAgg],
    backorder: Optional[BackorderAgg],
) -> VendorRecord:
    """Build a fully-computed :class:`VendorRecord` for a single vendor.

    Applies all LYNX business rules: KPI scores, incentive eligibility, and
    rifa qualification.  Ranking fields (``lugar_*``) and ``participacion``
    are set to their zero defaults here and are populated by
    :func:`calculate_general` and :func:`calculate_rankings` respectively.

    Args:
        zona_info: Geographic and contact information for the vendor.
        budget: Monthly targets; use ``None`` if the vendor has no budget for
            this period (all presup fields will be 0).
        facturado: Aggregated invoiced-sales data; ``None`` â†’ zeros.
        pedidos: Aggregated orders data; ``None`` â†’ zeros.
        devoluciones: Aggregated returns data; ``None`` â†’ zeros.
        cobranza: Aggregated accounts-receivable data; ``None`` â†’ zeros.
        por_facturar: Aggregated pending-invoicing data; ``None`` â†’ zeros.
        backorder: Aggregated backorder data; ``None`` â†’ zeros.

    Returns:
        A populated :class:`VendorRecord`.
    """
    # â”€â”€ Budget defaults â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    presup_ventas = budget.presup_ventas if budget else 0.0
    presup_activacion = budget.presup_activacion if budget else 0.0
    presup_skus = budget.presup_skus if budget else 0.0

    # â”€â”€ Pedidos â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    cantidad_pedidos = pedidos.cantidad_pedidos if pedidos else 0
    monto_pedidos = pedidos.monto_pedidos if pedidos else 0.0

    # â”€â”€ Facturado y devoluciones â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    total_facturado = facturado.total_facturado if facturado else 0.0
    clientes_activados = facturado.clientes_activados if facturado else 0
    skus_vendidos = facturado.skus_vendidos if facturado else 0
    total_devoluciones = devoluciones.total_devoluciones if devoluciones else 0.0
    ventas_netas = total_facturado - total_devoluciones

    # â”€â”€ LogÃ­stica â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    monto_por_procesar = por_facturar.total_base_imponible if por_facturar else 0.0
    backorder_val = backorder.backorder_value if backorder else 0.0

    # â”€â”€ Ventas KPI â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    diferencia_ventas = abs(ventas_netas - presup_ventas)
    alcance_ventas, puntos_ventas = _calc_alcance_puntos(
        ventas_netas, presup_ventas, config.MAX_PUNTOS_VENTAS
    )

    # â”€â”€ ActivaciÃ³n KPI â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    diferencia_activacion = abs(clientes_activados - presup_activacion)
    alcance_activacion, puntos_activacion = _calc_alcance_puntos(
        float(clientes_activados), presup_activacion, config.MAX_PUNTOS_ACTIVACION
    )

    # â”€â”€ SKUs KPI â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    diferencia_skus = abs(skus_vendidos - presup_skus)
    alcance_skus, puntos_skus = _calc_alcance_puntos(
        float(skus_vendidos), presup_skus, config.MAX_PUNTOS_SKUS
    )

    # â”€â”€ Cobranza KPI â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # CobranzaAgg stores raw totals; pct_sin_vencer is computed here
    if cobranza:
        total_cxc = cobranza.total_sin_vencer + cobranza.total_vencido
        pct_sin_vencer = _safe_divide(cobranza.total_sin_vencer, total_cxc)
    else:
        pct_sin_vencer = 0.0
    puntos_cobranza = config.MAX_PUNTOS_COBRANZA * pct_sin_vencer

    # â”€â”€ Totals and bonuses â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    total_indicadores = puntos_ventas + puntos_activacion + puntos_skus + puntos_cobranza

    incentivo = (
        "SI"
        if alcance_ventas >= config.THRESHOLD_INCENTIVO
        and pct_sin_vencer >= config.THRESHOLD_INCENTIVO
        else "NO"
    )

    rifa = (
        "SI"
        if alcance_activacion >= config.THRESHOLD_RIFA
        and alcance_skus >= config.THRESHOLD_RIFA
        else "NO"
    )

    return VendorRecord(
        vendor_name=zona_info.vendor_name,
        zona=zona_info.zona,
        region=zona_info.region,
        cantidad_pedidos=cantidad_pedidos,
        monto_pedidos=monto_pedidos,
        devoluciones=total_devoluciones,
        total_facturado=total_facturado,
        ventas_netas=ventas_netas,
        monto_por_procesar=monto_por_procesar,
        backorder=backorder_val,
        presup_ventas=presup_ventas,
        diferencia_ventas=diferencia_ventas,
        alcance_ventas=alcance_ventas,
        puntos_ventas=puntos_ventas,
        clientes_activados=clientes_activados,
        presup_activacion=presup_activacion,
        diferencia_activacion=diferencia_activacion,
        alcance_activacion=alcance_activacion,
        puntos_activacion=puntos_activacion,
        skus_vendidos=skus_vendidos,
        presup_skus=presup_skus,
        diferencia_skus=diferencia_skus,
        alcance_skus=alcance_skus,
        puntos_skus=puntos_skus,
        pct_sin_vencer=pct_sin_vencer,
        puntos_cobranza=puntos_cobranza,
        total_indicadores=total_indicadores,
        incentivo=incentivo,
        rifa=rifa,
        # Rankings and participaciÃ³n filled in later
    )


def calculate_general(
    zonas: Dict[str, ZonaInfo],
    facturado: Dict[str, FacturadoAgg],
    pedidos: Dict[str, PedidosAgg],
    devoluciones: Dict[str, DevolucionesAgg],
    cobranza: Dict[str, CobranzaAgg],
    por_facturar: Dict[str, PorFacturarAgg],
    backorder: Dict[str, BackorderAgg],
    presupuestos: Dict[str, BudgetRecord],
    cobranza_override: Optional[Dict[str, CobranzaAgg]] = None,
) -> List[VendorRecord]:
    """Build all vendor records and compute rankings and participación."""
    records: List[VendorRecord] = []

    # Apply vendor alias map so Odoo names are resolved to canonical zonas names
    facturado   = _apply_vendor_aliases(facturado)
    pedidos     = _apply_vendor_aliases(pedidos)
    devoluciones = _apply_vendor_aliases(devoluciones)
    cobranza    = _apply_vendor_aliases(cobranza)
    por_facturar = _apply_vendor_aliases(por_facturar)
    backorder   = _apply_vendor_aliases(backorder)

    fac_lookup = {k.upper().strip(): v for k, v in facturado.items()}
    ped_lookup = {k.upper().strip(): v for k, v in pedidos.items()}
    dev_lookup = {k.upper().strip(): v for k, v in devoluciones.items()}
    cxc_lookup = {k.upper().strip(): v for k, v in cobranza.items()}
    pf_lookup = {k.upper().strip(): v for k, v in por_facturar.items()}
    bo_lookup = {k.upper().strip(): v for k, v in backorder.items()}
    pres_lookup = {k.upper().strip(): v for k, v in presupuestos.items()}
    cob_override_lookup = {k.upper().strip(): v for k, v in cobranza_override.items()} if cobranza_override else {}

    processed_keys: set[str] = set()

    for name, zona_info in zonas.items():
        key = name.upper().strip()
        processed_keys.add(key)
        cob_entry = cob_override_lookup.get(key) or cxc_lookup.get(key)

        try:
            record = build_vendor_record(
                zona_info=zona_info,
                budget=pres_lookup.get(key),
                facturado=fac_lookup.get(key),
                pedidos=ped_lookup.get(key),
                devoluciones=dev_lookup.get(key),
                cobranza=cob_entry,
                por_facturar=pf_lookup.get(key),
                backorder=bo_lookup.get(key),
            )
            records.append(record)
        except Exception as exc:
            raise EngineError(f"Failed to build record for vendor '{name}': {exc}") from exc

    # Automatically incorporate any new vendor present in presupuestos that is not yet in zonas
    for p_key, budget_rec in pres_lookup.items():
        if p_key in processed_keys:
            continue

        # Check if this budget entry is an alias of an already processed vendor
        alias_target = (
            config.VENDOR_ALIAS.get(budget_rec.vendor_name)
            or config.VENDOR_ALIAS.get(p_key)
        )
        if alias_target and alias_target.upper().strip() in processed_keys:
            continue

        vendor_name = budget_rec.vendor_name
        region_str = (budget_rec.region or config.DEFAULT_REGION).upper().strip()
        if region_str not in config.REGIONES:
            region_str = config.DEFAULT_REGION

        synthetic_zona = ZonaInfo(
            vendor_name=vendor_name,
            zona=getattr(budget_rec, "zona", "") or "",
            region=region_str,
            correo="",
            codigo="",
        )
        cob_entry = cob_override_lookup.get(p_key) or cxc_lookup.get(p_key)

        try:
            record = build_vendor_record(
                zona_info=synthetic_zona,
                budget=budget_rec,
                facturado=fac_lookup.get(p_key),
                pedidos=ped_lookup.get(p_key),
                devoluciones=dev_lookup.get(p_key),
                cobranza=cob_entry,
                por_facturar=pf_lookup.get(p_key),
                backorder=bo_lookup.get(p_key),
            )
            records.append(record)
            processed_keys.add(p_key)
            logger.info(
                "calculate_general: Nuevo vendedor en presupuestos incluido automáticamente: '%s' (%s)",
                vendor_name,
                region_str,
            )
        except Exception as exc:
            raise EngineError(
                f"Failed to build record for new budget vendor '{vendor_name}': {exc}"
            ) from exc

    # Check if there are active sellers in Odoo not registered in zonas/presupuestos
    unrecognized = [
        k for k in fac_lookup.keys()
        if k not in processed_keys and k not in {"DIRECTOR", "ADMIN", "OFICINA", "GERENCIA", "DEFAULT", "NAN", ""}
    ]
    if unrecognized:
        logger.warning(
            "calculate_general: Vendedores con ventas en Odoo NO registrados en zonas ni presupuestos: %s",
            unrecognized,
        )

    total_ventas_global = sum(r.ventas_netas for r in records)
    for rec in records:
        rec.participacion = _safe_divide(rec.ventas_netas, total_ventas_global)

    # â”€â”€ Post-process: rankings â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    calculate_rankings(records)

    logger.info(
        "calculate_general: %d records | total_ventas_global=$%.2f",
        len(records),
        total_ventas_global,
    )
    return records


def calculate_rankings(records: List[VendorRecord]) -> List[VendorRecord]:
    """Assign ranking positions to vendors who have an active budget.

    Vendors with ``presup_ventas == 0`` AND ``presup_activacion == 0``
    AND ``presup_skus == 0`` are considered unranked and keep their
    ``lugar_*`` fields at ``0``.

    Rankings are assigned across four dimensions:
    - **Lugar Ventas**: by ``alcance_ventas`` descending.
    - **Lugar ActivaciÃ³n**: by ``alcance_activacion`` descending.
    - **Lugar SKUs**: by ``alcance_skus`` descending.
    - **Lugar Indicadores**: by ``total_indicadores`` descending.

    Args:
        records: List of :class:`VendorRecord` objects (mutated in place).

    Returns:
        The same *records* list (mutation is in place; return value aids chaining).
    """
    ranked = [
        r for r in records
        if r.presup_ventas > 0 or r.presup_activacion > 0 or r.presup_skus > 0
    ]

    dimensions = [
        ("alcance_ventas", "lugar_ventas"),
        ("alcance_activacion", "lugar_activacion"),
        ("alcance_skus", "lugar_skus"),
        ("total_indicadores", "lugar_indicadores"),
    ]

    for metric_attr, rank_attr in dimensions:
        sorted_ranked = sorted(ranked, key=lambda r: getattr(r, metric_attr), reverse=True)
        for position, rec in enumerate(sorted_ranked, start=1):
            setattr(rec, rank_attr, position)

    logger.info("calculate_rankings: %d vendors ranked.", len(ranked))
    return records


def filter_by_region(
    records: List[VendorRecord],
    region: str,
) -> List[VendorRecord]:
    """Return records belonging to *region*, sorted by total score descending.

    Args:
        records: Full list of :class:`VendorRecord` objects.
        region: Region name to filter by (case-insensitive, e.g. ``"CAPITAL"``).

    Returns:
        Filtered and sorted sub-list; empty list if no records match.
    """
    target = region.upper().strip()
    filtered = [r for r in records if r.region.upper().strip() == target]
    return sorted(filtered, key=lambda r: r.total_indicadores, reverse=True)


def get_regional_totals(records: List[VendorRecord]) -> dict:
    """Compute aggregated totals for the totals row of a sheet.

    Monetary fields are summed; achievement ratios and scores are computed
    from those sums (not averaged from individual ratios) to ensure
    mathematical correctness.

    Args:
        records: List of :class:`VendorRecord` objects for a sheet (can be
            the full general list or a regional sub-list).

    Returns:
        A dictionary of computed totals ready for :func:`reports.excel_writer.write_totals_row`.
    """
    if not records:
        return {}

    sum_ventas_netas = sum(r.ventas_netas for r in records)
    sum_presup_ventas = sum(r.presup_ventas for r in records)
    sum_facturado = sum(r.total_facturado for r in records)
    sum_devoluciones = sum(r.devoluciones for r in records)
    sum_pedidos_monto = sum(r.monto_pedidos for r in records)
    sum_cantidad_pedidos = sum(r.cantidad_pedidos for r in records)
    sum_por_procesar = sum(r.monto_por_procesar for r in records)
    sum_backorder = sum(r.backorder for r in records)
    sum_clientes = sum(r.clientes_activados for r in records)
    sum_presup_activacion = sum(r.presup_activacion for r in records)
    sum_skus = sum(r.skus_vendidos for r in records)
    sum_presup_skus = sum(r.presup_skus for r in records)
    sum_participacion = sum(r.participacion for r in records)

    # Weighted aggregate achievements
    alcance_ventas_total = _safe_divide(sum_ventas_netas, sum_presup_ventas)
    alcance_activacion_total = _safe_divide(float(sum_clientes), sum_presup_activacion)
    alcance_skus_total = _safe_divide(float(sum_skus), sum_presup_skus)

    puntos_ventas_total = min(config.MAX_PUNTOS_VENTAS, config.MAX_PUNTOS_VENTAS * alcance_ventas_total)
    puntos_activacion_total = min(config.MAX_PUNTOS_ACTIVACION, config.MAX_PUNTOS_ACTIVACION * alcance_activacion_total)
    puntos_skus_total = min(config.MAX_PUNTOS_SKUS, config.MAX_PUNTOS_SKUS * alcance_skus_total)

    # Cobranza: simple average of % sin vencer for vendors with any cobranza data
    cob_records = [r for r in records if (r.pct_sin_vencer > 0 or r.puntos_cobranza > 0)]
    pct_sin_vencer_total = (
        sum(r.pct_sin_vencer for r in cob_records) / len(cob_records)
        if cob_records else 0.0
    )
    puntos_cobranza_total = config.MAX_PUNTOS_COBRANZA * pct_sin_vencer_total
    total_indicadores_total = (
        puntos_ventas_total + puntos_activacion_total
        + puntos_skus_total + puntos_cobranza_total
    )

    return {
        "zona_count": len(records),
        "cantidad_pedidos": sum_cantidad_pedidos,
        "monto_pedidos": sum_pedidos_monto,
        "devoluciones": sum_devoluciones,
        "total_facturado": sum_facturado,
        "ventas_netas": sum_ventas_netas,
        "monto_por_procesar": sum_por_procesar,
        "backorder": sum_backorder,
        "presup_ventas": sum_presup_ventas,
        "diferencia_ventas": abs(sum_ventas_netas - sum_presup_ventas),
        "alcance_ventas": alcance_ventas_total,
        "puntos_ventas": puntos_ventas_total,
        "clientes_activados": sum_clientes,
        "presup_activacion": sum_presup_activacion,
        "diferencia_activacion": abs(float(sum_clientes) - sum_presup_activacion),
        "alcance_activacion": alcance_activacion_total,
        "puntos_activacion": puntos_activacion_total,
        "skus_vendidos": sum_skus,
        "presup_skus": sum_presup_skus,
        "diferencia_skus": abs(float(sum_skus) - sum_presup_skus),
        "alcance_skus": alcance_skus_total,
        "puntos_skus": puntos_skus_total,
        "pct_sin_vencer": pct_sin_vencer_total,
        "puntos_cobranza": puntos_cobranza_total,
        "total_indicadores": total_indicadores_total,
        "participacion": sum_participacion,
        "incentivo_count": sum(1 for r in records if r.incentivo == "SI"),
        "rifa_count": sum(1 for r in records if r.rifa == "SI"),
    }

