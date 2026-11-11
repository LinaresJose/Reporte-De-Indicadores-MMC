"""tests/test_budget_loader.py — Tests for maintenance/budget_loader.py."""
from __future__ import annotations

import tempfile
from pathlib import Path
import pandas as pd
import pytest

from core.models import BudgetRecord
from maintenance.budget_loader import (
    sync_zonas_with_presupuestos,
    save_presupuestos_config,
    load_presupuestos,
)


class TestSyncZonasWithPresupuestos:
    def test_adds_new_vendor_to_zonas(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_zonas = Path(tmpdir) / "zonas.xlsx"
            df_init = pd.DataFrame([
                {
                    "Nombre": "VENDEDOR A",
                    "Zona": "100",
                    "Correo": "a@empresa.com",
                    "Codigo": "VMMC-01",
                    "Región": "CAPITAL",
                }
            ])
            df_init.to_excel(tmp_zonas, index=False)

            pres = {
                "VENDEDOR A": BudgetRecord(
                    vendor_name="VENDEDOR A",
                    presup_activacion=10,
                    presup_skus=50,
                    presup_ventas=10000,
                    region="CAPITAL",
                    zona="100",
                ),
                "NUEVO VENDEDOR": BudgetRecord(
                    vendor_name="NUEVO VENDEDOR",
                    presup_activacion=15,
                    presup_skus=80,
                    presup_ventas=20000,
                    region="OCCIDENTE",
                    zona="200",
                ),
            }

            added = sync_zonas_with_presupuestos(tmp_zonas, pres)
            assert added == ["NUEVO VENDEDOR"]

            df_updated = pd.read_excel(tmp_zonas)
            assert len(df_updated) == 2
            row_new = df_updated[df_updated["Nombre"] == "NUEVO VENDEDOR"].iloc[0]
            assert row_new["Zona"] == 200 or str(row_new["Zona"]) == "200"
            assert row_new["Región"] == "OCCIDENTE"
            assert row_new["Codigo"] == "VMMC-02"

    def test_does_not_duplicate_existing_vendors(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_zonas = Path(tmpdir) / "zonas.xlsx"
            df_init = pd.DataFrame([
                {
                    "Nombre": "VENDEDOR A",
                    "Zona": "100",
                    "Correo": "a@empresa.com",
                    "Codigo": "VMMC-01",
                    "Región": "CAPITAL",
                }
            ])
            df_init.to_excel(tmp_zonas, index=False)

            pres = {
                "VENDEDOR A": BudgetRecord(
                    vendor_name="VENDEDOR A",
                    presup_activacion=10,
                    presup_skus=50,
                    presup_ventas=10000,
                    region="CAPITAL",
                ),
            }

            added = sync_zonas_with_presupuestos(tmp_zonas, pres)
            assert added == []

            df_check = pd.read_excel(tmp_zonas)
            assert len(df_check) == 1
