"""tests/test_worker_directory.py — Pruebas unitarias para el directorio y modificación de correos."""

from pathlib import Path
import tempfile
import openpyxl
import pandas as pd
import pytest

from services.worker_directory import (
    get_workers_directory_data,
    export_workers_directory_excel,
    update_worker_email,
)


@pytest.fixture
def sample_data_paths():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        zonas_file = tmp_path / "zonas.xlsx"
        presupuestos_file = tmp_path / "presupuestos.xlsx"

        df_z = pd.DataFrame([
            {
                "Nombre": "CARLOS PEREZ (CARACAS)",
                "Zona": "101",
                "Correo": "carlos@mmc.com",
                "Codigo": "VMMC-01",
                "Región": "CAPITAL",
            },
            {
                "Nombre": "ANA LOPEZ (COORD CAPITAL)",
                "Zona": "190",
                "Correo": "ana.coord@mmc.com",
                "Codigo": "CMMC-01",
                "Región": "CAPITAL",
            },
            {
                "Nombre": "PEDRO GOMEZ (SIN CORREO)",
                "Zona": "102",
                "Correo": "",
                "Codigo": "VMMC-02",
                "Región": "CAPITAL",
            },
        ])
        df_z.to_excel(zonas_file, index=False)

        df_p = pd.DataFrame([
            {
                "Nombre": "CARLOS PEREZ (CARACAS)",
                "Zona": "101",
                "VENTAS": 50000.0,
                "ACTIVACIÓN": 25,
                "SKU": 250,
                "Region": "CAPITAL",
            },
            {
                "Nombre": "NUEVO VENDEDOR EN PRESUPUESTO",
                "Zona": "103",
                "VENTAS": 30000.0,
                "ACTIVACIÓN": 15,
                "SKU": 150,
                "Region": "ORIENTE",
            },
        ])
        df_p.to_excel(presupuestos_file, index=False)

        yield zonas_file, presupuestos_file


def test_get_workers_directory_data(sample_data_paths):
    zonas_file, presupuestos_file = sample_data_paths
    records = get_workers_directory_data(zonas_file, presupuestos_file)

    assert len(records) == 4
    names = [r["nombre"] for r in records]
    assert "CARLOS PEREZ (CARACAS)" in names
    assert "ANA LOPEZ (COORD CAPITAL)" in names
    assert "PEDRO GOMEZ (SIN CORREO)" in names
    assert "NUEVO VENDEDOR EN PRESUPUESTO" in names

    carlos = next(r for r in records if r["nombre"] == "CARLOS PEREZ (CARACAS)")
    assert carlos["tiene_correo"] is True
    assert carlos["correo"] == "carlos@mmc.com"
    assert carlos["meta_ventas"] == 50000.0

    pedro = next(r for r in records if r["nombre"] == "PEDRO GOMEZ (SIN CORREO)")
    assert pedro["tiene_correo"] is False
    assert "FALTA CORREO" in pedro["estado_correo"]


def test_export_workers_directory_excel(sample_data_paths):
    zonas_file, presupuestos_file = sample_data_paths
    output_excel = zonas_file.parent / "test_out_directorio.xlsx"

    out = export_workers_directory_excel(
        zonas_path=zonas_file,
        presupuestos_path=presupuestos_file,
        output_file=output_excel,
        mes_label="Noviembre 2026",
    )
    assert out.exists()

    wb = openpyxl.load_workbook(out, data_only=False)
    ws = wb["Directorio_Metas"]
    assert ws["A1"].value == "MATERIALES MODERNOS DE CONSTRUCCIÓN (MMC)"
    # Header is on row 6
    assert ws["B6"].value == "Nombre del Trabajador"
    assert ws["H6"].value == "Meta Ventas ($)"


def test_update_worker_email(sample_data_paths):
    zonas_file, _ = sample_data_paths

    # Update existing email
    ok, msg = update_worker_email(zonas_file, "PEDRO GOMEZ (SIN CORREO)", "pedro.nuevo@mmc.com")
    assert ok is True
    assert "actualizado" in msg.lower()

    df = pd.read_excel(zonas_file)
    row = df[df["Nombre"] == "PEDRO GOMEZ (SIN CORREO)"]
    assert row.iloc[0]["Correo"] == "pedro.nuevo@mmc.com"

    # Reject invalid email format
    ok_inv, msg_inv = update_worker_email(zonas_file, "CARLOS PEREZ (CARACAS)", "correo-sin-arroba")
    assert ok_inv is False
    assert "no es un correo electrónico válido" in msg_inv
