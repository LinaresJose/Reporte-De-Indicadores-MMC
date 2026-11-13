"""tests/test_email_sender.py — Pruebas unitarias para el servicio de envío de correos."""

from pathlib import Path
import tempfile
import pandas as pd
import pytest

from services.email_sender import (
    get_default_templates,
    map_vendors_to_pdfs_and_emails,
    get_coordinator_dests,
    build_coordinator_extra_attachments,
)


def test_get_default_templates():
    subj_parcial, body_parcial = get_default_templates("Noviembre 2026", es_cierre=False)
    assert "Noviembre 2026" in subj_parcial
    assert "lo que va del mes" in subj_parcial.lower() or "noviembre" in subj_parcial.lower()
    assert "Estimado(a) Comercial" in body_parcial

    subj_cierre, body_cierre = get_default_templates("Noviembre 2026", es_cierre=True)
    assert "cierre" in subj_cierre.lower()
    assert "cierre definitivo" in body_cierre.lower()


def test_coordinator_extra_attachments():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        zonas_file = tmp_path / "zonas.xlsx"
        regionales_dir = tmp_path / "REGIONALES"
        regionales_dir.mkdir()

        # Mock zonas.xlsx with a coordinator
        df_zonas = pd.DataFrame([
            {"nombre": "COORDINADOR ORIENTE", "correo": "coord.oriente@mmc.com", "región": "ORIENTE"},
            {"nombre": "JUAN PEREZ", "correo": "juan@mmc.com", "región": "ORIENTE"},
        ])
        df_zonas.to_excel(zonas_file, index=False)

        # Create mock regional excel
        reg_excel = regionales_dir / "ORIENTE_Indicadores_Noviembre 2026.xlsx"
        reg_excel.write_text("dummy")

        extra_map = build_coordinator_extra_attachments(zonas_file, regionales_dir)
        assert "coord.oriente@mmc.com" in extra_map
        assert len(extra_map["coord.oriente@mmc.com"]) == 1
        assert extra_map["coord.oriente@mmc.com"][0].name == reg_excel.name
