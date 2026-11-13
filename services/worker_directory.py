"""services/worker_directory.py — Gestión y descarga del Directorio de Trabajadores y Metas.

Permite:
1. Consolidar la información de todos los trabajadores (vendedores, coordinadores y gerencia),
   sus correos de destino, sus zonas, regiones y sus metas de presupuesto mensual.
2. Exportar un reporte profesional en Excel con formato corporativo MMC.
3. Actualizar y persistir el correo electrónico de cualquier trabajador en data/zonas.xlsx.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

import config

logger = logging.getLogger(__name__)

_COORD_KEYWORDS = ("COORD", "SUPERVISOR", "COORDINADOR")
_GERENCIA_KEYWORDS = ("GERENCIA", "GERENTE", "DIRECTOR", "ADMINISTRACION")

_CLR_HEADER_BG = "1F3864"   # Azul corporativo MMC
_CLR_HEADER_FG = "FFFFFF"   # Texto blanco
_CLR_ALT_ROW   = "F2F5F9"   # Gris/azul claro alterno
_CLR_WARN_BG   = "FCE4D6"   # Naranja suave para advertencias (sin correo)
_CLR_WARN_FG   = "C00000"   # Texto rojo advertencia
_CLR_OK_FG     = "006100"   # Verde oscuro para estados válidos
_CLR_OK_BG     = "C6EFCE"   # Verde claro fondo


def _classify_worker_role(nombre: str) -> str:
    """Clasifica el rol del trabajador en base a su nombre."""
    up = nombre.upper()
    if any(k in up for k in _GERENCIA_KEYWORDS):
        return "Gerencia / Dirección"
    if any(k in up for k in _COORD_KEYWORDS):
        return "Coordinador / Supervisor"
    return "Comercial / Vendedor"


def get_workers_directory_data(
    zonas_path: Path,
    presupuestos_path: Path,
) -> List[Dict[str, Any]]:
    """Genera la lista consolidada de trabajadores combinando zonas.xlsx y presupuestos.xlsx.

    Args:
        zonas_path: Ruta a data/zonas.xlsx
        presupuestos_path: Ruta a data/presupuestos.xlsx

    Returns:
        Lista de diccionarios ordenada por Región y Nombre con la información consolidada.
    """
    zonas_path = Path(zonas_path)
    presupuestos_path = Path(presupuestos_path)

    df_zonas = pd.DataFrame()
    if zonas_path.exists():
        try:
            df_zonas = pd.read_excel(zonas_path)
        except Exception as exc:
            logger.warning("Error leyendo zonas.xlsx: %s", exc)

    df_pres = pd.DataFrame()
    if presupuestos_path.exists():
        try:
            df_pres = pd.read_excel(presupuestos_path)
        except Exception as exc:
            logger.warning("Error leyendo presupuestos.xlsx: %s", exc)

    # Normalizar columnas de presupuestos
    pres_lookup: Dict[str, Dict[str, Any]] = {}
    if not df_pres.empty:
        col_map = {c: str(c).strip().upper() for c in df_pres.columns}
        df_p_norm = df_pres.rename(columns=col_map)
        for _, row in df_p_norm.iterrows():
            nom = str(row.get("NOMBRE", "") or "").strip()
            if not nom or nom.lower() == "nan":
                continue
            key = nom.upper()
            try:
                ventas = float(row.get("VENTAS", 0) or 0)
            except (ValueError, TypeError):
                ventas = 0.0
            try:
                activacion = float(row.get("ACTIVACIÓN", 0) or row.get("ACTIVACION", 0) or 0)
            except (ValueError, TypeError):
                activacion = 0.0
            try:
                skus = float(row.get("SKU", 0) or row.get("SKUS", 0) or 0)
            except (ValueError, TypeError):
                skus = 0.0
            region = str(row.get("REGION", "") or row.get("REGIÓN", "") or "").strip().upper()
            zona = str(row.get("ZONA", "") or "").strip()
            if zona.endswith(".0"):
                zona = zona[:-2]

            pres_lookup[key] = {
                "nombre_original": nom,
                "ventas": ventas,
                "activacion": activacion,
                "skus": skus,
                "region": region,
                "zona": zona,
            }

    records: List[Dict[str, Any]] = []
    seen_names: set[str] = set()

    # 1. Procesar trabajadores de zonas.xlsx
    if not df_zonas.empty:
        for _, row in df_zonas.iterrows():
            nombre = str(row.get("Nombre", "") or "").strip()
            if not nombre or nombre.lower() == "nan":
                continue

            upper_nombre = nombre.upper()
            seen_names.add(upper_nombre)

            zona = str(row.get("Zona", "") or "").strip()
            if zona.endswith(".0"):
                zona = zona[:-2]
            if zona.lower() == "nan":
                zona = ""

            correo = str(row.get("Correo", "") or "").strip()
            if correo.lower() == "nan":
                correo = ""

            codigo = str(row.get("Codigo", "") or row.get("Código", "") or "").strip()
            if codigo.lower() == "nan":
                codigo = ""

            region = str(row.get("Región", "") or row.get("Region", "") or "").strip().upper()
            if region.lower() == "nan":
                region = ""

            # Buscar metas en presupuestos (incluyendo alias)
            p_data = pres_lookup.get(upper_nombre)
            if not p_data:
                # Chequear alias inversos
                for alias_odoo, canonical in getattr(config, "VENDOR_ALIAS", {}).items():
                    if canonical.upper().strip() == upper_nombre:
                        p_data = pres_lookup.get(alias_odoo.upper().strip())
                        if p_data:
                            break

            meta_ventas = p_data["ventas"] if p_data else 0.0
            meta_activacion = p_data["activacion"] if p_data else 0.0
            meta_skus = p_data["skus"] if p_data else 0.0

            if not region and p_data and p_data.get("region"):
                region = p_data["region"]
            if not zona and p_data and p_data.get("zona"):
                zona = p_data["zona"]

            rol = _classify_worker_role(nombre)
            tiene_correo = bool(correo and "@" in correo)

            records.append({
                "nombre": nombre,
                "rol": rol,
                "zona": zona,
                "region": region or config.DEFAULT_REGION,
                "correo": correo,
                "tiene_correo": tiene_correo,
                "estado_correo": "Válido" if tiene_correo else "⚠ FALTA CORREO",
                "meta_ventas": meta_ventas,
                "meta_activacion": meta_activacion,
                "meta_skus": meta_skus,
                "codigo": codigo,
            })

    # 2. Agregar cualquier vendedor de presupuestos que no esté en zonas.xlsx
    for key, p_data in pres_lookup.items():
        canonical_target = getattr(config, "VENDOR_ALIAS", {}).get(p_data["nombre_original"]) or getattr(config, "VENDOR_ALIAS", {}).get(key)
        if key in seen_names or (canonical_target and canonical_target.upper().strip() in seen_names):
            continue

        v_name = canonical_target or p_data["nombre_original"]
        seen_names.add(key)
        seen_names.add(v_name.upper().strip())

        records.append({
            "nombre": v_name,
            "rol": _classify_worker_role(v_name),
            "zona": p_data["zona"],
            "region": p_data["region"] or config.DEFAULT_REGION,
            "correo": "",
            "tiene_correo": False,
            "estado_correo": "⚠ FALTA CORREO",
            "meta_ventas": p_data["ventas"],
            "meta_activacion": p_data["activacion"],
            "meta_skus": p_data["skus"],
            "codigo": "NUEVO",
        })

    # Ordenar por Región y luego por Nombre
    records.sort(key=lambda r: (r["region"], r["rol"], r["nombre"]))
    return records


def export_workers_directory_excel(
    zonas_path: Path,
    presupuestos_path: Path,
    output_file: Path,
    mes_label: str = "",
) -> Path:
    """Exporta el directorio de trabajadores y metas a un libro Excel corporativo.

    Args:
        zonas_path: Ruta a data/zonas.xlsx
        presupuestos_path: Ruta a data/presupuestos.xlsx
        output_file: Ruta destino donde se guardará el .xlsx
        mes_label: Etiqueta de mes opcional (ej: "Noviembre 2026")

    Returns:
        Ruta final del archivo creado.
    """
    records = get_workers_directory_data(zonas_path, presupuestos_path)
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Directorio_Metas"
    ws.views.sheetView[0].showGridLines = True

    # Estilos
    font_title = Font(name="Calibri", size=15, bold=True, color="1F3864")
    font_sub = Font(name="Calibri", size=10, italic=True, color="555555")
    font_header = Font(name="Calibri", size=11, bold=True, color=_CLR_HEADER_FG)
    font_data = Font(name="Calibri", size=10)
    font_data_bold = Font(name="Calibri", size=10, bold=True)
    font_warn = Font(name="Calibri", size=10, bold=True, color=_CLR_WARN_FG)
    font_ok = Font(name="Calibri", size=10, color=_CLR_OK_FG)

    fill_header = PatternFill(start_color=_CLR_HEADER_BG, end_color=_CLR_HEADER_BG, fill_type="solid")
    fill_alt = PatternFill(start_color=_CLR_ALT_ROW, end_color=_CLR_ALT_ROW, fill_type="solid")
    fill_warn = PatternFill(start_color=_CLR_WARN_BG, end_color=_CLR_WARN_BG, fill_type="solid")
    fill_ok = PatternFill(start_color=_CLR_OK_BG, end_color=_CLR_OK_BG, fill_type="solid")
    fill_totals = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    font_totals = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    # 1. Título y Banner
    ws["A1"] = "MATERIALES MODERNOS DE CONSTRUCCIÓN (MMC)"
    ws["A1"].font = font_title
    sub_text = f"Directorio de Trabajadores, Correos Asignados y Metas del Mes — {mes_label}" if mes_label else "Directorio de Trabajadores, Correos Asignados y Metas Mensuales"
    ws["A2"] = sub_text
    ws["A2"].font = font_sub

    # 2. Resumen rápido en cabecera
    total_trabajadores = len(records)
    con_correo = sum(1 for r in records if r["tiene_correo"])
    sin_correo = total_trabajadores - con_correo
    total_metas_ventas = sum(r["meta_ventas"] for r in records)

    ws["A4"] = f"Total Trabajadores: {total_trabajadores}   |   Con Correo Configurado: {con_correo}   |   Sin Correo: {sin_correo}   |   Presupuesto Total: ${total_metas_ventas:,.2f}"
    ws["A4"].font = Font(name="Calibri", size=10, bold=True, color="1F3864")

    # 3. Encabezados de Tabla
    headers = [
        ("N°", 6, "center"),
        ("Nombre del Trabajador", 36, "left"),
        ("Rol / Cargo", 24, "left"),
        ("Zona", 10, "center"),
        ("Región", 16, "left"),
        ("Correo Electrónico (Destino)", 36, "left"),
        ("Estado Correo", 16, "center"),
        ("Meta Ventas ($)", 18, "right"),
        ("Meta Activación", 16, "right"),
        ("Meta SKUs", 14, "right"),
        ("Código", 12, "center"),
    ]

    header_row = 6
    for col_idx, (title, width, align) in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=title)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[header_row].height = 28

    # 4. Filas de Datos
    curr_row = header_row + 1
    for idx, r in enumerate(records, start=1):
        row_fill = fill_alt if idx % 2 == 0 else PatternFill(fill_type=None)

        # N°
        c1 = ws.cell(row=curr_row, column=1, value=idx)
        c1.alignment = Alignment(horizontal="center")

        # Nombre
        c2 = ws.cell(row=curr_row, column=2, value=r["nombre"])
        c2.alignment = Alignment(horizontal="left")

        # Rol
        c3 = ws.cell(row=curr_row, column=3, value=r["rol"])
        c3.alignment = Alignment(horizontal="left")

        # Zona
        c4 = ws.cell(row=curr_row, column=4, value=r["zona"])
        c4.alignment = Alignment(horizontal="center")

        # Región
        c5 = ws.cell(row=curr_row, column=5, value=r["region"])
        c5.alignment = Alignment(horizontal="left")

        # Correo
        c6 = ws.cell(row=curr_row, column=6, value=r["correo"])
        c6.alignment = Alignment(horizontal="left")

        # Estado Correo
        c7 = ws.cell(row=curr_row, column=7, value=r["estado_correo"])
        c7.alignment = Alignment(horizontal="center")
        if not r["tiene_correo"]:
            c7.fill = fill_warn
            c7.font = font_warn
            c6.fill = fill_warn
        else:
            c7.font = font_ok

        # Meta Ventas
        c8 = ws.cell(row=curr_row, column=8, value=r["meta_ventas"])
        c8.number_format = "$#,##0.00"
        c8.alignment = Alignment(horizontal="right")

        # Meta Activación
        c9 = ws.cell(row=curr_row, column=9, value=r["meta_activacion"])
        c9.number_format = "#,##0"
        c9.alignment = Alignment(horizontal="right")

        # Meta SKUs
        c10 = ws.cell(row=curr_row, column=10, value=r["meta_skus"])
        c10.number_format = "#,##0"
        c10.alignment = Alignment(horizontal="right")

        # Código
        c11 = ws.cell(row=curr_row, column=11, value=r["codigo"])
        c11.alignment = Alignment(horizontal="center")

        # Aplicar bordes y fondo alterno si no tiene advertencia
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.border = thin_border
            if not cell.font or cell.font == font_data:
                cell.font = font_data
            if not cell.fill.fill_type and row_fill.fill_type:
                cell.fill = row_fill

        ws.row_dimensions[curr_row].height = 20
        curr_row += 1

    # 5. Fila de Totales
    totals_row = curr_row
    ws.merge_cells(start_row=totals_row, start_column=1, end_row=totals_row, end_column=7)
    tot_lbl = ws.cell(row=totals_row, column=1, value="TOTALES GENERALES")
    tot_lbl.font = font_totals
    tot_lbl.fill = fill_totals
    tot_lbl.alignment = Alignment(horizontal="center", vertical="center")

    # Suma Ventas
    tot_v = ws.cell(row=totals_row, column=8, value=f"=SUM(H{header_row+1}:H{totals_row-1})")
    tot_v.font = font_totals
    tot_v.fill = fill_totals
    tot_v.number_format = "$#,##0.00"
    tot_v.alignment = Alignment(horizontal="right", vertical="center")

    # Suma Activación
    tot_a = ws.cell(row=totals_row, column=9, value=f"=SUM(I{header_row+1}:I{totals_row-1})")
    tot_a.font = font_totals
    tot_a.fill = fill_totals
    tot_a.number_format = "#,##0"
    tot_a.alignment = Alignment(horizontal="right", vertical="center")

    # Suma SKUs
    tot_s = ws.cell(row=totals_row, column=10, value=f"=SUM(J{header_row+1}:J{totals_row-1})")
    tot_s.font = font_totals
    tot_s.fill = fill_totals
    tot_s.number_format = "#,##0"
    tot_s.alignment = Alignment(horizontal="right", vertical="center")

    tot_end = ws.cell(row=totals_row, column=11, value="")
    tot_end.fill = fill_totals

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=totals_row, column=col_idx)
        cell.border = thin_border
        cell.fill = fill_totals

    ws.row_dimensions[totals_row].height = 24

    # Congelar panel de encabezado
    ws.freeze_panes = f"A{header_row+1}"

    wb.save(output_file)
    logger.info("Directorio de trabajadores exportado exitosamente a: %s", output_file)
    return output_file


def update_worker_email(
    zonas_path: Path,
    worker_name: str,
    new_email: str,
) -> Tuple[bool, str]:
    """Actualiza de forma persistente el correo de un trabajador en data/zonas.xlsx.

    Args:
        zonas_path: Ruta al archivo data/zonas.xlsx
        worker_name: Nombre exacto del trabajador a actualizar
        new_email: Nueva dirección de correo (o vacío)

    Returns:
        Tupla (éxito: bool, mensaje: str).
    """
    zonas_path = Path(zonas_path)
    if not zonas_path.exists():
        return False, f"El archivo {zonas_path} no existe."

    cleaned_email = new_email.strip()
    if cleaned_email and ("@" not in cleaned_email or "." not in cleaned_email):
        return False, f"'{cleaned_email}' no es un correo electrónico válido."

    try:
        df = pd.read_excel(zonas_path)
    except Exception as exc:
        return False, f"No se pudo abrir {zonas_path.name}: {exc}"

    if config.COL_ZONA_NOMBRE not in df.columns or config.COL_ZONA_CORREO not in df.columns:
        return False, f"Columnas requeridas '{config.COL_ZONA_NOMBRE}' o '{config.COL_ZONA_CORREO}' no encontradas."

    target_upper = worker_name.strip().upper()
    match_idx = None

    for idx, val in df[config.COL_ZONA_NOMBRE].items():
        if str(val).strip().upper() == target_upper:
            match_idx = idx
            break

    if match_idx is None:
        # Intentar por alias
        canonical = getattr(config, "VENDOR_ALIAS", {}).get(worker_name) or getattr(config, "VENDOR_ALIAS", {}).get(target_upper)
        if canonical:
            can_upper = canonical.strip().upper()
            for idx, val in df[config.COL_ZONA_NOMBRE].items():
                if str(val).strip().upper() == can_upper:
                    match_idx = idx
                    break

    if match_idx is not None:
        actual_name = df.at[match_idx, config.COL_ZONA_NOMBRE]
        old_email = str(df.at[match_idx, config.COL_ZONA_CORREO] or "").strip()
        df.at[match_idx, config.COL_ZONA_CORREO] = cleaned_email
        try:
            df.to_excel(zonas_path, index=False, engine="openpyxl")
            logger.info("Correo actualizado para '%s': de '%s' a '%s'", actual_name, old_email, cleaned_email)
            return True, f"Correo de '{actual_name}' actualizado exitosamente a: {cleaned_email or '(Sin correo)'}"
        except Exception as exc:
            return False, f"Error al guardar cambios en {zonas_path.name}: {exc}"
    else:
        # El trabajador no estaba en zonas.xlsx, agregarlo
        new_row = {
            config.COL_ZONA_NOMBRE: worker_name.strip(),
            config.COL_ZONA_ZONA: "",
            config.COL_ZONA_CORREO: cleaned_email,
            config.COL_ZONA_CODIGO: f"VMMC-{len(df)+1:02d}",
            config.COL_ZONA_REGION: config.DEFAULT_REGION,
        }
        df_updated = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        try:
            df_updated.to_excel(zonas_path, index=False, engine="openpyxl")
            logger.info("Nuevo trabajador '%s' registrado en zonas con correo '%s'", worker_name, cleaned_email)
            return True, f"Trabajador '{worker_name}' registrado exitosamente con correo: {cleaned_email}"
        except Exception as exc:
            return False, f"Error al guardar nuevo trabajador en {zonas_path.name}: {exc}"
