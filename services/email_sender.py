"""services/email_sender.py — Servicio de envío de reportes PDF por correo (Gmail/SMTP).

Permite enviar a cada comercial su reporte PDF individual con asunto y cuerpo
personalizados según sea informe parcial (durante el mes) o cierre de mes.

Coordinators (COORD / SUPERVISOR) automatically receive their regional Excel
file as an extra attachment on top of their individual PDF (if one exists).
"""
from __future__ import annotations

import logging
import mimetypes
import os
import smtplib
import ssl
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

DEFAULT_SENDER_EMAIL = "sistemas3.mmcve@gmail.com"
DEFAULT_SENDER_NAME = "Materiales Modernos de Construcción"
GMAIL_SMTP_SERVER = "smtp.gmail.com"
GMAIL_SMTP_PORT = 587  # STARTTLS


def get_default_templates(mes_nombre: str, es_cierre: bool = False) -> Tuple[str, str]:
    """Genera el asunto y cuerpo de correo predeterminado.

    Args:
        mes_nombre: Nombre del mes (e.g. 'Noviembre' o 'Noviembre 2026').
        es_cierre: True si es cierre definitivo de mes, False si es reporte parcial.

    Returns:
        Tupla con (asunto, cuerpo_html).
    """
    if es_cierre:
        asunto = f"Reporte de Indicadores de Ventas al cierre de {mes_nombre}"
        detalle_tiempo = f"al cierre definitivo del mes de <strong>{mes_nombre}</strong>"
    else:
        asunto = f"Reporte de Indicadores de Ventas de lo que va del mes en curso - {mes_nombre}"
        detalle_tiempo = f"de lo que va del mes en curso (<strong>{mes_nombre}</strong>)"

    cuerpo = f"""\
Estimado(a) Comercial,

Adjunto encontrará su Reporte de Indicadores de Ventas {detalle_tiempo}.

Por favor revise los indicadores correspondientes a su gestión (Ventas netas, Activación de clientes, SKUs y Cobranza).

Cualquier duda o aclaratoria, favor comunicarse con la Gerencia Comercial.

Atentamente,
Sistemas MMC
Materiales Modernos de Construcción
"""
    return asunto, cuerpo


def map_vendors_to_pdfs_and_emails(
    manifest_csv: Path,
    zonas_file: Path,
    pdf_dir: Path,
) -> List[Dict[str, str]]:
    """Empareja a cada comercial con su correo y su archivo PDF individual.

    Args:
        manifest_csv: Ruta al CSV generado de vendedores.
        zonas_file: Ruta al archivo data/zonas.xlsx.
        pdf_dir: Directorio donde están los PDFs (output/PDF_INDICADORES).

    Returns:
        Lista de diccionarios con keys: 'nombre', 'zona', 'region', 'email', 'pdf_path', 'existe_pdf'.
    """
    manifest_csv = Path(manifest_csv)
    zonas_file = Path(zonas_file)
    pdf_dir = Path(pdf_dir)

    # 1. Leer zonas para correos
    df_zonas = pd.read_excel(zonas_file)
    email_map: Dict[str, str] = {}
    for _, row in df_zonas.iterrows():
        nombre = str(row.get("Nombre", "")).strip().upper()
        correo = str(row.get("Correo", "")).strip()
        if correo and correo.lower() != "nan" and "@" in correo:
            email_map[nombre] = correo

    # 2. Leer manifiesto
    df_man = pd.read_csv(manifest_csv, sep=";")
    destinatarios: List[Dict[str, str]] = []

    for _, row in df_man.iterrows():
        vendedor = str(row.get("VendedorNombre", "")).strip()
        key = vendedor.upper()
        zona = str(row.get("Zona", "")).strip()
        region = str(row.get("Region", "")).strip()
        rel_pdf = str(row.get("NombreArchivo", "")).strip()

        # Ruta del PDF
        pdf_filename = Path(rel_pdf).name
        full_pdf_path = pdf_dir / pdf_filename
        if not full_pdf_path.exists():
            # Intentar ruta directa si venía relativa
            candidate = manifest_csv.parent / rel_pdf
            if candidate.exists():
                full_pdf_path = candidate

        correo = email_map.get(key, "")

        destinatarios.append({
            "nombre": vendedor,
            "zona": zona,
            "region": region,
            "email": correo,
            "pdf_path": str(full_pdf_path),
            "existe_pdf": str(full_pdf_path.exists()),
        })

    return destinatarios


# ---------------------------------------------------------------------------
# Coordinator helpers
# ---------------------------------------------------------------------------

#: Keywords that identify a coordinator/supervisor row in zonas.xlsx.
_COORD_KEYWORDS = ("COORD", "SUPERVISOR")

#: Emails of the company owners / gerencia — they receive the General Excel.
OWNER_EMAILS: frozenset[str] = frozenset({
    "hhurtado.mmcve@gmail.com",
    "jatouman.mmcve@gmail.com",
})


def _region_to_filename_prefix(region: str) -> str:
    """Convert a region name to the filename prefix used by generate_excel_report.

    Examples:
        "OCCIDENTE"    → "OCCIDENTE"
        "CENTRO COSTA" → "CENTRO_COSTA"
        "ORIENTE NORTE"→ "ORIENTE_NORTE"
    """
    return region.strip().upper().replace(" ", "_")


def build_coordinator_extra_attachments(
    zonas_file: Path,
    regional_dir: Path,
) -> Dict[str, List[Path]]:
    """Build a mapping of coordinator email → list of extra file paths to attach.

    Scans ``zonas.xlsx`` for rows whose *Nombre* contains 'COORD' or 'SUPERVISOR'
    (excluding GERENCIA owners who already get the General Excel).  For each such
    coordinator, looks for a matching regional Excel file inside *regional_dir*.

    The regional file name pattern is::

        <REGION_PREFIX>_Indicadores_<Mes>_<Año>.xlsx

    Args:
        zonas_file: Path to ``data/zonas.xlsx``.
        regional_dir: Path to the ``REGIONALES/`` subfolder inside the output dir.

    Returns:
        Dict mapping lowercase email → list of :class:`~pathlib.Path` objects.
        Only coordinators whose regional file actually exists are included.
    """
    result: Dict[str, List[Path]] = {}

    if not zonas_file.exists() or not regional_dir.exists():
        logger.warning(
            "build_coordinator_extra_attachments: missing path(s) — zonas=%s regional=%s",
            zonas_file, regional_dir,
        )
        return result

    df = pd.read_excel(zonas_file)

    # Normalise column names (handle encoding quirks like 'Región' vs 'Regi?n')
    col_map = {}
    for col in df.columns:
        normalized = col.strip().lower().replace("\ufffd", "").replace("?", "")
        col_map[col] = normalized
    df = df.rename(columns=col_map)

    for _, row in df.iterrows():
        nombre = str(row.get("nombre", "") or "").strip().upper()
        correo = str(row.get("correo", "") or "").strip()
        region = str(row.get("región", "") or row.get("regin", "") or row.get("region", "") or "").strip().upper()

        # Skip rows that are not coordinator/supervisor, or lack email
        is_coord = any(kw in nombre for kw in _COORD_KEYWORDS)
        is_gerencia = "GERENCIA" in nombre
        if not is_coord or is_gerencia or not correo or "@" not in correo:
            continue

        if not region:
            logger.warning("Coordinator '%s' has no region in zonas.xlsx — skipping", nombre)
            continue

        prefix = _region_to_filename_prefix(region)
        # Find matching file (glob to avoid hardcoding Mes/Año)
        candidates = list(regional_dir.glob(f"{prefix}_Indicadores_*.xlsx"))
        if not candidates:
            logger.warning(
                "No regional Excel found for coordinator '%s' (region=%s, prefix=%s) in %s",
                nombre, region, prefix, regional_dir,
            )
            continue

        # Pick most recently modified if multiple matches
        excel_path = max(candidates, key=lambda p: p.stat().st_mtime)
        result.setdefault(correo.lower(), []).append(excel_path)
        logger.info(
            "Coordinator '%s' (%s) → regional Excel: %s", nombre, correo, excel_path.name
        )

    return result


def get_coordinator_dests(
    zonas_file: Path,
    regional_dir: Path,
    pdf_dir: Path,
) -> List[Dict[str, str]]:
    """Build destination entries for coordinators who are NOT in the vendor manifest.

    Coordinators typically have no (or zero) budget and are therefore absent from
    the vendor manifest CSV.  This function reads ``zonas.xlsx`` directly to find
    them and pairs them with their regional Excel file (as the attachment) rather
    than an individual PDF.

    Args:
        zonas_file: Path to ``data/zonas.xlsx``.
        regional_dir: Path to the ``REGIONALES/`` subfolder inside the output dir.
        pdf_dir: Path to the ``PDF_INDICADORES/`` subfolder (to check for a PDF).

    Returns:
        List of destination dicts with the same keys as
        :func:`map_vendors_to_pdfs_and_emails` output, but ``pdf_path`` is set to
        ``""`` and ``existe_pdf`` to ``"False"`` (coordinators have no PDF unless
        one was explicitly generated for them).  The regional Excel path is stored
        under the ``regional_excel`` key.
    """
    dests: List[Dict[str, str]] = []

    if not zonas_file.exists():
        return dests

    df = pd.read_excel(zonas_file)
    col_map = {}
    for col in df.columns:
        normalized = col.strip().lower().replace("\ufffd", "").replace("?", "")
        col_map[col] = normalized
    df = df.rename(columns=col_map)

    for _, row in df.iterrows():
        nombre = str(row.get("nombre", "") or "").strip()
        nombre_up = nombre.upper()
        correo = str(row.get("correo", "") or "").strip()
        zona = str(row.get("zona", "") or "").strip()
        region = str(row.get("región", "") or row.get("regin", "") or row.get("region", "") or "").strip().upper()

        is_coord = any(kw in nombre_up for kw in _COORD_KEYWORDS)
        is_gerencia = "GERENCIA" in nombre_up
        if not is_coord or is_gerencia or not correo or "@" not in correo:
            continue

        # Look for individual PDF (may or may not exist)
        # Strip parentheses content for a simpler filename search
        import re
        base_name = re.sub(r'\(.*?\)', '', nombre).strip()
        pdf_candidates = list(pdf_dir.glob(f"*{base_name.replace(' ', '_')}*.pdf")) if pdf_dir.exists() else []
        if not pdf_candidates:
            pdf_candidates = list(pdf_dir.glob(f"*{base_name.split()[0]}*.pdf")) if pdf_dir.exists() else []

        pdf_path_str = str(pdf_candidates[0]) if pdf_candidates else ""
        tiene_pdf = bool(pdf_candidates and Path(pdf_path_str).exists())

        # Look for regional Excel
        prefix = _region_to_filename_prefix(region) if region else ""
        regional_excel_str = ""
        if prefix and regional_dir.exists():
            reg_cands = list(regional_dir.glob(f"{prefix}_Indicadores_*.xlsx"))
            if reg_cands:
                regional_excel_str = str(max(reg_cands, key=lambda p: p.stat().st_mtime))

        dests.append({
            "nombre": nombre,
            "zona": zona,
            "region": region,
            "email": correo,
            "pdf_path": pdf_path_str,
            "existe_pdf": str(tiene_pdf),
            "regional_excel": regional_excel_str,
            "es_coordinador": "True",
        })

    return dests


class EmailSender:
    """Gestiona el envío de correos vía SMTP (Gmail)."""

    def __init__(
        self,
        sender_email: str = DEFAULT_SENDER_EMAIL,
        sender_password: str = "",
        smtp_server: str = GMAIL_SMTP_SERVER,
        smtp_port: int = GMAIL_SMTP_PORT,
    ):
        self.sender_email = sender_email
        self.sender_password = sender_password.strip()
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def test_connection(self) -> Tuple[bool, str]:
        """Prueba la conexión y autenticación con el servidor SMTP."""
        if not self.sender_password:
            return False, "Debe ingresar la Contraseña de Aplicación de Gmail."

        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=15) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(self.sender_email, self.sender_password)
            return True, "Conexión y autenticación exitosas con Gmail."
        except smtplib.SMTPAuthenticationError:
            return False, (
                "Error de autenticación: Contraseña de Aplicación incorrecta.\n"
                "Nota: Para Gmail debe usar una 'Contraseña de aplicación' de 16 letras, "
                "no la contraseña normal."
            )
        except Exception as exc:
            return False, f"Error al conectar con el servidor: {exc}"

    def send_single_email(
        self,
        to_email: str,
        subject: str,
        body_text: str,
        attachments: Optional[List[Path]] = None,
        attachment_path: Optional[Path] = None,
    ) -> Tuple[bool, str]:
        """Envía un correo individual con uno o varios archivos adjuntos."""
        if not to_email or "@" not in to_email:
            return False, "Dirección de correo inválida o ausente."

        msg = MIMEMultipart()
        msg["From"] = f"MMC Sistemas <{self.sender_email}>"
        msg["To"] = to_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        files_to_attach: List[Path] = []
        if attachments:
            files_to_attach.extend(attachments)
        if attachment_path:
            files_to_attach.append(Path(attachment_path))

        for att_path in files_to_attach:
            att_path = Path(att_path)
            if att_path.exists():
                ctype, encoding = mimetypes.guess_type(str(att_path))
                if ctype is None or encoding is not None:
                    ctype = "application/octet-stream"
                maintype, subtype = ctype.split("/", 1)

                with open(att_path, "rb") as f:
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f'attachment; filename="{att_path.name}"',
                )
                msg.attach(part)
            else:
                logger.warning("Archivo adjunto no encontrado: %s", att_path)

        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=20) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
            return True, "Enviado con éxito"
        except Exception as exc:
            return False, str(exc)

    def send_batch(
        self,
        destinatarios: List[Dict[str, str]],
        subject: str,
        body_template: str,
        general_excel_path: Optional[Path] = None,
        extra_attachments: Optional[Dict[str, List[Path]]] = None,
        on_progress: Optional[Callable[[int, int, str, bool, str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None,
        delay_seconds: float = 1.0,
    ) -> Dict[str, int]:
        """Envía los correos en lote a todos los destinatarios con seguimiento.

        Args:
            destinatarios: Lista retornada por `map_vendors_to_pdfs_and_emails` o
                `get_coordinator_dests`.
            subject: Asunto del correo.
            body_template: Cuerpo del mensaje.
            general_excel_path: Ruta al archivo Excel General (se adjunta a los dueños/gerencia).
            extra_attachments: Dict mapping lowercase email → list of extra file Paths.
                Used to attach regional Excel files to coordinators.
            on_progress: Callback `fn(current, total, vendor_name, success, msg)`.
            cancel_check: Callback que retorna True si el usuario canceló el envío.
            delay_seconds: Pausa entre envíos para respetar límites de Gmail.

        Returns:
            Dict con contadores: 'enviados', 'fallidos', 'omitidos'.
        """
        stats = {"enviados": 0, "fallidos": 0, "omitidos": 0}
        total = len(destinatarios)
        extra_attachments = extra_attachments or {}

        context = ssl.create_default_context()
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=25) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(self.sender_email, self.sender_password)

                for idx, dest in enumerate(destinatarios, start=1):
                    if cancel_check and cancel_check():
                        logger.info("Envío cancelado por el usuario.")
                        break

                    nombre = dest["nombre"]
                    to_email = dest.get("email", "").strip()
                    pdf_path_str = dest.get("pdf_path", "")
                    pdf_path = Path(pdf_path_str) if pdf_path_str else None
                    regional_excel_str = dest.get("regional_excel", "")
                    es_coordinador = dest.get("es_coordinador", "False") == "True"

                    if not to_email:
                        stats["omitidos"] += 1
                        if on_progress:
                            on_progress(idx, total, nombre, False, "Sin correo registrado")
                        continue

                    # Coordinators: need at least a regional Excel; PDF is optional
                    tiene_pdf = pdf_path and pdf_path.exists()
                    tiene_regional = bool(regional_excel_str and Path(regional_excel_str).exists())

                    if es_coordinador:
                        if not tiene_pdf and not tiene_regional:
                            stats["omitidos"] += 1
                            if on_progress:
                                on_progress(idx, total, nombre, False, "Sin PDF ni Excel regional")
                            continue
                    else:
                        if not tiene_pdf:
                            stats["omitidos"] += 1
                            if on_progress:
                                on_progress(idx, total, nombre, False, "PDF no encontrado")
                            continue

                    # Determine if this recipient is an owner (gets General Excel)
                    es_dueno = (
                        to_email.lower() in OWNER_EMAILS
                        or "GERENCIA" in str(dest.get("region", "")).upper()
                        or "GERENCIA" in nombre.upper()
                    )

                    # Build attachment list
                    archivos_adjuntos: List[Path] = []
                    if tiene_pdf:
                        archivos_adjuntos.append(pdf_path)
                    if es_dueno and general_excel_path and Path(general_excel_path).exists():
                        archivos_adjuntos.append(Path(general_excel_path))
                    if tiene_regional:
                        archivos_adjuntos.append(Path(regional_excel_str))
                    # Extra attachments from caller (e.g. regional Excel keyed by email)
                    for extra_path in extra_attachments.get(to_email.lower(), []):
                        extra_path = Path(extra_path)
                        if extra_path.exists() and extra_path not in archivos_adjuntos:
                            archivos_adjuntos.append(extra_path)

                    if not archivos_adjuntos:
                        stats["omitidos"] += 1
                        if on_progress:
                            on_progress(idx, total, nombre, False, "Sin archivos adjuntos")
                        continue

                    # Prepare message
                    msg = MIMEMultipart()
                    msg["From"] = f"MMC Sistemas <{self.sender_email}>"
                    msg["To"] = to_email
                    msg["Subject"] = subject

                    custom_body = body_template.replace("Estimado(a) Comercial,", f"Estimado(a) {nombre},")
                    if es_dueno and general_excel_path:
                        custom_body += "\n\nNota: Adicionalmente se adjunta el archivo Reporte General de Indicadores en formato Excel."
                    if tiene_regional and es_coordinador:
                        custom_body += "\n\nNota: Se adjunta el Reporte Regional de Indicadores en formato Excel correspondiente a su zona."

                    msg.attach(MIMEText(custom_body, "plain", "utf-8"))

                    for att_file in archivos_adjuntos:
                        with open(att_file, "rb") as f:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f'attachment; filename="{att_file.name}"',
                        )
                        msg.attach(part)

                    # Send
                    try:
                        server.send_message(msg)
                        stats["enviados"] += 1
                        det_adj = f"{len(archivos_adjuntos)} adjuntos" if len(archivos_adjuntos) > 1 else "PDF adjunto"
                        if on_progress:
                            on_progress(idx, total, nombre, True, f"Enviado a {to_email} ({det_adj})")
                    except Exception as exc:
                        stats["fallidos"] += 1
                        if on_progress:
                            on_progress(idx, total, nombre, False, f"Error: {exc}")

                    if delay_seconds > 0:
                        time.sleep(delay_seconds)

        except Exception as exc:
            logger.error("Error global de conexión durante el lote: %s", exc)
            if on_progress:
                on_progress(0, total, "Servidor", False, f"Error de conexión: {exc}")

        return stats
