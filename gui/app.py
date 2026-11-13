"""gui/app.py — LYNX Report System.

Main Tkinter application window.

The window is organised into three tabs:

1. **Importar Datos** — Six file-picker rows for the Odoo source files.
2. **Generar Reportes** — Buttons to trigger Excel/PDF/all report generation.
3. **Mantenimiento** — Monthly maintenance tasks (Presupuestos, Cobranza Cierre,
   month label).

A progress bar and scrollable log area span the bottom of the window and are
shared across all tabs.

Long-running operations (ETL + report generation) run in a background
:class:`threading.Thread` to keep the UI responsive.  Results and log messages
are posted back to the main thread via Tkinter's ``after()`` mechanism using a
thread-safe :class:`queue.Queue`.

Typical usage (see also ``main.py``)::

    from gui.app import LynxApp
    app = LynxApp()
    app.mainloop()
"""

from __future__ import annotations

import logging
import os
import queue
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Any, Dict, Optional

import config
from core.engine import build_vendor_record, calculate_general, calculate_rankings
from core.etl import (
    calculate_working_days,
    read_backorder,
    read_devoluciones,
    read_facturado,
    read_pedidos,
    read_por_cobrar,
    read_por_facturar,
    read_presupuestos,
    read_zonas,
)
from core.exceptions import LynxBaseError
from core.models import AppConfig, SourceFiles
from gui.widgets import FilePickerRow, LogTextArea, SectionFrame
from maintenance.budget_loader import (
    load_presupuestos,
    load_presupuestos_from_config,
    save_presupuestos_config,
    sync_zonas_with_presupuestos,
)
from maintenance.cobranza_loader import (
    load_cobranza_cierre,
    save_cobranza_cierre,
)
from reports.excel_writer import generate_excel_report
from reports.manifest_writer import write_regional_manifest, write_vendor_manifest
from reports.pdf_writer import generate_pdf_reports
from services.email_sender import EmailSender, get_default_templates, map_vendors_to_pdfs_and_emails, get_coordinator_dests, build_coordinator_extra_attachments
from services.worker_directory import export_workers_directory_excel
from gui.worker_dialog import WorkerEmailDialog

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_WINDOW_TITLE = "LYNX — Sistema de Indicadores de Ventas"
_WINDOW_SIZE  = "1050x700"
_MIN_SIZE     = (900, 580)

# Source-file picker definitions: (label_text, internal_key)
_SOURCE_FILE_DEFS = [
    ("Facturado (cierre):",         "facturado"),
    ("Pedidos Recibidos:",          "pedidos"),
    ("Devoluciones:",               "devoluciones"),
    ("Cuentas por Cobrar:",         "por_cobrar"),
    ("Pedidos por Facturar:",       "por_facturar"),
    ("Backorder:",                  "backorder"),
]

# Default paths (relative to project root)
_DATA_DIR   = Path(__file__).parent.parent / "data"
_OUTPUT_DIR = Path(__file__).parent.parent / "output"


# ---------------------------------------------------------------------------
# Queue message protocol
# ---------------------------------------------------------------------------

class _Msg:
    """Simple message envelope posted from worker threads to the GUI queue."""

    def __init__(self, kind: str, payload: Any = None) -> None:
        self.kind    = kind    # "log", "progress", "done", "error"
        self.payload = payload


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------


class LynxApp(tk.Tk):
    """Root Tkinter window for the LYNX sales indicator system.

    Inherits from :class:`tkinter.Tk` so it *is* the main event loop.
    Call :meth:`mainloop` to start the application.

    Args:
        output_dir: Directory where reports are saved.  Defaults to
            ``<project_root>/output``.
        presupuestos_config: Path to the saved presupuestos Excel file.
            Defaults to ``<project_root>/data/presupuestos.xlsx``.
        zonas_config: Path to the zonas/zones definition file.
            Defaults to ``<project_root>/data/zonas.xlsx``.
    """

    def __init__(
        self,
        output_dir: Path = _OUTPUT_DIR,
        presupuestos_config: Path = _DATA_DIR / "presupuestos.xlsx",
        zonas_config: Path = _DATA_DIR / "zonas.xlsx",
    ) -> None:
        super().__init__()

        self.title(_WINDOW_TITLE)
        self.geometry(_WINDOW_SIZE)
        self.minsize(*_MIN_SIZE)

        # State
        self._output_dir          = output_dir
        self._presupuestos_config = presupuestos_config
        self._zonas_config        = zonas_config
        self._gui_queue: queue.Queue[_Msg] = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._cobranza_cierre_path: Optional[str] = None

        # Month / year spinners (must be created before _build_ui)
        self._mes_var  = tk.StringVar(value=config.MESES_ES[0])
        self._anio_var = tk.IntVar(value=2026)

        self._build_ui()
        self._load_saved_state()
        self._poll_queue()

        logger.info("LYNX application started.")

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        """Construct and layout all UI widgets."""
        # ---- Notebook (tabs) ----------------------------------------
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        self._tab_import    = ttk.Frame(notebook, padding=10)
        self._tab_reports   = ttk.Frame(notebook, padding=10)
        self._tab_email     = ttk.Frame(notebook, padding=10)
        self._tab_maint     = ttk.Frame(notebook, padding=10)

        notebook.add(self._tab_import,  text="📂  Importar Datos")
        notebook.add(self._tab_reports, text="📊  Generar Reportes")
        notebook.add(self._tab_email,   text="📧  Envío de Correos")
        notebook.add(self._tab_maint,   text="🔧  Mantenimiento")

        self._build_import_tab()
        self._build_reports_tab()
        self._build_email_tab()
        self._build_maintenance_tab()

        # ---- Bottom area: progress + log ----------------------------
        bottom = ttk.Frame(self, padding=(10, 5))
        bottom.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        # Progress bar
        prog_frame = ttk.Frame(bottom)
        prog_frame.pack(fill="x")
        ttk.Label(prog_frame, text="Estado:").pack(side="left")
        self._progress_lbl = ttk.Label(prog_frame, text="Listo.", foreground="#444444")
        self._progress_lbl.pack(side="left", padx=8)
        self._progressbar = ttk.Progressbar(
            prog_frame, mode="indeterminate", length=200
        )
        self._progressbar.pack(side="right", padx=(0, 4))
        self._progress_bar = self._progressbar  # alias for backward compatibility

        # Log area
        log_frame = SectionFrame(bottom, "Registro de Eventos")
        log_frame.pack(fill="both", expand=True, pady=(6, 0))
        self._log_area = LogTextArea(log_frame, height=8)
        self._log_area.pack(fill="both", expand=True)

        # Clear log button
        ttk.Button(
            bottom, text="Limpiar registro", command=self._log_area.clear, width=16
        ).pack(anchor="e", pady=(2, 0))

    def _build_import_tab(self) -> None:
        """Build the 'Importar Datos' tab contents."""
        frame = SectionFrame(self._tab_import, "Archivos fuente de Odoo")
        frame.pack(fill="x", pady=(0, 10))

        self._file_pickers: Dict[str, FilePickerRow] = {}

        for label, key in _SOURCE_FILE_DEFS:
            row = FilePickerRow(frame, label_text=label, label_width=24)
            row.pack(fill="x", pady=2)
            self._file_pickers[key] = row

        # Buttons row
        btn_frame = ttk.Frame(self._tab_import)
        btn_frame.pack(fill="x", pady=4)
        ttk.Button(
            btn_frame, text="Validar archivos seleccionados",
            command=self._validate_files, width=28,
        ).pack(side="left", padx=4)
        ttk.Button(
            btn_frame, text="Limpiar selección",
            command=self._clear_file_selections, width=18,
        ).pack(side="left")

        # Info label
        ttk.Label(
            self._tab_import,
            text="Seleccione los 6 archivos exportados desde Odoo antes de generar reportes.",
            foreground="#555555",
            font=("Segoe UI", 8, "italic"),
        ).pack(anchor="w", pady=(6, 0))

    def _build_reports_tab(self) -> None:
        """Build the 'Generar Reportes' tab contents."""
        # Output directory selector
        out_frame = SectionFrame(self._tab_reports, "Directorio de salida")
        out_frame.pack(fill="x", pady=(0, 10))

        self._output_var = tk.StringVar(value=str(self._output_dir))
        ttk.Label(out_frame, text="Carpeta de salida:", width=18, anchor="w").pack(
            side="left"
        )
        ttk.Entry(out_frame, textvariable=self._output_var, width=50).pack(
            side="left", padx=4, fill="x", expand=True
        )
        ttk.Button(
            out_frame, text="Cambiar…", command=self._browse_output_dir, width=10
        ).pack(side="left")

        # Report buttons
        rpt_frame = SectionFrame(self._tab_reports, "Acciones")
        rpt_frame.pack(fill="x", pady=(0, 10))

        btn_cfg = [
            ("📊  Reporte General + Regionales (Excel)",
             self._run_excel_reports, "#1F3864", "white"),
            ("📄  Generar PDFs Individuales",
             self._run_pdf_reports,   "#2E7D32", "white"),
            ("🚀  Generar TODO (Excel + PDFs + Manifiestos)",
             self._run_all_reports,   "#4A148C", "white"),
        ]

        for text, cmd, bg, fg in btn_cfg:
            btn = tk.Button(
                rpt_frame, text=text, command=cmd,
                bg=bg, fg=fg, font=("Segoe UI", 10, "bold"),
                relief="flat", padx=12, pady=8, cursor="hand2",
                activebackground="#374151", activeforeground="white",
            )
            btn.pack(fill="x", pady=3)

    def _build_maintenance_tab(self) -> None:
        """Build the 'Mantenimiento' tab contents."""
        # --- Presupuestos section ---
        pres_frame = SectionFrame(self._tab_maint, "Presupuestos (metas mensuales)")
        pres_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            pres_frame,
            text=(
                "El gerente entrega un archivo Excel con las metas de cada vendedor.\n"
                "Cárguelo aquí una vez al mes para actualizar los presupuestos."
            ),
            foreground="#444444",
        ).pack(anchor="w", pady=(0, 6))

        self._pres_picker = FilePickerRow(
            pres_frame, label_text="Archivo de Presupuestos:", label_width=24
        )
        self._pres_picker.pack(fill="x", pady=2)

        ttk.Button(
            pres_frame,
            text="Cargar y guardar Presupuestos",
            command=self._load_presupuestos,
            width=30,
        ).pack(anchor="w", pady=(6, 0))

        # --- Mes / Año section ---
        mes_frame = SectionFrame(self._tab_maint, "Mes del reporte")
        mes_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(mes_frame, text="Mes:", width=8, anchor="w").pack(side="left")
        mes_combo = ttk.Combobox(
            mes_frame, textvariable=self._mes_var,
            values=config.MESES_ES, state="readonly", width=14,
        )
        mes_combo.pack(side="left", padx=(0, 12))

        ttk.Label(mes_frame, text="Año:", width=6, anchor="w").pack(side="left")
        ttk.Spinbox(
            mes_frame, from_=2024, to=2035, textvariable=self._anio_var,
            width=6, format="%4.0f",
        ).pack(side="left", padx=(0, 12))

        ttk.Button(
            mes_frame, text="Aplicar mes",
            command=self._apply_mes_label, width=14,
        ).pack(side="left")

        self._mes_label_lbl = ttk.Label(
            mes_frame, text="", foreground="#006400", font=("Segoe UI", 9, "bold")
        )
        self._mes_label_lbl.pack(side="left", padx=8)

        # --- Cobranza Cierre section ---
        cob_frame = SectionFrame(
            self._tab_maint, "Cobranza Cierre (fin de mes)"
        )
        cob_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            cob_frame,
            text=(
                "Al cierre de mes el gerente envía el archivo INDICADOR_COBRANZA_MES_AÑO.xlsx.\n"
                "Cárguelo aquí para que los datos de cobranza reflejen el cierre real."
            ),
            foreground="#444444",
        ).pack(anchor="w", pady=(0, 6))

        self._cob_picker = FilePickerRow(
            cob_frame, label_text="Archivo Indicador Cobranza:", label_width=26
        )
        self._cob_picker.pack(fill="x", pady=2)

        cob_btn_row = ttk.Frame(cob_frame)
        cob_btn_row.pack(fill="x", pady=(6, 0))

        ttk.Button(
            cob_btn_row,
            text="Cargar Cobranza Cierre",
            command=self._load_cobranza_cierre,
            width=24,
        ).pack(side="left")

        ttk.Button(
            cob_btn_row,
            text="Modo Automático (Odoo)",
            command=self._clear_cobranza_cierre,
            width=24,
        ).pack(side="left", padx=(8, 0))

        self._cob_status_lbl = ttk.Label(
            cob_frame,
            text="Modo actual: Automático desde Odoo (mes en curso)",
            foreground="#006400",
            font=("Segoe UI", 9, "bold"),
        )
        self._cob_status_lbl.pack(anchor="w", pady=(6, 0))

        # --- Directorio de Trabajadores y Correos ---
        worker_maint_frame = SectionFrame(self._tab_maint, "Directorio de Trabajadores, Correos y Metas")
        worker_maint_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            worker_maint_frame,
            text=(
                "Consulte la lista completa de comerciales y coordinadores, descargue un archivo Excel\n"
                "con sus correos y metas del mes, o modifique los correos de destino cuando haya cambios de personal."
            ),
            foreground="#444444",
        ).pack(anchor="w", pady=(0, 6))

        row_w_btns = ttk.Frame(worker_maint_frame)
        row_w_btns.pack(fill="x", pady=(4, 2))

        ttk.Button(
            row_w_btns,
            text="📥 Descargar Directorio y Metas (Excel)",
            command=self._download_workers_directory,
            width=36,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            row_w_btns,
            text="✏ Modificar Correos de Trabajadores",
            command=self._open_worker_email_dialog,
            width=34,
        ).pack(side="left")

    def _build_email_tab(self) -> None:
        """Construct the 'Envío de Correos' tab widgets."""
        # --- Configuración del Remitente ---
        cfg_frame = SectionFrame(self._tab_email, "1. Configuración de Cuenta Gmail (Remitente)")
        cfg_frame.pack(fill="x", pady=(0, 10))

        row1 = ttk.Frame(cfg_frame)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="Correo Remitente:", width=20, anchor="w").pack(side="left")
        self._email_sender_var = tk.StringVar(value="sistemas3.mmcve@gmail.com")
        ttk.Entry(row1, textvariable=self._email_sender_var, width=32).pack(side="left", padx=5)

        ttk.Label(row1, text="Contraseña de Aplicación:", width=24, anchor="w").pack(side="left", padx=(15, 0))
        self._email_pass_var = tk.StringVar(value="zymz wtjv rodt rpyu")
        ttk.Entry(row1, textvariable=self._email_pass_var, show="•", width=22).pack(side="left", padx=5)

        ttk.Button(row1, text="Probar Conexión", command=self._test_email_connection, width=16).pack(side="left", padx=10)

        help_lbl = ttk.Label(
            cfg_frame,
            text="ℹ Nota: Debe usar una 'Contraseña de Aplicación' de 16 caracteres generada en su cuenta Google, NO la clave habitual.",
            font=("Segoe UI", 8, "italic"),
            foreground="#666666",
        )
        help_lbl.pack(anchor="w", pady=(2, 0))

        # --- Tipo de Reporte y Plantilla ---
        msg_frame = SectionFrame(self._tab_email, "2. Modo de Envío y Mensaje")
        msg_frame.pack(fill="x", pady=(0, 10))

        row_mode = ttk.Frame(msg_frame)
        row_mode.pack(fill="x", pady=4)
        ttk.Label(row_mode, text="Tipo de Reporte:", width=20, anchor="w").pack(side="left")

        self._email_tipo_var = tk.StringVar(value="parcial")
        r1 = ttk.Radiobutton(
            row_mode,
            text="Mes en Curso (de lo que va del mes)",
            variable=self._email_tipo_var,
            value="parcial",
            command=self._update_email_templates,
        )
        r1.pack(side="left", padx=(0, 15))

        r2 = ttk.Radiobutton(
            row_mode,
            text="Cierre de Mes (definitivo)",
            variable=self._email_tipo_var,
            value="cierre",
            command=self._update_email_templates,
        )
        r2.pack(side="left")

        # Asunto
        row_sub = ttk.Frame(msg_frame)
        row_sub.pack(fill="x", pady=4)
        ttk.Label(row_sub, text="Asunto del Correo:", width=20, anchor="w").pack(side="left")
        self._email_subject_var = tk.StringVar()
        ttk.Entry(row_sub, textvariable=self._email_subject_var, width=65).pack(side="left", fill="x", expand=True, padx=5)

        # Cuerpo
        ttk.Label(msg_frame, text="Cuerpo del Mensaje (El nombre del comercial se personaliza automáticamente):").pack(anchor="w", pady=(6, 2))
        self._email_body_text = tk.Text(msg_frame, height=5, font=("Segoe UI", 9), wrap="word")
        self._email_body_text.pack(fill="x", pady=(0, 6))

        # --- Destinatarios y Envío ---
        send_frame = SectionFrame(self._tab_email, "3. Envío y Pruebas")
        send_frame.pack(fill="x", pady=(0, 10))

        # Selector de carpeta de reportes / PDFs
        row_dir = ttk.Frame(send_frame)
        row_dir.pack(fill="x", pady=(2, 4))
        ttk.Label(row_dir, text="Carpeta de Reportes:", width=20, anchor="w").pack(side="left")
        ttk.Entry(row_dir, textvariable=self._output_var, state="readonly").pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row_dir, text="Cambiar...", command=self._browse_output_dir, width=12).pack(side="left")

        # Fila de gestión de directorio y modificación de correos
        row_dir_tools = ttk.Frame(send_frame)
        row_dir_tools.pack(fill="x", pady=(2, 6))

        ttk.Button(
            row_dir_tools,
            text="📥 Descargar Directorio y Metas (Excel)",
            command=self._download_workers_directory,
            width=36,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            row_dir_tools,
            text="✏ Modificar Correos de Trabajadores",
            command=self._open_worker_email_dialog,
            width=34,
        ).pack(side="left")

        row_actions = ttk.Frame(send_frame)
        row_actions.pack(fill="x", pady=4)

        ttk.Button(
            row_actions,
            text="Verificar Destinatarios",
            command=self._verify_recipients,
            width=22,
        ).pack(side="left", padx=(0, 10))

        ttk.Label(row_actions, text="Correo de prueba:", width=16, anchor="w").pack(side="left")
        self._email_test_var = tk.StringVar(value="sistemas3.mmcve@gmail.com")
        ttk.Entry(row_actions, textvariable=self._email_test_var, width=28).pack(side="left", padx=5)

        ttk.Button(
            row_actions,
            text="Enviar Prueba",
            command=self._send_test_email,
            width=16,
        ).pack(side="left", padx=5)

        self._btn_send_all = ttk.Button(
            row_actions,
            text="✉ ENVIAR A TODOS LOS COMERCIALES",
            command=self._start_batch_email_send,
            style="Accent.TButton",
            width=34,
        )
        self._btn_send_all.pack(side="right", padx=5)

        self._update_email_templates()

    def _update_email_templates(self) -> None:
        """Actualiza el asunto y cuerpo según el modo seleccionado (parcial vs cierre)."""
        mes_label = self._get_mes_label()
        es_cierre = (self._email_tipo_var.get() == "cierre")
        subj, body = get_default_templates(mes_label, es_cierre=es_cierre)
        self._email_subject_var.set(subj)
        self._email_body_text.delete("1.0", "end")
        self._email_body_text.insert("1.0", body)

    def _test_email_connection(self) -> None:
        """Prueba autenticación SMTP con Gmail."""
        sender = EmailSender(
            sender_email=self._email_sender_var.get(),
            sender_password=self._email_pass_var.get(),
        )
        ok, msg = sender.test_connection()
        if ok:
            messagebox.showinfo("Conexión Exitosa", msg)
            self._log(f"✔ {msg}", "SUCCESS")
        else:
            messagebox.showerror("Error de Conexión", msg)
            self._log(f"✖ {msg}", "ERROR")

    def _get_active_output_dir(self) -> Path:
        """Return the current output directory from the UI entry or fallback."""
        val = self._output_var.get().strip() if hasattr(self, "_output_var") else ""
        if val:
            return Path(val)
        return self._output_dir

    def _verify_recipients(self) -> None:
        """Verifica la lista de comerciales, correos y PDFs generados."""
        out_dir = self._get_active_output_dir()
        manifest_path = out_dir / "manifest_indicadores_vendedores.csv"
        if not manifest_path.exists():
            alt_m = out_dir / "manifest_vendedores.csv" / "manifest_indicadores_vendedores.csv"
            if alt_m.exists():
                manifest_path = alt_m

        zonas_path = Path("data/zonas.xlsx")
        pdf_dir = out_dir / "PDF_INDICADORES"
        regional_dir = out_dir / "REGIONALES"

        if not manifest_path.exists():
            messagebox.showwarning(
                "Manifiesto no encontrado",
                "Primero debe generar los reportes en la pestaña 'Generar Reportes' para crear los PDFs y la lista de comerciales.",
                parent=self,
            )
            return

        dests = map_vendors_to_pdfs_and_emails(manifest_path, zonas_path, pdf_dir)
        coord_dests = get_coordinator_dests(zonas_path, regional_dir, pdf_dir)

        listos = [d for d in dests if d["email"] and d["existe_pdf"] == "True"]
        sin_correo = [d for d in dests if not d["email"]]
        sin_pdf = [d for d in dests if d["existe_pdf"] != "True"]
        coord_listos = [
            d for d in coord_dests
            if d["email"] and (d["existe_pdf"] == "True" or d.get("regional_excel"))
        ]

        self._log(
            f"Verificación: {len(listos)} vendedores + {len(coord_listos)} coordinadores listos "
            f"(total manifiesto: {len(dests)}).",
            "INFO",
        )
        if sin_correo:
            self._log(
                f"Sin correo ({len(sin_correo)}): {', '.join(d['nombre'] for d in sin_correo[:5])}...",
                "WARNING",
            )
        if sin_pdf:
            self._log(
                f"Sin PDF ({len(sin_pdf)}): {', '.join(d['nombre'] for d in sin_pdf[:5])}...",
                "WARNING",
            )

        messagebox.showinfo(
            "Verificación de Destinatarios",
            f"Total vendedores en manifiesto: {len(dests)}\n"
            f"  Listos con PDF y Correo: {len(listos)}\n"
            f"  Sin correo asignado: {len(sin_correo)}\n"
            f"  Sin PDF encontrado: {len(sin_pdf)}\n\n"
            f"Coordinadores/Supervisores: {len(coord_dests)}\n"
            f"  Listos (PDF y/o Excel Regional): {len(coord_listos)}",
            parent=self,
        )

    def _send_test_email(self) -> None:
        """Envía un correo de prueba con un PDF muestra."""
        test_email = self._email_test_var.get().strip()
        if not test_email or "@" not in test_email:
            messagebox.showwarning("Atención", "Ingrese una dirección de correo válida para la prueba.")
            return

        sender = EmailSender(
            sender_email=self._email_sender_var.get(),
            sender_password=self._email_pass_var.get(),
        )

        # Buscar cualquier PDF generado en el directorio de salida
        out_dir = self._get_active_output_dir()
        pdf_dir = out_dir / "PDF_INDICADORES"
        muestra_pdf = None
        if pdf_dir.exists():
            pdfs = list(pdf_dir.glob("*.pdf"))
            if pdfs:
                muestra_pdf = pdfs[0]

        subject = "[PRUEBA] " + self._email_subject_var.get()
        body = self._email_body_text.get("1.0", "end-1c")

        self._log(f"Enviando correo de prueba a {test_email}...", "INFO")
        ok, msg = sender.send_single_email(
            to_email=test_email,
            subject=subject,
            body_text=body,
            attachment_path=muestra_pdf,
        )
        if ok:
            adjunto_txt = f" (Adjunto: {muestra_pdf.name})" if muestra_pdf else " (Sin adjunto)"
            messagebox.showinfo("Éxito", f"Correo de prueba enviado a {test_email}{adjunto_txt}", parent=self)
            self._log(f"✔ Correo de prueba enviado a {test_email}", "SUCCESS")
        else:
            messagebox.showerror("Error de Envío", f"No se pudo enviar el correo de prueba:\n{msg}", parent=self)
            self._log(f"✖ Error enviando correo de prueba: {msg}", "ERROR")

    def _start_batch_email_send(self) -> None:
        """Inicia el envío masivo en un hilo secundario.

        Sends individual PDFs to all vendors in the manifest, plus the regional
        Excel file to each coordinator/supervisor found in zonas.xlsx.
        """
        try:
            if not self._email_pass_var.get().strip():
                messagebox.showwarning("Falta Contraseña", "Ingrese la Contraseña de Aplicación de Gmail.", parent=self)
                return

            out_dir = self._get_active_output_dir()
            manifest_path = out_dir / "manifest_indicadores_vendedores.csv"
            if not manifest_path.exists():
                alt_m = out_dir / "manifest_vendedores.csv" / "manifest_indicadores_vendedores.csv"
                if alt_m.exists():
                    manifest_path = alt_m

            if not manifest_path.exists():
                messagebox.showwarning(
                    "Reportes no encontrados",
                    f"No se encontró el manifiesto de vendedores en:\n{out_dir}\n\n"
                    "Primero genere los reportes desde la pestaña 'Generar Reportes' o seleccione la carpeta donde están guardados.",
                    parent=self,
                )
                return

            zonas_path = Path("data/zonas.xlsx")
            pdf_dir = out_dir / "PDF_INDICADORES"
            regional_dir = out_dir / "REGIONALES"

            # --- Vendor destinations from manifest ---
            dests = map_vendors_to_pdfs_and_emails(manifest_path, zonas_path, pdf_dir)
            listos = [d for d in dests if d["email"] and d["existe_pdf"] == "True"]

            # --- Coordinator destinations (not in manifest) ---
            coord_dests = get_coordinator_dests(zonas_path, regional_dir, pdf_dir)
            coord_listos = [
                d for d in coord_dests
                if d["email"] and (d["existe_pdf"] == "True" or d.get("regional_excel"))
            ]

            # Combined list for the batch
            todos = dests + coord_dests
            total_listos = len(listos) + len(coord_listos)

            if total_listos == 0:
                messagebox.showwarning(
                    "Sin destinatarios listos",
                    f"Se leyeron {len(dests)} vendedores del manifiesto y {len(coord_dests)} coordinadores, "
                    "pero ninguno tiene PDF/Excel y correo asignado.\n"
                    "Verifique que los PDFs estén en la carpeta PDF_INDICADORES.",
                    parent=self,
                )
                return

            # --- General Excel for owners ---
            excel_general = None
            cand_excels = list(out_dir.glob("Reporte_General_*.xlsx")) + list(out_dir.glob("Reporte_General_*.xls"))
            if cand_excels:
                excel_general = cand_excels[0]

            # --- Extra attachments for coordinators ---
            extra_map = build_coordinator_extra_attachments(zonas_path, regional_dir)

            duenos_info = f"\n(Se adjuntará también {excel_general.name} a los directores/dueños)" if excel_general else ""
            coord_info = f"\n({len(coord_listos)} coordinadores recibirán además su Excel Regional)" if coord_listos else ""

            confirm = messagebox.askyesno(
                "Confirmar Envío Masivo",
                f"Se enviarán correos a {len(listos)} vendedores y {len(coord_listos)} coordinadores.{duenos_info}{coord_info}\n\n"
                f"Remitente: {self._email_sender_var.get()}\n"
                f"Asunto: {self._email_subject_var.get()}\n\n"
                "¿Desea iniciar el proceso de envío ahora?",
                parent=self,
            )
            if not confirm:
                return

            self._set_ui_busy(True)
            self._progressbar["mode"] = "determinate"
            self._progressbar["value"] = 0
            self._progress_lbl.config(text=f"Iniciando envío a {total_listos} destinatarios...")
            self._log(f"Iniciando proceso de envío masivo para {total_listos} destinatarios...", "INFO")

            sender = EmailSender(
                sender_email=self._email_sender_var.get(),
                sender_password=self._email_pass_var.get(),
            )
            subject = self._email_subject_var.get()
            body = self._email_body_text.get("1.0", "end-1c")

            def worker():
                try:
                    self._gui_queue.put(_Msg("log", (
                        f"Iniciando envío — {len(listos)} vendedores + {len(coord_listos)} coordinadores...",
                        "INFO",
                    )))
                    if excel_general:
                        self._gui_queue.put(_Msg("log", (f"Excel General para dueños: {excel_general.name}", "INFO")))
                    if coord_listos:
                        self._gui_queue.put(_Msg("log", (
                            f"Coordinadores con Excel Regional: {', '.join(d['nombre'] for d in coord_listos)}",
                            "INFO",
                        )))

                    def on_prog(curr, tot, nom, ok, m):
                        lvl = "SUCCESS" if ok else "WARNING"
                        self._gui_queue.put(_Msg("progress", (curr, tot, f"Enviando {curr}/{tot}: {nom}")))
                        self._gui_queue.put(_Msg("log", (f"[{curr}/{tot}] {nom} → {m}", lvl)))

                    stats = sender.send_batch(
                        destinatarios=todos,
                        subject=subject,
                        body_template=body,
                        general_excel_path=excel_general,
                        extra_attachments=extra_map,
                        on_progress=on_prog,
                        delay_seconds=0.8,
                    )
                    self._gui_queue.put(_Msg("email_done", stats))
                except Exception as exc:
                    logger.error("Error en hilo de envío de correos: %s", exc, exc_info=True)
                    self._gui_queue.put(_Msg("log", (f"Error crítico en envío de correos: {exc}", "ERROR")))
                    self._gui_queue.put(_Msg("email_done", {"enviados": 0, "fallidos": total_listos, "omitidos": 0}))

            t = threading.Thread(target=worker, daemon=True)
            t.start()
        except Exception as exc:
            logger.error("Error al preparar envío masivo de correos: %s", exc, exc_info=True)
            self._set_ui_busy(False)
            self._log(f"Error al iniciar envío masivo: {exc}", "ERROR")
            messagebox.showerror("Error", f"No se pudo iniciar el envío masivo:\n{exc}", parent=self)

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------

    def _set_ui_busy(self, busy: bool) -> None:
        """Habilita o deshabilita controles interactivos durante tareas en segundo plano."""
        state = "disabled" if busy else "normal"
        cursor = "watch" if busy else ""
        try:
            self.config(cursor=cursor)
        except Exception:
            pass

        if hasattr(self, "_btn_send_all"):
            try:
                self._btn_send_all.config(state=state)
            except Exception:
                pass
        if hasattr(self, "_btn_run"):
            try:
                self._btn_run.config(state=state)
            except Exception:
                pass

    def _open_worker_email_dialog(self) -> None:
        """Abre la ventana para consultar trabajadores, metas y modificar correos."""
        WorkerEmailDialog(
            parent=self,
            zonas_path=self._zonas_config,
            presupuestos_path=self._presupuestos_config,
            mes_label=self._get_mes_label(),
            on_log=self._log,
        )

    def _download_workers_directory(self) -> None:
        """Genera y descarga un archivo en Excel con los trabajadores, correos y metas del mes."""
        mes_label = self._get_mes_label()
        initial_file = f"Directorio_Trabajadores_Metas_{mes_label.replace(' ', '_')}.xlsx"

        dest = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar Directorio de Trabajadores y Metas en Excel",
            initialfile=initial_file,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not dest:
            return

        try:
            out_path = export_workers_directory_excel(
                zonas_path=self._zonas_config,
                presupuestos_path=self._presupuestos_config,
                output_file=Path(dest),
                mes_label=mes_label,
            )
            msg = f"Directorio de trabajadores y metas exportado con éxito a:\n{out_path}"
            self._log(f"✔ {msg}", "SUCCESS")
            messagebox.showinfo("Exportación Exitosa", msg, parent=self)
        except Exception as exc:
            err_msg = f"Error al exportar directorio de trabajadores:\n{exc}"
            self._log(f"✖ {err_msg}", "ERROR")
            messagebox.showerror("Error", err_msg, parent=self)

    def _browse_output_dir(self) -> None:
        """Open a folder-picker and update the output directory entry."""
        path = filedialog.askdirectory(initialdir=self._output_var.get())
        if path:
            self._output_var.set(path)
            self._output_dir = Path(path)

    def _validate_files(self) -> None:
        """Check that all 6 source files are selected and exist."""
        all_ok = True
        for key, picker in self._file_pickers.items():
            path = picker.get()
            if not path or not os.path.isfile(path):
                self._log(f"[FALTA] Archivo no seleccionado: {key}", "WARNING")
                all_ok = False
        if all_ok:
            self._log("✔ Todos los archivos fuente están seleccionados.", "SUCCESS")
        else:
            self._log("⚠ Seleccione todos los archivos antes de generar reportes.", "WARNING")

    def _clear_file_selections(self) -> None:
        """Clear all file picker entries."""
        for picker in self._file_pickers.values():
            picker.set("")
        self._log("Selección de archivos limpiada.", "INFO")

    def _apply_mes_label(self) -> None:
        """Update the month label from the spinners."""
        label = f"{self._mes_var.get()} {self._anio_var.get()}"
        self._mes_label_lbl.config(text=f"→ {label}")
        self._log(f"Mes actualizado: {label}", "SUCCESS")

    def _get_mes_label(self) -> str:
        """Return the currently configured month label string.

        Returns:
            E.g. ``"Noviembre 2026"``.
        """
        return f"{self._mes_var.get()} {self._anio_var.get()}"

    def _log(self, message: str, level: str = "INFO") -> None:
        """Append a message to the on-screen log area (thread-safe via queue).

        Args:
            message: Text to display.
            level: Severity level string.
        """
        if threading.current_thread() is threading.main_thread():
            try:
                self._log_area.append(message, level)
            except Exception:
                self._gui_queue.put(_Msg("log", (message, level)))
        else:
            self._gui_queue.put(_Msg("log", (message, level)))

    # ------------------------------------------------------------------
    # State persistence
    # ------------------------------------------------------------------

    def _load_saved_state(self) -> None:
        """Load previously saved presupuestos config and cobranza cierre if available."""
        if self._presupuestos_config.exists():
            self._log(
                f"Presupuestos guardados encontrados: {self._presupuestos_config.name}",
                "INFO",
            )

        # Auto-detect saved cobranza cierre file in data/
        _default_cob = _DATA_DIR / "cobranza_cierre.xlsx"
        if _default_cob.exists() and not self._cobranza_cierre_path:
            self._cobranza_cierre_path = str(_default_cob)
            self._log(
                f"Cobranza cierre encontrada: {_default_cob.name} — se aplicará como cierre oficial.",
                "INFO",
            )
            # Pre-fill the picker entry if it exists
            try:
                self._cob_picker.set(str(_default_cob))
                if hasattr(self, "_cob_status_lbl"):
                    self._cob_status_lbl.config(
                        text=f"Modo: Cierre definitivo activo ({_default_cob.name})",
                        foreground="#1F3864",
                    )
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Queue polling (thread → main thread communication)
    # ------------------------------------------------------------------

    def _poll_queue(self) -> None:
        """Check the GUI queue and dispatch messages from worker threads.

        Schedules itself to run again via :meth:`after` every 100 ms.
        """
        try:
            while True:
                msg = self._gui_queue.get_nowait()
                if msg.kind == "log":
                    text, level = msg.payload
                    self._log_area.append(text, level)
                elif msg.kind == "progress":
                    if isinstance(msg.payload, tuple) and len(msg.payload) == 3:
                        curr, tot, text = msg.payload
                        pct = int((curr / tot) * 100) if tot > 0 else 0
                        self._progressbar["value"] = pct
                        self._progress_lbl.config(text=text)
                    else:
                        self._progress_lbl.config(text=str(msg.payload))
                elif msg.kind == "done":
                    self._on_worker_done(msg.payload)
                elif msg.kind == "email_done":
                    self._on_email_done(msg.payload)
                elif msg.kind == "cob_status":
                    if hasattr(self, "_cob_status_lbl"):
                        mode, text = msg.payload
                        fg = "#1F3864" if mode == "cierre" else "#006400"
                        self._cob_status_lbl.config(text=text, foreground=fg)
                elif msg.kind == "error":
                    self._on_worker_error(msg.payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _post(self, kind: str, payload: Any = None) -> None:
        """Post a message to the GUI queue (callable from any thread).

        Args:
            kind: Message type string.
            payload: Associated data.
        """
        self._gui_queue.put(_Msg(kind, payload))

    # ------------------------------------------------------------------
    # Worker lifecycle
    # ------------------------------------------------------------------

    def _start_worker(self, target, *args) -> None:
        """Launch a background thread for a long-running operation.

        Only one worker runs at a time; a new one cannot start until the
        current one finishes.

        Args:
            target: Callable to run in the background thread.
            *args: Positional arguments passed to *target*.
        """
        if self._worker and self._worker.is_alive():
            messagebox.showwarning(
                "Operación en curso",
                "Ya hay una operación en ejecución. Espere a que termine.",
            )
            return

        self._progressbar.start(10)
        self._progress_lbl.config(text="Procesando…")

        self._worker = threading.Thread(target=target, args=args, daemon=True)
        self._worker.start()

    def _on_worker_done(self, summary: str) -> None:
        """Handle successful completion of a worker thread.

        Args:
            summary: Human-readable summary of what was produced.
        """
        self._progressbar.stop()
        self._progress_lbl.config(text="Listo.")
        self._log(f"✔ {summary}", "SUCCESS")
        messagebox.showinfo("Éxito", summary)

    def _on_worker_error(self, error_msg: str) -> None:
        """Handle an error from a worker thread.

        Args:
            error_msg: Description of the error.
        """
        self._progressbar.stop()
        self._progress_lbl.config(text="Error.")
        self._log(f"✖ {error_msg}", "ERROR")
        messagebox.showerror("Error", error_msg)

    # ------------------------------------------------------------------
    # ETL helpers
    # ------------------------------------------------------------------

    def _build_source_files(self) -> Optional[SourceFiles]:
        """Collect file picker paths into a :class:`~core.models.SourceFiles` object.

        Returns:
            Populated :class:`~core.models.SourceFiles`, or ``None`` if any
            required file is missing.
        """
        paths = {key: picker.get() for key, picker in self._file_pickers.items()}
        missing = [k for k, v in paths.items() if not v or not os.path.isfile(v)]
        if missing:
            self._post("error", f"Archivos faltantes: {', '.join(missing)}")
            return None

        return SourceFiles(
            facturado   = paths["facturado"],
            pedidos     = paths["pedidos"],
            devoluciones= paths["devoluciones"],
            por_cobrar  = paths["por_cobrar"],
            por_facturar= paths["por_facturar"],
            backorder   = paths["backorder"],
        )

    def _build_app_config(self) -> AppConfig:
        """Build the :class:`~core.models.AppConfig` from current UI state.

        Returns:
            :class:`~core.models.AppConfig` instance.
        """
        return AppConfig(
            mes_label  = self._get_mes_label(),
            output_dir = str(self._output_dir),
            regions    = config.REGIONES,
        )

    def _run_etl(
        self, source_files: SourceFiles, app_config: AppConfig
    ):
        """Execute the full ETL pipeline and return ``(records, dias_info, presupuestos)``.

        This method is called from a background thread.

        Args:
            source_files: Paths to the 6 Odoo source files.
            app_config: Application configuration.

        Returns:
            Tuple ``(records, dias_info, presupuestos)`` or raises on error.
        """
        self._post("log", ("Leyendo archivo: Facturado…", "INFO"))
        df_facturado = read_facturado(source_files.facturado)

        self._post("log", ("Leyendo archivo: Pedidos…", "INFO"))
        df_pedidos = read_pedidos(source_files.pedidos)

        self._post("log", ("Leyendo archivo: Devoluciones…", "INFO"))
        df_devoluciones = read_devoluciones(source_files.devoluciones)

        self._post("log", ("Leyendo archivo: Cuentas por Cobrar…", "INFO"))
        df_cxc = read_por_cobrar(source_files.por_cobrar)

        self._post("log", ("Leyendo archivo: Pedidos por Facturar…", "INFO"))
        df_por_facturar = read_por_facturar(source_files.por_facturar)

        self._post("log", ("Leyendo archivo: Backorder…", "INFO"))
        df_backorder = read_backorder(source_files.backorder)

        self._post("log", ("Leyendo zonas y presupuestos…", "INFO"))
        zonas = read_zonas(str(self._zonas_config))
        presupuestos = load_presupuestos_from_config(self._presupuestos_config)

        if not presupuestos:
            self._post("log", (
                "⚠ No se encontraron presupuestos guardados. "
                "Cargue el archivo de presupuestos en la pestaña Mantenimiento.",
                "WARNING",
            ))

        self._post("log", ("Calculando días hábiles…", "INFO"))
        dias_info = calculate_working_days()

        # Load cobranza cierre if available
        cobranza_override = None
        if self._cobranza_cierre_path and os.path.isfile(self._cobranza_cierre_path):
            self._post(
                "log",
                (f"Aplicando Cobranza Cierre definitivo: {Path(self._cobranza_cierre_path).name}", "INFO"),
            )
            cobranza_override = load_cobranza_cierre(self._cobranza_cierre_path)
        else:
            self._post(
                "log",
                ("Cobranza: Calculando automáticamente desde archivo Odoo (mes en curso)…", "INFO"),
            )

        self._post("log", ("Calculando indicadores por vendedor…", "INFO"))
        records = calculate_general(
            zonas           = zonas,
            facturado       = df_facturado,
            pedidos         = df_pedidos,
            devoluciones    = df_devoluciones,
            cobranza        = df_cxc,
            por_facturar    = df_por_facturar,
            backorder       = df_backorder,
            presupuestos    = presupuestos,
            cobranza_override = cobranza_override,
        )
        records = calculate_rankings(records)

        self._post("log", (f"ETL completo: {len(records)} vendedores procesados.", "SUCCESS"))
        return records, dias_info, presupuestos

    # ------------------------------------------------------------------
    # Report action handlers
    # ------------------------------------------------------------------

    def _run_excel_reports(self) -> None:
        """Trigger Excel report generation in a background thread."""
        self._start_worker(self._worker_excel)

    def _run_pdf_reports(self) -> None:
        """Trigger PDF report generation in a background thread."""
        self._start_worker(self._worker_pdf)

    def _run_all_reports(self) -> None:
        """Trigger full report generation (Excel + PDF + manifests) in a background thread."""
        self._start_worker(self._worker_all)

    # --- Worker callables (run in background threads) ---

    def _worker_excel(self) -> None:
        """Background worker: generate Excel reports."""
        try:
            source_files = self._build_source_files()
            if source_files is None:
                return
            app_config  = self._build_app_config()
            output_dir  = Path(self._output_var.get())

            records, dias_info, _ = self._run_etl(source_files, app_config)

            self._post("log", ("Generando workbook Excel…", "INFO"))
            consolidated, regional_paths = generate_excel_report(
                records, app_config, dias_info, output_dir
            )
            self._post("done", (
                f"Reporte Excel generado:\n{consolidated}\n"
                f"Regionales: {len(regional_paths)} archivos en {output_dir / 'REGIONALES'}"
            ))
        except LynxBaseError as exc:
            self._post("error", str(exc))
        except Exception as exc:
            logger.exception("Unexpected error in Excel worker")
            self._post("error", f"Error inesperado: {exc}")

    def _worker_pdf(self) -> None:
        """Background worker: generate individual PDF reports."""
        try:
            source_files = self._build_source_files()
            if source_files is None:
                return
            app_config = self._build_app_config()
            output_dir = Path(self._output_var.get())

            records, _, _ = self._run_etl(source_files, app_config)

            self._post("log", ("Generando PDFs individuales…", "INFO"))
            pdf_paths = generate_pdf_reports(records, app_config, output_dir)
            self._post("done", (
                f"{len(pdf_paths)} PDFs generados en "
                f"{output_dir / 'PDF_INDICADORES'}"
            ))
        except LynxBaseError as exc:
            self._post("error", str(exc))
        except Exception as exc:
            logger.exception("Unexpected error in PDF worker")
            self._post("error", f"Error inesperado: {exc}")

    def _worker_all(self) -> None:
        """Background worker: generate all reports and manifests."""
        try:
            source_files = self._build_source_files()
            if source_files is None:
                return
            app_config = self._build_app_config()
            output_dir = Path(self._output_var.get())

            records, dias_info, _ = self._run_etl(source_files, app_config)

            self._post("log", ("Generando reporte Excel consolidado…", "INFO"))
            consolidated, regional_paths = generate_excel_report(
                records, app_config, dias_info, output_dir
            )

            self._post("log", ("Generando PDFs individuales…", "INFO"))
            pdf_paths = generate_pdf_reports(records, app_config, output_dir)

            self._post("log", ("Escribiendo manifiestos CSV…", "INFO"))
            vendor_manifest   = write_vendor_manifest(records, output_dir)
            regional_manifest = write_regional_manifest(regional_paths, output_dir)

            summary = (
                f"¡Generación completa!\n\n"
                f"• Excel consolidado: {consolidated.name}\n"
                f"• Regionales: {len(regional_paths)} archivos\n"
                f"• PDFs: {len(pdf_paths)} archivos\n"
                f"• Manifiestos: {vendor_manifest.name}, {regional_manifest.name}\n\n"
                f"Carpeta de salida: {output_dir}"
            )
            self._post("done", summary)

        except LynxBaseError as exc:
            self._post("error", str(exc))
        except Exception as exc:
            logger.exception("Unexpected error in all-reports worker")
            self._post("error", f"Error inesperado: {exc}")

    # ------------------------------------------------------------------
    # Maintenance handlers
    # ------------------------------------------------------------------

    def _load_presupuestos(self) -> None:
        """Load and persist the manager's budget file."""
        path = self._pres_picker.get()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("Archivo no válido", "Seleccione el archivo de presupuestos.")
            return

        def _worker():
            try:
                presupuestos = load_presupuestos(path)
                save_presupuestos_config(presupuestos, self._presupuestos_config)
                added = sync_zonas_with_presupuestos(self._zonas_config, presupuestos)
                msg = (
                    f"Presupuestos cargados: {len(presupuestos)} vendedores.\n"
                    f"Guardados en: {self._presupuestos_config}"
                )
                if added:
                    msg += f"\n\nNuevos vendedores registrados en zonas.xlsx ({len(added)}):\n• " + "\n• ".join(added)
                self._post("done", msg)
            except LynxBaseError as exc:
                self._post("error", str(exc))
            except Exception as exc:
                self._post("error", f"Error al cargar presupuestos: {exc}")

        self._start_worker(_worker)

    def _load_cobranza_cierre(self) -> None:
        """Validate and register the cobranza cierre file path."""
        path = self._cob_picker.get()
        if not path or not os.path.isfile(path):
            messagebox.showwarning(
                "Archivo no válido",
                "Seleccione el archivo INDICADOR_COBRANZA_*.xlsx."
            )
            return

        def _worker():
            try:
                cobranza = load_cobranza_cierre(path)
                saved_dest = save_cobranza_cierre(path)
                self._cobranza_cierre_path = str(saved_dest)
                self._post("done", (
                    f"Cobranza cierre cargada y guardada: {len(cobranza)} vendedores.\n"
                    f"Guardado en: {saved_dest.name}\n"
                    f"Se aplicará como cierre definitivo al próximo reporte generado."
                ))
                self._post(
                    "cob_status",
                    ("cierre", f"Modo: Cierre definitivo activo ({Path(path).name})"),
                )
            except LynxBaseError as exc:
                self._post("error", str(exc))
            except Exception as exc:
                self._post("error", f"Error al cargar cobranza cierre: {exc}")

        self._start_worker(_worker)

    def _clear_cobranza_cierre(self) -> None:
        """Clear active closing cobranza override, returning to dynamic Odoo mode."""
        self._cobranza_cierre_path = None
        self._cob_picker.set("")
        if hasattr(self, "_cob_status_lbl"):
            self._cob_status_lbl.config(
                text="Modo actual: Automático desde Odoo (mes en curso)",
                foreground="#006400",
            )
        self._log(
            "Cobranza Cierre desactivada. Se calculará automáticamente desde Odoo.",
            "INFO",
        )
        messagebox.showinfo(
            "Modo Automático",
            "Cobranza Cierre desactivada.\nLos próximos reportes calcularán % Sin Vencer desde los archivos de Odoo.",
        )

    def _on_email_done(self, stats: Dict[str, int]) -> None:
        """Callback ejecutado en el hilo principal cuando termina el envío masivo."""
        self._set_ui_busy(False)
        self._progressbar["value"] = 100
        self._progress_lbl.config(text="Envío de correos completado.")

        resumen = (
            f"Proceso de Envío Finalizado:\n\n"
            f"✔ Correos enviados con éxito: {stats.get('enviados', 0)}\n"
            f"✖ Envíos fallidos: {stats.get('fallidos', 0)}\n"
            f"⚠ Omitidos (sin correo o sin PDF): {stats.get('omitidos', 0)}"
        )
        self._log(resumen, "SUCCESS" if stats.get("fallidos", 0) == 0 else "WARNING")
        messagebox.showinfo("Envío de Correos Finalizado", resumen, parent=self)
