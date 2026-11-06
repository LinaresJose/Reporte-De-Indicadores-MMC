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
)
from maintenance.cobranza_loader import load_cobranza_cierre
from reports.excel_writer import generate_excel_report
from reports.manifest_writer import write_regional_manifest, write_vendor_manifest
from reports.pdf_writer import generate_pdf_reports

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
        self._tab_maint     = ttk.Frame(notebook, padding=10)

        notebook.add(self._tab_import,  text="📂  Importar Datos")
        notebook.add(self._tab_reports, text="📊  Generar Reportes")
        notebook.add(self._tab_maint,   text="🔧  Mantenimiento")

        self._build_import_tab()
        self._build_reports_tab()
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

        ttk.Button(
            cob_frame,
            text="Cargar Cobranza Cierre",
            command=self._load_cobranza_cierre,
            width=25,
        ).pack(anchor="w", pady=(6, 0))

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------

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
        # If called from the main thread, write directly
        try:
            self._log_area.append(message, level)
        except Exception:
            self._gui_queue.put(_Msg("log", (message, level)))

    # ------------------------------------------------------------------
    # State persistence
    # ------------------------------------------------------------------

    def _load_saved_state(self) -> None:
        """Load previously saved presupuestos config if available."""
        if self._presupuestos_config.exists():
            self._log(
                f"Presupuestos guardados encontrados: {self._presupuestos_config.name}",
                "INFO",
            )

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
                    self._progress_lbl.config(text=msg.payload)
                elif msg.kind == "done":
                    self._on_worker_done(msg.payload)
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
            self._post("log", ("Aplicando Cobranza Cierre…", "INFO"))
            cobranza_override = load_cobranza_cierre(self._cobranza_cierre_path)

        self._post("log", ("Calculando indicadores por vendedor…", "INFO"))
        records = calculate_general(
            df_facturado    = df_facturado,
            df_pedidos      = df_pedidos,
            df_devoluciones = df_devoluciones,
            df_cxc          = df_cxc,
            df_por_facturar = df_por_facturar,
            df_backorder    = df_backorder,
            zonas           = zonas,
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
                self._post("done", (
                    f"Presupuestos cargados: {len(presupuestos)} vendedores.\n"
                    f"Guardados en: {self._presupuestos_config}"
                ))
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
                self._cobranza_cierre_path = path
                self._post("done", (
                    f"Cobranza cierre cargada: {len(cobranza)} vendedores.\n"
                    f"Se aplicará al próximo reporte generado."
                ))
            except LynxBaseError as exc:
                self._post("error", str(exc))
            except Exception as exc:
                self._post("error", f"Error al cargar cobranza cierre: {exc}")

        self._start_worker(_worker)
