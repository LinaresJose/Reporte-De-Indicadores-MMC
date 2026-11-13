"""gui/worker_dialog.py — Ventana de Gestión y Modificación de Correos de Trabajadores.

Permite buscar rápidamente cualquier trabajador (vendedor o coordinador),
ver sus metas del mes y actualizar de manera inmediata y persistente
su correo electrónico en data/zonas.xlsx.
"""

from __future__ import annotations

import logging
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Any, Callable, Dict, List, Optional

from services.worker_directory import (
    export_workers_directory_excel,
    get_workers_directory_data,
    update_worker_email,
)

logger = logging.getLogger(__name__)


class WorkerEmailDialog(tk.Toplevel):
    """Ventana emergente para consultar trabajadores, metas y modificar correos."""

    def __init__(
        self,
        parent: tk.Tk,
        zonas_path: Path,
        presupuestos_path: Path,
        mes_label: str = "",
        on_log: Optional[Callable[[str, str], None]] = None,
    ) -> None:
        super().__init__(parent)

        self.title("Directorio y Modificación de Correos — MMC")
        self.geometry("980x620")
        self.minsize(820, 500)
        self.transient(parent)
        self.grab_set()

        self._zonas_path = Path(zonas_path)
        self._presupuestos_path = Path(presupuestos_path)
        self._mes_label = mes_label
        self._on_log = on_log

        self._all_records: List[Dict[str, Any]] = []
        self._filtered_records: List[Dict[str, Any]] = []

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._on_search_change)

        self._selected_nombre_var = tk.StringVar(value="(Ningún trabajador seleccionado)")
        self._selected_role_var = tk.StringVar(value="")
        self._email_edit_var = tk.StringVar()

        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        """Construye los componentes visuales de la ventana."""
        # Top banner
        top_frame = ttk.Frame(self, padding=(12, 10))
        top_frame.pack(fill="x")

        ttk.Label(
            top_frame,
            text="Directorio de Trabajadores, Correos y Metas del Mes",
            font=("Segoe UI", 12, "bold"),
            foreground="#1F3864",
        ).pack(anchor="w")

        ttk.Label(
            top_frame,
            text="Seleccione un trabajador de la lista para modificar su correo electrónico o descargue el archivo consolidado en Excel.",
            font=("Segoe UI", 9),
            foreground="#555555",
        ).pack(anchor="w", pady=(2, 8))

        # Search bar row
        search_frame = ttk.Frame(top_frame)
        search_frame.pack(fill="x", pady=2)

        ttk.Label(search_frame, text="🔍 Buscar trabajador:", width=18, anchor="w").pack(side="left")
        search_entry = ttk.Entry(search_frame, textvariable=self._search_var, width=35)
        search_entry.pack(side="left", padx=(0, 10))
        search_entry.focus_set()

        ttk.Button(
            search_frame,
            text="Limpiar búsqueda",
            command=lambda: self._search_var.set(""),
            width=16,
        ).pack(side="left")

        self._stats_lbl = ttk.Label(
            search_frame,
            text="Cargando trabajadores...",
            font=("Segoe UI", 9, "bold"),
            foreground="#1F3864",
        )
        self._stats_lbl.pack(side="right")

        # Treeview / Table area
        tree_frame = ttk.Frame(self, padding=(12, 4))
        tree_frame.pack(fill="both", expand=True)

        columns = ("nombre", "rol", "zona", "region", "correo", "estado", "meta_ventas", "meta_act", "meta_skus")
        self._tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")

        self._tree.heading("nombre", text="Nombre del Trabajador", anchor="w")
        self._tree.heading("rol", text="Rol / Cargo", anchor="w")
        self._tree.heading("zona", text="Zona", anchor="center")
        self._tree.heading("region", text="Región", anchor="w")
        self._tree.heading("correo", text="Correo Electrónico Actual", anchor="w")
        self._tree.heading("estado", text="Estado Correo", anchor="center")
        self._tree.heading("meta_ventas", text="Meta Ventas ($)", anchor="e")
        self._tree.heading("meta_act", text="Meta Act.", anchor="e")
        self._tree.heading("meta_skus", text="Meta SKUs", anchor="e")

        self._tree.column("nombre", width=240, anchor="w")
        self._tree.column("rol", width=150, anchor="w")
        self._tree.column("zona", width=60, anchor="center")
        self._tree.column("region", width=100, anchor="w")
        self._tree.column("correo", width=220, anchor="w")
        self._tree.column("estado", width=110, anchor="center")
        self._tree.column("meta_ventas", width=110, anchor="e")
        self._tree.column("meta_act", width=80, anchor="e")
        self._tree.column("meta_skus", width=80, anchor="e")

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=scroll_y.set)
        self._tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Tags para colores
        self._tree.tag_configure("warn", background="#FCE4D6", foreground="#9C0006")
        self._tree.tag_configure("ok", background="#FFFFFF", foreground="#000000")
        self._tree.tag_configure("alt", background="#F7FAFC", foreground="#000000")

        # Edit panel at bottom
        edit_frame = ttk.LabelFrame(self, text="Modificar Correo del Trabajador Seleccionado", padding=(12, 8))
        edit_frame.pack(fill="x", padx=12, pady=(6, 12))

        info_row = ttk.Frame(edit_frame)
        info_row.pack(fill="x", pady=2)

        ttk.Label(info_row, text="Trabajador:", width=14, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(info_row, textvariable=self._selected_nombre_var, font=("Segoe UI", 9, "bold"), foreground="#1F3864").pack(side="left")
        ttk.Label(info_row, textvariable=self._selected_role_var, font=("Segoe UI", 9, "italic"), foreground="#555555").pack(side="left", padx=10)

        edit_row = ttk.Frame(edit_frame)
        edit_row.pack(fill="x", pady=(6, 4))

        ttk.Label(edit_row, text="Correo Destino:", width=14, anchor="w").pack(side="left")
        self._entry_email = ttk.Entry(edit_row, textvariable=self._email_edit_var, width=36)
        self._entry_email.pack(side="left", padx=(0, 10))
        self._entry_email.bind("<Return>", lambda event: self._save_email())

        self._btn_save = ttk.Button(
            edit_row,
            text="💾 Guardar Correo",
            command=self._save_email,
            width=18,
            style="Accent.TButton",
        )
        self._btn_save.pack(side="left", padx=(0, 15))

        self._save_status_lbl = ttk.Label(edit_row, text="", font=("Segoe UI", 9))
        self._save_status_lbl.pack(side="left")

        # Actions row (Excel and Close)
        actions_frame = ttk.Frame(self, padding=(12, 0, 12, 10))
        actions_frame.pack(fill="x")

        ttk.Button(
            actions_frame,
            text="📥 Descargar Directorio en Excel",
            command=self._export_excel,
            width=30,
        ).pack(side="left")

        ttk.Button(
            actions_frame,
            text="Cerrar",
            command=self.destroy,
            width=14,
        ).pack(side="right")

    def _load_data(self) -> None:
        """Carga los datos consolidados y puebla la tabla."""
        try:
            self._all_records = get_workers_directory_data(self._zonas_path, self._presupuestos_path)
            self._apply_filter()
        except Exception as exc:
            logger.error("Error al cargar directorio de trabajadores: %s", exc)
            messagebox.showerror("Error", f"No se pudo cargar el directorio de trabajadores:\n{exc}", parent=self)

    def _apply_filter(self) -> None:
        """Filtra y renderiza las filas del Treeview según el texto de búsqueda."""
        query = self._search_var.get().strip().upper()

        if query:
            self._filtered_records = [
                r for r in self._all_records
                if (
                    query in r["nombre"].upper()
                    or query in r["zona"].upper()
                    or query in r["region"].upper()
                    or query in r["correo"].upper()
                    or query in r["rol"].upper()
                )
            ]
        else:
            self._filtered_records = list(self._all_records)

        # Limpiar Treeview
        for item in self._tree.get_children():
            self._tree.delete(item)

        total_con_correo = sum(1 for r in self._filtered_records if r["tiene_correo"])
        total_sin_correo = len(self._filtered_records) - total_con_correo

        for idx, r in enumerate(self._filtered_records):
            tag = "warn" if not r["tiene_correo"] else ("alt" if idx % 2 == 0 else "ok")
            mv_str = f"${r['meta_ventas']:,.2f}" if r["meta_ventas"] > 0 else "-"
            ma_str = f"{int(r['meta_activacion'])}" if r["meta_activacion"] > 0 else "-"
            ms_str = f"{int(r['meta_skus'])}" if r["meta_skus"] > 0 else "-"

            self._tree.insert(
                "",
                "end",
                iid=r["nombre"],
                values=(
                    r["nombre"],
                    r["rol"],
                    r["zona"],
                    r["region"],
                    r["correo"] or "(Sin correo)",
                    r["estado_correo"],
                    mv_str,
                    ma_str,
                    ms_str,
                ),
                tags=(tag,),
            )

        self._stats_lbl.config(
            text=f"Total: {len(self._filtered_records)} | Con correo: {total_con_correo} | Sin correo: {total_sin_correo}"
        )

    def _on_search_change(self, *args) -> None:
        """Callback ejecutado cuando cambia el texto de búsqueda."""
        self._apply_filter()

    def _on_tree_select(self, event=None) -> None:
        """Callback al seleccionar una fila en la tabla."""
        selected = self._tree.selection()
        if not selected:
            return

        worker_name = selected[0]
        # Buscar en los registros
        match = next((r for r in self._all_records if r["nombre"] == worker_name), None)
        if match:
            self._selected_nombre_var.set(match["nombre"])
            self._selected_role_var.set(f"({match['rol']} — Zona {match['zona'] or 'N/A'}, {match['region']})")
            self._email_edit_var.set(match["correo"])
            self._save_status_lbl.config(text="")
            self._entry_email.focus_set()

    def _save_email(self) -> None:
        """Guarda el nuevo correo en zonas.xlsx y actualiza la tabla."""
        worker_name = self._selected_nombre_var.get()
        if not worker_name or worker_name == "(Ningún trabajador seleccionado)":
            messagebox.showwarning("Atención", "Por favor seleccione un trabajador de la lista.", parent=self)
            return

        new_email = self._email_edit_var.get().strip()

        ok, msg = update_worker_email(self._zonas_path, worker_name, new_email)
        if ok:
            self._save_status_lbl.config(text="✔ Guardado con éxito", foreground="#006400")
            if self._on_log:
                self._on_log(f"✔ {msg}", "SUCCESS")

            # Actualizar en memoria
            for r in self._all_records:
                if r["nombre"] == worker_name:
                    r["correo"] = new_email
                    r["tiene_correo"] = bool(new_email and "@" in new_email)
                    r["estado_correo"] = "Válido" if r["tiene_correo"] else "⚠ FALTA CORREO"
                    break

            # Re-aplicar filtro para actualizar la tabla
            self._apply_filter()
            # Restaurar selección
            if self._tree.exists(worker_name):
                self._tree.selection_set(worker_name)
                self._tree.see(worker_name)

            messagebox.showinfo("Correo Actualizado", msg, parent=self)
        else:
            self._save_status_lbl.config(text="✖ Error al guardar", foreground="#CC0000")
            if self._on_log:
                self._on_log(f"✖ {msg}", "ERROR")
            messagebox.showerror("Error", msg, parent=self)

    def _export_excel(self) -> None:
        """Exporta el directorio a un archivo Excel seleccionado por el usuario."""
        initial_name = f"Directorio_Trabajadores_Metas_{self._mes_label.replace(' ', '_')}.xlsx" if self._mes_label else "Directorio_Trabajadores_Metas.xlsx"
        dest = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar Directorio de Trabajadores en Excel",
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not dest:
            return

        try:
            exported_path = export_workers_directory_excel(
                zonas_path=self._zonas_path,
                presupuestos_path=self._presupuestos_path,
                output_file=Path(dest),
                mes_label=self._mes_label,
            )
            msg = f"Directorio de trabajadores guardado exitosamente en:\n{exported_path}"
            if self._on_log:
                self._on_log(f"✔ {msg}", "SUCCESS")
            messagebox.showinfo("Exportación Exitosa", msg, parent=self)
        except Exception as exc:
            err_msg = f"No se pudo exportar el directorio en Excel:\n{exc}"
            if self._on_log:
                self._on_log(f"✖ {err_msg}", "ERROR")
            messagebox.showerror("Error de Exportación", err_msg, parent=self)
