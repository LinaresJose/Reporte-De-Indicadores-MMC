"""gui/worker_dialog.py — Ventana de Gestión y Modificación de Trabajadores (Correos y Rol de Coordinador).

Permite:
1. Buscar rápidamente cualquier trabajador (vendedor o coordinador).
2. Ver sus metas del mes, correo actual y región.
3. Modificar su correo electrónico de destino.
4. Asignar o cambiar el rol de un trabajador a Coordinador (o revertirlo a Vendedor)
   y definir su región para que le llegue automáticamente el reporte regional de su zona.
5. Descargar el directorio completo en Excel.
"""

from __future__ import annotations

import logging
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Any, Callable, Dict, List, Optional

import config
from services.worker_directory import (
    export_workers_directory_excel,
    get_workers_directory_data,
    update_worker_email,
    update_worker_role,
)

logger = logging.getLogger(__name__)


class WorkerEmailDialog(tk.Toplevel):
    """Ventana emergente para consultar trabajadores, metas y modificar correos y roles."""

    def __init__(
        self,
        parent: tk.Tk,
        zonas_path: Path,
        presupuestos_path: Path,
        mes_label: str = "",
        on_log: Optional[Callable[[str, str], None]] = None,
    ) -> None:
        super().__init__(parent)

        self.title("Directorio y Gestión de Trabajadores (Correos y Roles) — MMC")
        self.geometry("1020x670")
        self.minsize(880, 540)
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
        self._role_edit_var = tk.StringVar(value="Comercial / Vendedor")
        self._region_edit_var = tk.StringVar(value="CAPITAL")

        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        """Construye los componentes visuales de la ventana."""
        # Top banner
        top_frame = ttk.Frame(self, padding=(12, 10))
        top_frame.pack(fill="x")

        ttk.Label(
            top_frame,
            text="Directorio de Trabajadores, Correos y Asignación de Coordinadores",
            font=("Segoe UI", 12, "bold"),
            foreground="#1F3864",
        ).pack(anchor="w")

        ttk.Label(
            top_frame,
            text=(
                "Seleccione un trabajador para modificar su correo o asignar/cambiar su rol de Coordinador/Vendedor.\n"
                "A los Coordinadores se les envía automáticamente el Reporte Regional en Excel de la zona/región asignada."
            ),
            font=("Segoe UI", 9),
            foreground="#555555",
        ).pack(anchor="w", pady=(2, 8))

        # Search bar row
        search_frame = ttk.Frame(top_frame)
        search_frame.pack(fill="x", pady=2)

        ttk.Label(search_frame, text="🔍 Buscar trabajador:", width=18, anchor="w").pack(side="left")
        search_entry = ttk.Entry(search_frame, textvariable=self._search_var, width=32)
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
        self._tree.heading("region", text="Región Asignada", anchor="w")
        self._tree.heading("correo", text="Correo Electrónico Actual", anchor="w")
        self._tree.heading("estado", text="Estado Correo", anchor="center")
        self._tree.heading("meta_ventas", text="Meta Ventas ($)", anchor="e")
        self._tree.heading("meta_act", text="Meta Act.", anchor="e")
        self._tree.heading("meta_skus", text="Meta SKUs", anchor="e")

        self._tree.column("nombre", width=230, anchor="w")
        self._tree.column("rol", width=155, anchor="w")
        self._tree.column("zona", width=55, anchor="center")
        self._tree.column("region", width=115, anchor="w")
        self._tree.column("correo", width=210, anchor="w")
        self._tree.column("estado", width=105, anchor="center")
        self._tree.column("meta_ventas", width=105, anchor="e")
        self._tree.column("meta_act", width=75, anchor="e")
        self._tree.column("meta_skus", width=75, anchor="e")

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=scroll_y.set)
        self._tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Tags para colores
        self._tree.tag_configure("warn", background="#FCE4D6", foreground="#9C0006")
        self._tree.tag_configure("coord", background="#E8F4F8", foreground="#003366")
        self._tree.tag_configure("ok", background="#FFFFFF", foreground="#000000")
        self._tree.tag_configure("alt", background="#F7FAFC", foreground="#000000")

        # Edit panel at bottom
        edit_frame = ttk.LabelFrame(self, text="Editar Datos del Trabajador Seleccionado (Correo y Rol de Coordinador)", padding=(12, 8))
        edit_frame.pack(fill="x", padx=12, pady=(6, 10))

        info_row = ttk.Frame(edit_frame)
        info_row.pack(fill="x", pady=2)

        ttk.Label(info_row, text="Trabajador:", width=12, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(info_row, textvariable=self._selected_nombre_var, font=("Segoe UI", 9, "bold"), foreground="#1F3864").pack(side="left")
        ttk.Label(info_row, textvariable=self._selected_role_var, font=("Segoe UI", 9, "italic"), foreground="#555555").pack(side="left", padx=10)

        # Fields row: Correo, Rol, Región
        fields_row = ttk.Frame(edit_frame)
        fields_row.pack(fill="x", pady=(6, 4))

        # Correo
        ttk.Label(fields_row, text="Correo:", width=8, anchor="w").pack(side="left")
        self._entry_email = ttk.Entry(fields_row, textvariable=self._email_edit_var, width=28)
        self._entry_email.pack(side="left", padx=(0, 15))
        self._entry_email.bind("<Return>", lambda event: self._save_changes())

        # Rol Combobox
        ttk.Label(fields_row, text="Rol:", width=5, anchor="w").pack(side="left")
        self._role_combo = ttk.Combobox(
            fields_row,
            textvariable=self._role_edit_var,
            values=["Comercial / Vendedor", "Coordinador / Supervisor"],
            state="readonly",
            width=22,
        )
        self._role_combo.pack(side="left", padx=(0, 15))

        # Región Combobox (por si es coordinador o se cambia de región)
        ttk.Label(fields_row, text="Región:", width=8, anchor="w").pack(side="left")
        region_values = list(config.REGIONES)
        if "GERENCIA" not in region_values:
            region_values.append("GERENCIA")
        if "TVENTAS" not in region_values:
            region_values.append("TVENTAS")

        self._region_combo = ttk.Combobox(
            fields_row,
            textvariable=self._region_edit_var,
            values=region_values,
            state="readonly",
            width=16,
        )
        self._region_combo.pack(side="left", padx=(0, 15))

        # Botón Guardar
        self._btn_save = ttk.Button(
            fields_row,
            text="💾 Guardar Cambios",
            command=self._save_changes,
            width=18,
            style="Accent.TButton",
        )
        self._btn_save.pack(side="left", padx=(0, 10))

        self._save_status_lbl = ttk.Label(edit_frame, text="", font=("Segoe UI", 9))
        self._save_status_lbl.pack(anchor="w", pady=(2, 0))

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
        total_coords = sum(1 for r in self._filtered_records if "coord" in r["rol"].lower() or "superv" in r["rol"].lower())

        for idx, r in enumerate(self._filtered_records):
            is_c = "coord" in r["rol"].lower() or "superv" in r["rol"].lower()
            if not r["tiene_correo"]:
                tag = "warn"
            elif is_c:
                tag = "coord"
            else:
                tag = "alt" if idx % 2 == 0 else "ok"

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
            text=f"Total: {len(self._filtered_records)} | Coordinadores: {total_coords} | Con correo: {total_con_correo} | Sin correo: {total_sin_correo}"
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
        match = next((r for r in self._all_records if r["nombre"] == worker_name), None)
        if match:
            self._selected_nombre_var.set(match["nombre"])
            self._selected_role_var.set(f"(Zona {match['zona'] or 'N/A'}, {match['region']})")
            self._email_edit_var.set(match["correo"])

            is_c = "coord" in match["rol"].lower() or "superv" in match["rol"].lower()
            self._role_edit_var.set("Coordinador / Supervisor" if is_c else "Comercial / Vendedor")
            self._region_edit_var.set(match["region"] or "CAPITAL")

            self._save_status_lbl.config(text="")
            self._entry_email.focus_set()

    def _save_changes(self) -> None:
        """Guarda tanto el correo como el rol y región en data/zonas.xlsx."""
        worker_name = self._selected_nombre_var.get()
        if not worker_name or worker_name == "(Ningún trabajador seleccionado)":
            messagebox.showwarning("Atención", "Por favor seleccione un trabajador de la lista.", parent=self)
            return

        new_email = self._email_edit_var.get().strip()
        new_role = self._role_edit_var.get()
        new_region = self._region_edit_var.get()

        # 1. Guardar correo
        ok_email, msg_email = update_worker_email(self._zonas_path, worker_name, new_email)
        if not ok_email:
            self._save_status_lbl.config(text="✖ Error en correo", foreground="#CC0000")
            if self._on_log:
                self._on_log(f"✖ {msg_email}", "ERROR")
            messagebox.showerror("Error de Correo", msg_email, parent=self)
            return

        # 2. Guardar rol y región
        target_role = "Coordinador" if "coord" in new_role.lower() else "Vendedor"
        ok_role, msg_role = update_worker_role(
            self._zonas_path,
            worker_name,
            new_role=target_role,
            new_region=new_region,
        )
        if not ok_role:
            self._save_status_lbl.config(text="✖ Error en rol", foreground="#CC0000")
            if self._on_log:
                self._on_log(f"✖ {msg_role}", "ERROR")
            messagebox.showerror("Error de Rol", msg_role, parent=self)
            return

        # Éxito general
        success_msg = f"Cambios guardados para '{worker_name}':\n• Correo: {new_email or '(Sin correo)'}\n• Rol: {new_role}\n• Región: {new_region}"
        self._save_status_lbl.config(text=f"✔ Guardado: {new_role} ({new_region})", foreground="#006400")
        if self._on_log:
            self._on_log(f"✔ Actualizado {worker_name}: {new_role} [{new_region}] — {new_email}", "SUCCESS")

        # Actualizar registros en memoria
        for r in self._all_records:
            if r["nombre"] == worker_name:
                r["correo"] = new_email
                r["tiene_correo"] = bool(new_email and "@" in new_email)
                r["estado_correo"] = "Válido" if r["tiene_correo"] else "⚠ FALTA CORREO"
                r["rol"] = new_role
                r["region"] = new_region
                break

        # Re-aplicar filtro y mantener selección
        self._apply_filter()
        if self._tree.exists(worker_name):
            self._tree.selection_set(worker_name)
            self._tree.see(worker_name)

        messagebox.showinfo("Trabajador Actualizado", success_msg, parent=self)

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
