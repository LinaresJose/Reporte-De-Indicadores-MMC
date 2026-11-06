"""gui/widgets.py — LYNX Report System.

Reusable Tkinter composite widgets used by the main application window.

Widgets defined
---------------
- :class:`FilePickerRow`: Label + Entry + Browse-button + status icon in one row.
- :class:`LogTextArea`: Scrollable text widget with colour-coded log levels.
- :class:`SectionFrame`: Labelled LabelFrame with consistent padding and style.

All widgets inherit from ``ttk.Frame`` for seamless integration with the
themed Tkinter widget set.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, ttk
from typing import Callable, Optional


# ---------------------------------------------------------------------------
# Colour scheme (matches the Excel/PDF dark-navy palette)
# ---------------------------------------------------------------------------

_LOG_COLOURS = {
    "DEBUG":    "#888888",
    "INFO":     "#000000",
    "WARNING":  "#FF8C00",
    "ERROR":    "#CC0000",
    "CRITICAL": "#8B0000",
    "SUCCESS":  "#006400",
}

_STATUS_ICONS = {
    "empty":   "○",   # hollow circle — no file selected
    "ok":      "✔",   # check mark — file selected
    "error":   "✖",   # cross — invalid file
}

_STATUS_COLOURS = {
    "empty": "#AAAAAA",
    "ok":    "#006400",
    "error": "#CC0000",
}


# ---------------------------------------------------------------------------
# FilePickerRow
# ---------------------------------------------------------------------------


class FilePickerRow(ttk.Frame):
    """A composite widget for picking a single file.

    Layout (left → right)::

        [Label (fixed width)] [Entry (expandable)] [Browse] [Status icon]

    Args:
        parent: Parent Tkinter widget.
        label_text: Descriptive label shown to the left of the entry.
        file_types: List of ``(description, pattern)`` tuples passed to
            :func:`tkinter.filedialog.askopenfilename`.  Defaults to Excel files.
        label_width: Fixed character width of the label column.
        on_file_selected: Optional callback called with the selected file path
            whenever the user picks a file successfully.

    Attributes:
        path_var (tk.StringVar): Observable string variable holding the current
            file path.  Bind to it or read ``.get()`` to access the value.
    """

    def __init__(
        self,
        parent: tk.Widget,
        label_text: str,
        file_types: Optional[list[tuple[str, str]]] = None,
        label_width: int = 22,
        on_file_selected: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent)

        self._file_types      = file_types or [("Archivos Excel", "*.xlsx *.xls")]
        self._on_file_selected = on_file_selected

        # --- Path variable ------------------------------------------------
        self.path_var = tk.StringVar()
        self.path_var.trace_add("write", self._on_path_changed)

        # --- Label --------------------------------------------------------
        ttk.Label(self, text=label_text, width=label_width, anchor="w").pack(
            side="left", padx=(0, 4)
        )

        # --- Entry --------------------------------------------------------
        self._entry = ttk.Entry(self, textvariable=self.path_var, width=42)
        self._entry.pack(side="left", padx=(0, 4), fill="x", expand=True)

        # --- Browse button ------------------------------------------------
        ttk.Button(self, text="Examinar…", command=self._browse, width=10).pack(
            side="left", padx=(0, 4)
        )

        # --- Status icon --------------------------------------------------
        self._status_lbl = tk.Label(
            self,
            text=_STATUS_ICONS["empty"],
            fg=_STATUS_COLOURS["empty"],
            font=("Segoe UI", 11),
            width=2,
        )
        self._status_lbl.pack(side="left")

    # ------------------------------------------------------------------
    # Private methods
    # ------------------------------------------------------------------

    def _browse(self) -> None:
        """Open a file-picker dialog and populate the entry."""
        path = filedialog.askopenfilename(filetypes=self._file_types)
        if path:
            self.path_var.set(path)

    def _on_path_changed(self, *_args) -> None:
        """React to changes in :attr:`path_var` to update the status icon."""
        path = self.path_var.get().strip()
        if not path:
            self._set_status("empty")
        else:
            # Basic validation: file must exist and have an Excel extension
            import os
            if os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls", ".xlsm")):
                self._set_status("ok")
                if self._on_file_selected:
                    self._on_file_selected(path)
            else:
                self._set_status("error")

    def _set_status(self, status: str) -> None:
        """Update the status icon appearance.

        Args:
            status: One of ``"empty"``, ``"ok"``, ``"error"``.
        """
        self._status_lbl.config(
            text=_STATUS_ICONS[status],
            fg=_STATUS_COLOURS[status],
        )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get(self) -> str:
        """Return the current file path string (may be empty).

        Returns:
            The file path entered by the user, or an empty string.
        """
        return self.path_var.get().strip()

    def set(self, path: str) -> None:
        """Programmatically set the file path.

        Args:
            path: File path string to set.
        """
        self.path_var.set(path)


# ---------------------------------------------------------------------------
# LogTextArea
# ---------------------------------------------------------------------------


class LogTextArea(ttk.Frame):
    """A scrollable, read-only text area for displaying log messages.

    Log entries are colour-coded by severity level, matching the Python
    :mod:`logging` level names (DEBUG, INFO, WARNING, ERROR, CRITICAL).
    A special ``SUCCESS`` pseudo-level is supported for positive confirmations.

    Args:
        parent: Parent Tkinter widget.
        height: Number of text lines visible without scrolling.
        font: Font specification tuple (family, size).
    """

    def __init__(
        self,
        parent: tk.Widget,
        height: int = 10,
        font: tuple = ("Consolas", 9),
    ) -> None:
        super().__init__(parent)

        self._text = tk.Text(
            self,
            height=height,
            font=font,
            state="disabled",
            wrap="word",
            bg="#FAFAFA",
            relief="sunken",
            bd=1,
        )
        scrollbar = ttk.Scrollbar(self, command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)

        self._text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Configure colour tags
        for level, colour in _LOG_COLOURS.items():
            self._text.tag_configure(level, foreground=colour)

    def append(self, message: str, level: str = "INFO") -> None:
        """Append a log message to the text area.

        Args:
            message: The message text to display.
            level: Severity level string (``"DEBUG"``, ``"INFO"``, etc.).
                   Unknown levels are treated as ``"INFO"``.
        """
        tag = level if level in _LOG_COLOURS else "INFO"
        self._text.configure(state="normal")
        self._text.insert("end", message + "\n", tag)
        self._text.see("end")  # auto-scroll to bottom
        self._text.configure(state="disabled")

    def clear(self) -> None:
        """Remove all content from the log area."""
        self._text.configure(state="normal")
        self._text.delete("1.0", "end")
        self._text.configure(state="disabled")


# ---------------------------------------------------------------------------
# SectionFrame
# ---------------------------------------------------------------------------


class SectionFrame(ttk.LabelFrame):
    """A :class:`ttk.LabelFrame` with consistent padding and bold title.

    Args:
        parent: Parent Tkinter widget.
        title: Text shown in the frame border.
        pad: Internal padding in pixels applied to all four sides.
    """

    def __init__(self, parent: tk.Widget, title: str, pad: int = 8) -> None:
        super().__init__(parent, text=f"  {title}  ", padding=pad)
        self.columnconfigure(0, weight=1)
