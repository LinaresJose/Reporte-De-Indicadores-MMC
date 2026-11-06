"""main.py — LYNX Report System.

Application entry point.

Responsibilities
----------------
1. Configure the Python :mod:`logging` framework:
   - File handler → ``output/lynx.log`` (rotating, UTF-8).
   - Stream handler → stdout (INFO and above).
2. Load the :class:`~core.models.AppConfig` from ``data/config.json`` if it
   exists (persists user preferences across runs: last output dir, mes_label).
3. Auto-detect the current month and pre-select it in the UI.
4. Instantiate and launch :class:`~gui.app.LynxApp`.

Usage::

    py main.py
"""

from __future__ import annotations

import json
import logging
import logging.handlers
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Project root on sys.path (allows ``import config``, ``from core import …``)
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).parent.resolve()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# ---------------------------------------------------------------------------
# Directory constants
# ---------------------------------------------------------------------------
DATA_DIR   = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"
LOG_FILE   = OUTPUT_DIR / "lynx.log"
CONFIG_FILE = DATA_DIR / "config.json"

# ---------------------------------------------------------------------------
# Default config
# ---------------------------------------------------------------------------
_MONTHS_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]


def _auto_mes_label() -> tuple[str, int]:
    """Detect the current month and year from the system clock.

    Returns:
        Tuple ``(month_name_es, year)`` e.g. ``("Noviembre", 2026)``.
    """
    now = datetime.now()
    return _MONTHS_ES[now.month - 1], now.year


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------


def configure_logging(log_file: Path) -> None:
    """Configure the application-wide logging framework.

    Sets up two handlers:

    - **RotatingFileHandler:** writes to *log_file* at DEBUG level, rotates at
      5 MB, keeps 3 backup files.
    - **StreamHandler:** writes to stdout at INFO level with a concise format.

    Args:
        log_file: Path to the log file.  Parent directory is created if needed.
    """
    log_file.parent.mkdir(parents=True, exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    fmt_file   = "%(asctime)s [%(levelname)-8s] %(name)s — %(message)s"
    fmt_stream = "[%(levelname)s] %(message)s"

    # File handler (rotating)
    fh = logging.handlers.RotatingFileHandler(
        log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(fmt_file, datefmt="%Y-%m-%d %H:%M:%S"))
    root_logger.addHandler(fh)

    # Stream handler (stdout)
    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.INFO)
    sh.setFormatter(logging.Formatter(fmt_stream))
    root_logger.addHandler(sh)

    logging.getLogger("PIL").setLevel(logging.WARNING)       # suppress pillow noise
    logging.getLogger("openpyxl").setLevel(logging.WARNING)  # suppress openpyxl noise


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------


def load_app_config() -> dict:
    """Load user preferences from ``data/config.json``.

    Returns:
        Dictionary with saved preferences, or an empty dict if the file does
        not exist or is malformed.
    """
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open(encoding="utf-8") as fh:
                return json.load(fh)
        except (json.JSONDecodeError, OSError) as exc:
            logging.warning("Could not load config.json: %s", exc)
    return {}


def save_app_config(cfg: dict) -> None:
    """Persist user preferences to ``data/config.json``.

    Args:
        cfg: Dictionary of preferences to save.
    """
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        with CONFIG_FILE.open("w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2, ensure_ascii=False)
    except OSError as exc:
        logging.warning("Could not save config.json: %s", exc)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Configure logging, load saved state, and launch the LYNX GUI."""
    configure_logging(LOG_FILE)
    logger = logging.getLogger(__name__)
    logger.info("=" * 60)
    logger.info("LYNX Sales Indicator System — starting up")
    logger.info("Project root: %s", PROJECT_ROOT)

    # Ensure required directories exist
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Load saved preferences
    saved_cfg = load_app_config()

    # Determine initial mes_label
    auto_mes, auto_anio = _auto_mes_label()
    saved_mes  = saved_cfg.get("mes",  auto_mes)
    saved_anio = saved_cfg.get("anio", auto_anio)
    saved_out  = saved_cfg.get("output_dir", str(OUTPUT_DIR))

    logger.info("Initial mes: %s %s", saved_mes, saved_anio)

    # Import GUI here (after sys.path is set up and logging is configured)
    try:
        from gui.app import LynxApp
    except ImportError as exc:
        logger.critical("Failed to import GUI: %s", exc)
        sys.exit(1)

    app = LynxApp(
        output_dir          = Path(saved_out),
        presupuestos_config = DATA_DIR / "presupuestos.xlsx",
        zonas_config        = DATA_DIR / "zonas.xlsx",
    )

    # Inject saved/auto-detected mes into the UI
    if hasattr(app, "_mes_var"):
        app._mes_var.set(saved_mes)
    if hasattr(app, "_anio_var"):
        app._anio_var.set(saved_anio)

    # Save preferences on close
    def _on_close():
        save_app_config({
            "mes":        app._mes_var.get() if hasattr(app, "_mes_var") else auto_mes,
            "anio":       app._anio_var.get() if hasattr(app, "_anio_var") else auto_anio,
            "output_dir": app._output_var.get() if hasattr(app, "_output_var") else str(OUTPUT_DIR),
        })
        logger.info("LYNX application closed.")
        app.destroy()

    app.protocol("WM_DELETE_WINDOW", _on_close)

    logger.info("Launching GUI event loop.")
    app.mainloop()


if __name__ == "__main__":
    main()
