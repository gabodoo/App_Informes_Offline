"""
Generador de Informes Offline - punto de entrada de la aplicacion.
"""

import sys
import traceback
from pathlib import Path


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def main() -> None:
    base_dir = get_base_dir()
    (base_dir / "plantillas").mkdir(parents=True, exist_ok=True)
    (base_dir / "Informes_Generados").mkdir(parents=True, exist_ok=True)

    from gui import AppInformes
    app = AppInformes()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        base_dir = get_base_dir()
        log_path = base_dir / "error_log.txt"
        error_text = traceback.format_exc()
        log_path.write_text(error_text, encoding="utf-8")
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Error al iniciar",
                f"La aplicacion fallo.\n\nLog en: {log_path}\n\n{error_text[:600]}"
            )
        except Exception:
            pass
