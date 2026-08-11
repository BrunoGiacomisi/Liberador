"""Punto de entrada del Generador de Libres.

Ejecutar con:  python main.py
(o el .exe generado con PyInstaller, ver build.bat / README.md)
"""

if __name__ == "__main__":
    try:
        from gui import iniciar_app

        iniciar_app()
    except Exception as exc:  # ultimo recurso: mostrar el error igual sin consola
        try:
            import tkinter.messagebox as messagebox

            messagebox.showerror(
                "Generador de Libres",
                f"No se pudo iniciar la aplicacion:\n\n{exc}",
            )
        except Exception:
            raise
