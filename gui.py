"""
Interfaz grafica del Generador de Libres (TECNORDI SA).

Pensada para usuarios sin conocimientos tecnicos:
  1. Elegir el Excel del manifiesto.
  2. Ingresar la fecha de llegada del barco.
  3. Apretar "Generar libres".

Los PDF resultantes se guardan automaticamente en la carpeta Descargas.
Los de deposito 1716 usan el numero de BL; los de 1714 (contenedor)
usan el BL seguido de CNT.
"""

from __future__ import annotations

import json
import os
import re
import threading
import webbrowser
from pathlib import Path
from tkinter import Tk, StringVar, BooleanVar, filedialog, messagebox, ttk, END, DISABLED, NORMAL, WORD
from tkinter.scrolledtext import ScrolledText

from libre_core import (
    LibreGeneratorError,
    carpeta_descargas,
    encontrar_wkhtmltopdf,
    generar_pdfs,
    leer_libres_desde_excel,
)

APP_TITLE = "Generador de Libres - TECNORDI SA"
CONFIG_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / "GeneradorDeLibres"
CONFIG_PATH = CONFIG_DIR / "config.json"

WKHTMLTOPDF_DESCARGA_URL = "https://wkhtmltopdf.org/downloads.html"

COLOR_FONDO = "#f4f6f8"
COLOR_PRIMARIO = "#1f3864"
COLOR_TEXTO = "#1f2937"


def _cargar_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _guardar_config(cfg: dict) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass  # No es critico si no se puede guardar la config.


class GeneradorLibresApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("620x600")
        self.root.minsize(560, 560)
        self.root.configure(bg=COLOR_FONDO)

        self.config_usuario = _cargar_config()
        self.ruta_excel_var = StringVar(value="")
        self.fecha_llegada_var = StringVar(value="")
        self.incluir_fechas_var = BooleanVar(value=True)
        self.ruta_wkhtmltopdf = self.config_usuario.get("ruta_wkhtmltopdf") or encontrar_wkhtmltopdf()

        self._construir_ui()
        self._actualizar_estado_wkhtmltopdf()

    # ------------------------------------------------------------------
    # Construccion de la interfaz
    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except Exception:
            pass
        estilo.configure("TFrame", background=COLOR_FONDO)
        estilo.configure("TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=("Segoe UI", 10))
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 15, "bold"), foreground=COLOR_PRIMARIO)
        estilo.configure("Paso.TLabel", font=("Segoe UI", 10, "bold"), foreground=COLOR_PRIMARIO)
        estilo.configure("Ayuda.TLabel", font=("Segoe UI", 8), foreground="#6b7280")
        estilo.configure("Estado.TLabel", font=("Segoe UI", 8))
        estilo.configure("TButton", font=("Segoe UI", 10), padding=6)
        estilo.configure(
            "Generar.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=10,
        )

        contenedor = ttk.Frame(self.root, padding=20)
        contenedor.pack(fill="both", expand=True)

        ttk.Label(contenedor, text="Generador de Libres", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(
            contenedor,
            text="Crea automaticamente los PDF de libre para las lineas con deposito 1716 y 1714.",
            style="Ayuda.TLabel",
        ).pack(anchor="w", pady=(0, 16))

        # Paso 1: Excel
        ttk.Label(contenedor, text="1. Archivo Excel del manifiesto", style="Paso.TLabel").pack(anchor="w")
        marco_excel = ttk.Frame(contenedor)
        marco_excel.pack(fill="x", pady=(4, 4))
        self.entrada_excel = ttk.Entry(marco_excel, textvariable=self.ruta_excel_var, state="readonly")
        self.entrada_excel.pack(side="left", fill="x", expand=True, ipady=3)
        ttk.Button(marco_excel, text="Seleccionar...", command=self._elegir_excel).pack(side="left", padx=(8, 0))

        ttk.Label(contenedor, text="", style="Ayuda.TLabel").pack(anchor="w", pady=(0, 10))

        # Paso 2: fecha
        ttk.Label(contenedor, text="2. Llegada del barco", style="Paso.TLabel").pack(anchor="w")
        marco_fecha = ttk.Frame(contenedor)
        marco_fecha.pack(fill="x", pady=(4, 2))
        self.entrada_fecha = ttk.Entry(marco_fecha, textvariable=self.fecha_llegada_var, width=16)
        self.entrada_fecha.pack(side="left", ipady=3)
        ttk.Label(marco_fecha, text="  Formato: DD/MM/AAAA (ej: 13/06/2026)", style="Ayuda.TLabel").pack(
            side="left"
        )
        ttk.Checkbutton(
            contenedor,
            text="Incluir fechas en el PDF",
            variable=self.incluir_fechas_var,
            command=self._al_cambiar_incluir_fechas,
        ).pack(anchor="w", pady=(6, 0))
        ttk.Label(
            contenedor,
            text="Si lo desmarcás, el PDF sale sin llegada del barco ni fecha de vencimiento.",
            style="Ayuda.TLabel",
        ).pack(anchor="w", pady=(0, 10))

        # Paso 3: generar
        ttk.Label(contenedor, text="3. Generar los libres", style="Paso.TLabel").pack(anchor="w", pady=(0, 6))
        self.boton_generar = ttk.Button(
            contenedor, text="Generar libres", style="Generar.TButton", command=self._al_generar
        )
        self.boton_generar.pack(fill="x")

        ttk.Label(
            contenedor,
            text="Los PDF se guardan en Descargas: 1716 como BL.pdf y 1714 (contenedor) como BL CNT.pdf.",
            style="Ayuda.TLabel",
        ).pack(anchor="w", pady=(6, 12))

        # Barra de progreso
        self.progreso = ttk.Progressbar(contenedor, mode="determinate")
        self.progreso.pack(fill="x", pady=(0, 10))

        # Registro de actividad
        ttk.Label(contenedor, text="Actividad", style="Paso.TLabel").pack(anchor="w")
        self.registro = ScrolledText(
            contenedor, height=10, wrap=WORD, font=("Consolas", 9), state=DISABLED, bg="white"
        )
        self.registro.pack(fill="both", expand=True, pady=(4, 10))

        # Estado de wkhtmltopdf
        marco_estado = ttk.Frame(contenedor)
        marco_estado.pack(fill="x")
        self.etiqueta_estado_wk = ttk.Label(marco_estado, text="", style="Estado.TLabel")
        self.etiqueta_estado_wk.pack(side="left")
        ttk.Button(
            marco_estado, text="Configurar wkhtmltopdf...", command=self._configurar_wkhtmltopdf
        ).pack(side="right")

    # ------------------------------------------------------------------
    # Registro de actividad en pantalla
    # ------------------------------------------------------------------
    def _log(self, mensaje: str) -> None:
        self.registro.configure(state=NORMAL)
        self.registro.insert(END, mensaje + "\n")
        self.registro.see(END)
        self.registro.configure(state=DISABLED)

    def _limpiar_log(self) -> None:
        self.registro.configure(state=NORMAL)
        self.registro.delete("1.0", END)
        self.registro.configure(state=DISABLED)

    # ------------------------------------------------------------------
    # wkhtmltopdf
    # ------------------------------------------------------------------
    def _actualizar_estado_wkhtmltopdf(self) -> None:
        if self.ruta_wkhtmltopdf:
            self.etiqueta_estado_wk.configure(
                text=f"wkhtmltopdf: {self.ruta_wkhtmltopdf}", foreground="#15803d"
            )
        else:
            self.etiqueta_estado_wk.configure(
                text="wkhtmltopdf no encontrado. Configuralo con el boton de la derecha.",
                foreground="#b91c1c",
            )

    def _configurar_wkhtmltopdf(self) -> None:
        respuesta = messagebox.askyesno(
            "wkhtmltopdf",
            "wkhtmltopdf es un programa gratuito necesario para crear los PDF.\n\n"
            "Si todavia no lo instalaste, se abrira la pagina de descarga.\n"
            "Si ya lo instalaste, elegi 'No' y selecciona el archivo wkhtmltopdf.exe.\n\n"
            "Abrir la pagina de descarga?",
        )
        if respuesta:
            webbrowser.open(WKHTMLTOPDF_DESCARGA_URL)
            return

        ruta = filedialog.askopenfilename(
            title="Selecciona wkhtmltopdf.exe",
            filetypes=[("Ejecutable", "wkhtmltopdf.exe"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self.ruta_wkhtmltopdf = ruta
            self.config_usuario["ruta_wkhtmltopdf"] = ruta
            _guardar_config(self.config_usuario)
            self._actualizar_estado_wkhtmltopdf()
            self._log(f"Ruta de wkhtmltopdf configurada: {ruta}")

    # ------------------------------------------------------------------
    # Seleccion de archivo Excel
    # ------------------------------------------------------------------
    def _elegir_excel(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel del manifiesto",
            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm *.xls"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self.ruta_excel_var.set(ruta)

    # ------------------------------------------------------------------
    # Validaciones
    # ------------------------------------------------------------------
    @staticmethod
    def _fecha_valida(texto: str) -> bool:
        return bool(re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", texto.strip()))

    # ------------------------------------------------------------------
    # Generacion (en un hilo aparte para no congelar la ventana)
    # ------------------------------------------------------------------
    def _al_cambiar_incluir_fechas(self) -> None:
        if self.incluir_fechas_var.get():
            self.entrada_fecha.configure(state=NORMAL)
        else:
            self.entrada_fecha.configure(state=DISABLED)

    def _al_generar(self) -> None:
        ruta_excel = self.ruta_excel_var.get().strip()
        incluir_fechas = bool(self.incluir_fechas_var.get())
        fecha_llegada = self.fecha_llegada_var.get().strip() if incluir_fechas else ""

        if not ruta_excel:
            messagebox.showwarning(APP_TITLE, "Primero selecciona el archivo Excel del manifiesto.")
            return
        if incluir_fechas and (not fecha_llegada or not self._fecha_valida(fecha_llegada)):
            messagebox.showwarning(
                APP_TITLE, "Ingresa la fecha de llegada del barco con formato DD/MM/AAAA."
            )
            return
        if not self.ruta_wkhtmltopdf:
            messagebox.showerror(
                APP_TITLE,
                "No se encontro wkhtmltopdf, que es necesario para crear los PDF.\n"
                "Usa el boton 'Configurar wkhtmltopdf...' para instalarlo o indicar su ubicacion.",
            )
            return

        self.boton_generar.configure(state=DISABLED)
        self._limpiar_log()
        self.progreso.configure(value=0, maximum=100)
        self._log("Leyendo el Excel...")

        hilo = threading.Thread(
            target=self._ejecutar_generacion,
            args=(ruta_excel, fecha_llegada, incluir_fechas),
            daemon=True,
        )
        hilo.start()

    def _ejecutar_generacion(self, ruta_excel: str, fecha_llegada: str, incluir_fechas: bool) -> None:
        try:
            libres = leer_libres_desde_excel(ruta_excel)
            total = len(libres)
            n_1716 = sum(1 for libre in libres if not libre.es_contenedor)
            n_1714 = sum(1 for libre in libres if libre.es_contenedor)
            self.root.after(
                0,
                self._log,
                f"Se armaron {total} libre(s): {n_1716} de retiro TMM (1716) y {n_1714} de contenedor (1714).",
            )
            self.root.after(0, self.progreso.configure, {"maximum": total, "value": 0})
            self.root.after(
                0,
                self._log,
                "Generando con fechas." if incluir_fechas else "Generando sin fechas.",
            )

            destino = carpeta_descargas()

            def progreso(indice: int, total_: int, nombre: str) -> None:
                self.root.after(0, self.progreso.configure, {"value": indice})
                self.root.after(0, self._log, f"[{indice}/{total_}] Generado: {nombre}")

            generados = generar_pdfs(
                libres,
                fecha_llegada=fecha_llegada,
                carpeta_destino=destino,
                ruta_wkhtmltopdf=self.ruta_wkhtmltopdf,
                incluir_fechas=incluir_fechas,
                on_progreso=progreso,
            )

            self.root.after(0, self._al_finalizar_ok, len(generados), str(destino))
        except LibreGeneratorError as exc:
            self.root.after(0, self._al_finalizar_error, str(exc))
        except Exception as exc:  # error inesperado, igual mostramos algo claro
            self.root.after(0, self._al_finalizar_error, f"Ocurrio un error inesperado: {exc}")

    def _al_finalizar_ok(self, cantidad: int, carpeta: str) -> None:
        self._log(f"Listo. Se generaron {cantidad} libre(s) en: {carpeta}")
        self.boton_generar.configure(state=NORMAL)
        messagebox.showinfo(
            APP_TITLE, f"Se generaron {cantidad} libre(s) en tu carpeta Descargas."
        )

    def _al_finalizar_error(self, mensaje: str) -> None:
        self._log(f"ERROR: {mensaje}")
        self.boton_generar.configure(state=NORMAL)
        messagebox.showerror(APP_TITLE, mensaje)


def iniciar_app() -> None:
    root = Tk()
    GeneradorLibresApp(root)
    root.mainloop()
