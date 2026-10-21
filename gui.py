"""
Interfaz grafica del Generador de Libres (TECNORDI SA).

Pensada para usuarios sin conocimientos tecnicos:
  1. Elegir el Excel de la planilla de facturacion.
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
from tkinter import (
    Tk,
    Toplevel,
    Frame,
    Label,
    StringVar,
    BooleanVar,
    PhotoImage,
    Checkbutton,
    filedialog,
    messagebox,
    ttk,
    END,
    DISABLED,
    NORMAL,
    WORD,
)
from tkinter.scrolledtext import ScrolledText

from libre_core import (
    CondicionesAlmacenaje,
    LibreGeneratorError,
    carpeta_descargas,
    condiciones_desde_dict,
    encontrar_wkhtmltopdf,
    generar_pdfs,
    leer_libres_desde_excel,
    filtrar_libres_por_bl,
)

APP_TITLE = "Generador de Libres - TECNORDI SA"
CONFIG_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / "GeneradorDeLibres"
CONFIG_PATH = CONFIG_DIR / "config.json"

WKHTMLTOPDF_DESCARGA_URL = "https://wkhtmltopdf.org/downloads.html"

COLOR_FONDO = "#eef1f5"
COLOR_TARJETA = "#ffffff"
COLOR_BORDE = "#dbe1e8"
COLOR_PRIMARIO = "#4a6fa5"
COLOR_PRIMARIO_HOVER = "#3f6194"
COLOR_PRIMARIO_DESACTIVADO = "#aebdd3"
COLOR_SUBTITULO = "#e1e8f3"
COLOR_TEXTO = "#1f2937"
COLOR_TEXTO_SUAVE = "#6b7280"
COLOR_TILDE = "#15803d"
FUENTE = "Segoe UI"


def _imagen_casilla(marcada: bool) -> PhotoImage:
    """Casilla de 18 px: vacia, o con tilde verde cuando esta marcada."""
    lado = 18
    img = PhotoImage(width=lado, height=lado)
    img.put("#ffffff", to=(0, 0, lado, lado))
    borde = COLOR_TILDE if marcada else "#9ca3af"
    img.put(borde, to=(0, 0, lado, 2))
    img.put(borde, to=(0, lado - 2, lado, lado))
    img.put(borde, to=(0, 0, 2, lado))
    img.put(borde, to=(lado - 2, 0, lado, lado))
    if not marcada:
        return img
    # Trazo grueso: baja a la derecha y sube hasta la punta.
    puntos = [
        (3, 8), (4, 9), (5, 10), (6, 11), (7, 12),
        (3, 9), (4, 10), (5, 11), (6, 12), (7, 13),
        (8, 12), (9, 11), (10, 10), (11, 9), (12, 8), (13, 7), (14, 6),
        (8, 11), (9, 10), (10, 9), (11, 8), (12, 7), (13, 6), (14, 5),
    ]
    for x, y in puntos:
        img.put(COLOR_TILDE, to=(x, y, x + 1, y + 1))
    return img


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
        self.root.geometry("640x660")
        self.root.minsize(580, 620)
        self.root.configure(bg=COLOR_FONDO)

        self.config_usuario = _cargar_config()
        self.ruta_excel_var = StringVar(value="")
        self.fecha_llegada_var = StringVar(value="")
        self.bl_var = StringVar(value="")
        self.incluir_fechas_var = BooleanVar(value=True)
        self.ruta_wkhtmltopdf = self.config_usuario.get("ruta_wkhtmltopdf") or encontrar_wkhtmltopdf()
        self.condiciones = condiciones_desde_dict(self.config_usuario.get("condiciones_almacenaje"))

        self._construir_ui()
        self._actualizar_estado_wkhtmltopdf()

    # ------------------------------------------------------------------
    # Construccion de la interfaz
    # ------------------------------------------------------------------
    def _configurar_estilos(self) -> None:
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except Exception:
            pass

        estilo.configure(".", font=(FUENTE, 10))
        estilo.configure("TFrame", background=COLOR_FONDO)
        estilo.configure("Tarjeta.TFrame", background=COLOR_TARJETA)
        estilo.configure("TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=(FUENTE, 10))
        estilo.configure("Tarjeta.TLabel", background=COLOR_TARJETA, foreground=COLOR_TEXTO)
        estilo.configure(
            "Codigo.TLabel", background=COLOR_TARJETA, foreground=COLOR_PRIMARIO, font=(FUENTE, 10, "bold")
        )
        estilo.configure(
            "Paso.TLabel", background=COLOR_TARJETA, foreground=COLOR_PRIMARIO, font=(FUENTE, 11, "bold")
        )
        estilo.configure(
            "Ayuda.TLabel", background=COLOR_TARJETA, foreground=COLOR_TEXTO_SUAVE, font=(FUENTE, 9)
        )
        estilo.configure(
            "Seccion.TLabel", background=COLOR_FONDO, foreground=COLOR_PRIMARIO, font=(FUENTE, 10, "bold")
        )
        estilo.configure("Pie.TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO_SUAVE, font=(FUENTE, 8))

        estilo.configure(
            "TEntry",
            padding=(8, 5),
            fieldbackground=COLOR_TARJETA,
            bordercolor=COLOR_BORDE,
            lightcolor=COLOR_BORDE,
            darkcolor=COLOR_BORDE,
        )
        estilo.map(
            "TEntry",
            bordercolor=[("focus", COLOR_PRIMARIO)],
            lightcolor=[("focus", COLOR_PRIMARIO)],
            fieldbackground=[("disabled", "#f3f4f6"), ("readonly", "#f8fafc")],
            foreground=[("disabled", "#9ca3af")],
        )

        estilo.configure(
            "TButton",
            padding=(12, 6),
            background=COLOR_TARJETA,
            foreground=COLOR_TEXTO,
            bordercolor=COLOR_BORDE,
            lightcolor=COLOR_TARJETA,
            darkcolor=COLOR_TARJETA,
            focuscolor=COLOR_TARJETA,
        )
        estilo.map(
            "TButton",
            background=[("disabled", "#f3f4f6"), ("pressed", "#e5e7eb"), ("active", "#f1f4f8")],
            lightcolor=[("active", "#f1f4f8")],
            darkcolor=[("active", "#f1f4f8")],
            foreground=[("disabled", "#9ca3af")],
        )

        for nombre, fuente, relleno in (
            ("Primario.TButton", (FUENTE, 12, "bold"), (12, 8)),
            ("Accion.TButton", (FUENTE, 10, "bold"), (14, 6)),
        ):
            estilo.configure(
                nombre,
                font=fuente,
                padding=relleno,
                background=COLOR_PRIMARIO,
                foreground="white",
                bordercolor=COLOR_PRIMARIO,
                lightcolor=COLOR_PRIMARIO,
                darkcolor=COLOR_PRIMARIO,
                focuscolor=COLOR_PRIMARIO,
            )
            estilo.map(
                nombre,
                background=[("disabled", COLOR_PRIMARIO_DESACTIVADO), ("active", COLOR_PRIMARIO_HOVER)],
                bordercolor=[("disabled", COLOR_PRIMARIO_DESACTIVADO), ("active", COLOR_PRIMARIO_HOVER)],
                lightcolor=[("disabled", COLOR_PRIMARIO_DESACTIVADO), ("active", COLOR_PRIMARIO_HOVER)],
                darkcolor=[("disabled", COLOR_PRIMARIO_DESACTIVADO), ("active", COLOR_PRIMARIO_HOVER)],
                foreground=[("disabled", "white")],
            )

        estilo.configure(
            "Enlace.TButton",
            font=(FUENTE, 8, "underline"),
            padding=(4, 2),
            background=COLOR_FONDO,
            foreground=COLOR_PRIMARIO,
            bordercolor=COLOR_FONDO,
            lightcolor=COLOR_FONDO,
            darkcolor=COLOR_FONDO,
            focuscolor=COLOR_FONDO,
        )
        estilo.map(
            "Enlace.TButton",
            background=[("active", COLOR_FONDO)],
            foreground=[("active", COLOR_PRIMARIO_HOVER)],
        )

        estilo.configure(
            "Horizontal.TProgressbar",
            thickness=6,
            troughcolor="#e2e7ee",
            background=COLOR_PRIMARIO,
            bordercolor="#e2e7ee",
            lightcolor=COLOR_PRIMARIO,
            darkcolor=COLOR_PRIMARIO,
        )
        estilo.configure("Tarjeta.TSeparator", background=COLOR_BORDE)

    def _tarjeta(self, padre, numero: int, titulo: str) -> tuple[ttk.Frame, ttk.Frame]:
        """Seccion blanca con borde fino y titulo numerado. Devuelve (cuerpo, cabecera)."""
        borde = Frame(padre, bg=COLOR_BORDE)
        borde.pack(fill="x", pady=(0, 10))
        cuerpo = ttk.Frame(borde, style="Tarjeta.TFrame", padding=(16, 10, 16, 12))
        cuerpo.pack(fill="both", expand=True, padx=1, pady=1)

        cabecera = ttk.Frame(cuerpo, style="Tarjeta.TFrame")
        cabecera.pack(fill="x", pady=(0, 8))
        Label(
            cabecera,
            text=str(numero),
            bg=COLOR_PRIMARIO,
            fg="white",
            font=(FUENTE, 9, "bold"),
            width=2,
        ).pack(side="left")
        ttk.Label(cabecera, text=titulo, style="Paso.TLabel").pack(side="left", padx=(10, 0))
        return cuerpo, cabecera

    def _construir_ui(self) -> None:
        self._configurar_estilos()

        # Cabecera
        cabecera = Frame(self.root, bg=COLOR_PRIMARIO)
        cabecera.pack(fill="x")
        Label(
            cabecera,
            text="Generador de Libres",
            bg=COLOR_PRIMARIO,
            fg="white",
            font=(FUENTE, 15, "bold"),
        ).pack(anchor="w", padx=22, pady=(10, 0))
        Label(
            cabecera,
            text="TECNORDI SA  ·  PDF de libre para depósitos 1716 y 1714",
            bg=COLOR_PRIMARIO,
            fg=COLOR_SUBTITULO,
            font=(FUENTE, 9),
        ).pack(anchor="w", padx=22, pady=(0, 10))

        contenedor = ttk.Frame(self.root, padding=(18, 12, 18, 8))
        contenedor.pack(fill="both", expand=True)

        # Paso 1: Excel
        paso1, _ = self._tarjeta(contenedor, 1, "Archivo Excel de la planilla de facturación")
        marco_excel = ttk.Frame(paso1, style="Tarjeta.TFrame")
        marco_excel.pack(fill="x")
        self.entrada_excel = ttk.Entry(marco_excel, textvariable=self.ruta_excel_var, state="readonly")
        self.entrada_excel.pack(side="left", fill="x", expand=True)
        ttk.Button(marco_excel, text="Seleccionar...", command=self._elegir_excel).pack(
            side="left", padx=(8, 0)
        )

        # Paso 2: fecha
        paso2, cabecera_paso2 = self._tarjeta(contenedor, 2, "Llegada del barco")
        ttk.Button(
            cabecera_paso2, text="Días de almacenaje", command=self._editar_condiciones
        ).pack(side="right")
        marco_fecha = ttk.Frame(paso2, style="Tarjeta.TFrame")
        marco_fecha.pack(fill="x")
        self.entrada_fecha = ttk.Entry(
            marco_fecha, textvariable=self.fecha_llegada_var, width=14, font=(FUENTE, 10)
        )
        self.entrada_fecha.pack(side="left")
        ttk.Label(marco_fecha, text="DD/MM/AAAA  (ej: 13/06/2026)", style="Ayuda.TLabel").pack(
            side="left", padx=(10, 0)
        )
        self._img_casilla_off = _imagen_casilla(False)
        self._img_casilla_on = _imagen_casilla(True)
        Checkbutton(
            paso2,
            text="  Incluir fechas en el PDF",
            variable=self.incluir_fechas_var,
            command=self._al_cambiar_incluir_fechas,
            image=self._img_casilla_off,
            selectimage=self._img_casilla_on,
            compound="left",
            indicatoron=False,
            bd=0,
            highlightthickness=0,
            relief="flat",
            offrelief="flat",
            overrelief="flat",
            bg=COLOR_TARJETA,
            activebackground=COLOR_TARJETA,
            selectcolor=COLOR_TARJETA,
            fg=COLOR_TEXTO,
            activeforeground=COLOR_TEXTO,
            font=(FUENTE, 10),
            cursor="hand2",
            padx=0,
            pady=2,
        ).pack(anchor="w", pady=(8, 0))

        # Paso 3: generar
        paso3, cabecera_paso3 = self._tarjeta(contenedor, 3, "Generar los libres")
        ttk.Label(
            cabecera_paso3,
            text="Se guardan en Descargas",
            style="Ayuda.TLabel",
        ).pack(side="right")
        self.boton_generar = ttk.Button(
            paso3,
            text="Generar todos los libres",
            style="Primario.TButton",
            command=self._al_generar_todos,
            cursor="hand2",
        )
        self.boton_generar.pack(fill="x")

        marco_o = ttk.Frame(paso3, style="Tarjeta.TFrame")
        marco_o.pack(fill="x", pady=(8, 6))
        ttk.Separator(marco_o, style="Tarjeta.TSeparator").pack(side="left", fill="x", expand=True)
        ttk.Label(marco_o, text="o uno solo", style="Ayuda.TLabel").pack(side="left", padx=10)
        ttk.Separator(marco_o, style="Tarjeta.TSeparator").pack(side="left", fill="x", expand=True)

        marco_bl = ttk.Frame(paso3, style="Tarjeta.TFrame")
        marco_bl.pack(fill="x")
        ttk.Label(marco_bl, text="BL", style="Codigo.TLabel").pack(side="left")
        self.entrada_bl = ttk.Entry(marco_bl, textvariable=self.bl_var)
        self.entrada_bl.pack(side="left", fill="x", expand=True, padx=(10, 8))
        self.entrada_bl.bind("<Return>", lambda _e: self._al_generar_un_bl())
        self.boton_generar_bl = ttk.Button(
            marco_bl,
            text="Generar este BL",
            style="Accion.TButton",
            command=self._al_generar_un_bl,
            cursor="hand2",
        )
        self.boton_generar_bl.pack(side="left")

        # Actividad
        ttk.Label(contenedor, text="Actividad", style="Seccion.TLabel").pack(anchor="w", pady=(0, 4))
        self.progreso = ttk.Progressbar(contenedor, mode="determinate")
        self.progreso.pack(fill="x", pady=(0, 6))
        self.registro = ScrolledText(
            contenedor,
            height=4,
            wrap=WORD,
            font=("Consolas", 9),
            state=DISABLED,
            bg=COLOR_TARJETA,
            fg=COLOR_TEXTO,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLOR_BORDE,
            highlightcolor=COLOR_BORDE,
            padx=10,
            pady=6,
        )
        self.registro.pack(fill="both", expand=True)

        # Estado de wkhtmltopdf
        marco_estado = ttk.Frame(contenedor)
        marco_estado.pack(fill="x", pady=(6, 0))
        self.etiqueta_estado_wk = ttk.Label(marco_estado, text="", style="Pie.TLabel")
        self.etiqueta_estado_wk.pack(side="left")
        ttk.Button(
            marco_estado,
            text="Configurar wkhtmltopdf",
            style="Enlace.TButton",
            command=self._configurar_wkhtmltopdf,
            cursor="hand2",
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
                text="wkhtmltopdf no encontrado. Configuralo con el enlace de la derecha.",
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
            title="Selecciona el Excel de la planilla de facturación",
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
    def _editar_condiciones(self) -> None:
        ventana = Toplevel(self.root)
        ventana.title("Días de almacenaje")
        ventana.configure(bg=COLOR_TARJETA)
        ventana.resizable(False, False)
        ventana.transient(self.root)
        ventana.grab_set()

        marco = ttk.Frame(ventana, style="Tarjeta.TFrame", padding=20)
        marco.pack(fill="both", expand=True)
        ttk.Label(marco, text="Días libres por tipo", style="Paso.TLabel").pack(anchor="w")
        ttk.Label(
            marco,
            text="Se cuenta el día de llegada del barco. Otros tipos salen sin vencimiento.",
            style="Ayuda.TLabel",
            wraplength=380,
        ).pack(anchor="w", pady=(2, 14))

        filas = (
            ("VEH", "Vehículos y motorhome", self.condiciones.dias_veh),
            ("MAQ", "Máquinas", self.condiciones.dias_maq),
            ("CNT", "Contenedores de 20 y 40", self.condiciones.dias_cnt),
        )
        entradas: dict[str, StringVar] = {}
        for codigo, descripcion, valor in filas:
            fila = ttk.Frame(marco, style="Tarjeta.TFrame")
            fila.pack(fill="x", pady=4)
            ttk.Label(fila, text=codigo, width=6, style="Codigo.TLabel").pack(side="left")
            variable = StringVar(value=str(valor))
            entradas[codigo] = variable
            ttk.Entry(fila, textvariable=variable, width=5, justify="center").pack(
                side="left", padx=(0, 6)
            )
            ttk.Label(fila, text="días", style="Ayuda.TLabel").pack(side="left", padx=(0, 12))
            ttk.Label(fila, text=descripcion, style="Tarjeta.TLabel").pack(side="left")

        botones = ttk.Frame(marco, style="Tarjeta.TFrame")
        botones.pack(fill="x", pady=(18, 0))

        def guardar() -> None:
            try:
                nuevas = CondicionesAlmacenaje(
                    dias_veh=self._dias_ingresados(entradas["VEH"].get()),
                    dias_maq=self._dias_ingresados(entradas["MAQ"].get()),
                    dias_cnt=self._dias_ingresados(entradas["CNT"].get()),
                )
            except ValueError:
                messagebox.showwarning(
                    APP_TITLE,
                    "Cada condición tiene que ser un número entero de días, mayor a 0.",
                    parent=ventana,
                )
                return
            self.condiciones = nuevas
            self.config_usuario["condiciones_almacenaje"] = {
                "veh": nuevas.dias_veh,
                "maq": nuevas.dias_maq,
                "cnt": nuevas.dias_cnt,
            }
            _guardar_config(self.config_usuario)
            self._log(
                f"Días de almacenaje: VEH {nuevas.dias_veh}, "
                f"MAQ {nuevas.dias_maq}, CNT {nuevas.dias_cnt}."
            )
            ventana.destroy()

        ttk.Button(botones, text="Guardar", style="Accion.TButton", command=guardar).pack(side="right")
        ttk.Button(botones, text="Cancelar", command=ventana.destroy).pack(side="right", padx=(0, 8))

        ventana.update_idletasks()
        x = self.root.winfo_rootx() + 40
        y = self.root.winfo_rooty() + 80
        ventana.geometry(f"+{x}+{y}")
        ventana.wait_window()

    @staticmethod
    def _dias_ingresados(texto: str) -> int:
        numero = int(texto.strip())
        if numero < 1:
            raise ValueError
        return numero

    def _al_cambiar_incluir_fechas(self) -> None:
        if self.incluir_fechas_var.get():
            self.entrada_fecha.configure(state=NORMAL)
        else:
            self.entrada_fecha.configure(state=DISABLED)

    def _al_generar_todos(self) -> None:
        self._iniciar_generacion("")

    def _al_generar_un_bl(self) -> None:
        bl = self.bl_var.get().strip()
        if not bl:
            messagebox.showwarning(APP_TITLE, "Ingresa el numero de BL que queres generar.")
            return
        self._iniciar_generacion(bl)

    def _iniciar_generacion(self, bl: str) -> None:
        ruta_excel = self.ruta_excel_var.get().strip()
        incluir_fechas = bool(self.incluir_fechas_var.get())
        fecha_llegada = self.fecha_llegada_var.get().strip() if incluir_fechas else ""

        if not ruta_excel:
            messagebox.showwarning(
                APP_TITLE, "Primero selecciona el archivo Excel de la planilla de facturación."
            )
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
                "Usa 'Configurar wkhtmltopdf' (abajo a la derecha) para instalarlo o indicar su ubicacion.",
            )
            return

        self._set_generando(True)
        self._limpiar_log()
        self.progreso.configure(value=0, maximum=100)
        if bl:
            self._log(f"Leyendo el Excel para el BL {bl}...")
        else:
            self._log("Leyendo el Excel...")

        hilo = threading.Thread(
            target=self._ejecutar_generacion,
            args=(ruta_excel, fecha_llegada, incluir_fechas, bl),
            daemon=True,
        )
        hilo.start()

    def _set_generando(self, generando: bool) -> None:
        estado = DISABLED if generando else NORMAL
        self.boton_generar.configure(state=estado)
        self.boton_generar_bl.configure(state=estado)

    def _ejecutar_generacion(
        self, ruta_excel: str, fecha_llegada: str, incluir_fechas: bool, bl: str
    ) -> None:
        try:
            libres = leer_libres_desde_excel(ruta_excel)
            if bl:
                libres = filtrar_libres_por_bl(libres, bl)
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
                "Generando con fechas. " + self.condiciones.resumen()
                if incluir_fechas
                else "Generando sin fechas.",
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
                condiciones=self.condiciones,
            )

            self.root.after(0, self._al_finalizar_ok, len(generados), str(destino))
        except LibreGeneratorError as exc:
            self.root.after(0, self._al_finalizar_error, str(exc))
        except Exception as exc:  # error inesperado, igual mostramos algo claro
            self.root.after(0, self._al_finalizar_error, f"Ocurrio un error inesperado: {exc}")

    def _al_finalizar_ok(self, cantidad: int, carpeta: str) -> None:
        self._log(f"Listo. Se generaron {cantidad} libre(s) en: {carpeta}")
        self._set_generando(False)
        messagebox.showinfo(
            APP_TITLE, f"Se generaron {cantidad} libre(s) en tu carpeta Descargas."
        )

    def _al_finalizar_error(self, mensaje: str) -> None:
        self._log(f"ERROR: {mensaje}")
        self._set_generando(False)
        messagebox.showerror(APP_TITLE, mensaje)


def iniciar_app() -> None:
    root = Tk()
    GeneradorLibresApp(root)
    root.mainloop()
