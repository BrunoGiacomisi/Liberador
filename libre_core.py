"""
Core del Generador de Libres.

Se encarga de:
  1. Leer el Excel de manifiesto.
  2. Quedarse solo con las lineas cuyo DEPOSITO sea 1716 (retiro TMM)
     o 1714 (libre de contenedor).
  3. Armar los datos de cada "libre" y generar su PDF a partir de una
     plantilla HTML (Jinja2 + wkhtmltopdf/pdfkit).

Pensado para ser usado tanto desde la interfaz grafica (gui.py) como
desde un script / consola si hiciera falta.
"""

from __future__ import annotations

import base64
import os
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
from jinja2 import Environment, FileSystemLoader

try:
    import pdfkit
except ImportError:  # pdfkit se valida igual mas adelante con un mensaje claro
    pdfkit = None


def _directorio_base() -> Path:
    """Carpeta base de la app: cuando corre como .exe (PyInstaller), los
    archivos como la plantilla HTML se extraen en sys._MEIPASS."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# Configuracion / constantes
# ---------------------------------------------------------------------------

DEPOSITO_RETIRO_TMM = "1716"
DEPOSITO_CONTENEDOR = "1714"

# Fila de Excel (1-indexada, tal cual se ve en la planilla) donde estan los
# titulos de columna. La fila 1 tiene el texto libre (PGD GRANDE FRANCIA...).
FILA_ENCABEZADOS_EXCEL = 2

# Ubicacion de las plantillas HTML.
TEMPLATE_DIR = _directorio_base() / "templates"
TEMPLATE_NAME = "libre_template.html"
RUTA_LOGO = _directorio_base() / "assets" / "kma_logo.png"

# Rutas donde se suele instalar wkhtmltopdf en Windows.
RUTAS_WKHTMLTOPDF_HABITUALES = [
    r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
    r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
    r"C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe",
]


class LibreGeneratorError(Exception):
    """Errores esperables y con mensaje amigable para mostrar al usuario."""


@dataclass
class ItemLibre:
    tp: str
    qty: str
    descripcion: str
    chasis: str
    cnee: str


@dataclass
class Libre:
    """Toda la informacion necesaria para renderizar un PDF de libre."""

    bl: str
    extra1: str
    extra2: str
    extra3: str
    fecha_vencimiento: str
    deposito: str = DEPOSITO_RETIRO_TMM
    items: list[ItemLibre] = field(default_factory=list)

    @property
    def es_contenedor(self) -> bool:
        return self.deposito == DEPOSITO_CONTENEDOR

    @property
    def nombre_base_archivo(self) -> str:
        if self.es_contenedor:
            return f"{self.bl} CNT"
        return self.bl


# ---------------------------------------------------------------------------
# Utilidades de texto / normalizacion
# ---------------------------------------------------------------------------

def _normalizar(texto: str) -> str:
    """Mayusculas, sin acentos y con espacios simples. Sirve para comparar
    encabezados de columnas sin importar tildes, mayusculas o espacios raros.
    """
    if texto is None:
        return ""
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = texto.upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def _valor_celda(valor) -> str:
    """Convierte un valor de celda de pandas a texto prolijo para mostrar."""
    if valor is None:
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    if pd.isna(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def _formatear_fecha(valor) -> str:
    """Devuelve la fecha en formato DD/MM/AAAA sin importar como venga
    guardada en el Excel (texto, datetime, etc.)."""
    texto = _valor_celda(valor)
    if not texto:
        return ""
    try:
        fecha = pd.to_datetime(valor, dayfirst=True, errors="raise")
        return fecha.strftime("%d/%m/%Y")
    except Exception:
        return texto


def extraer_encabezado_hasta_cuatro_digitos(texto: str) -> str:
    """A partir del texto libre de la fila 1 del Excel, devuelve todo hasta
    (e incluyendo) la primera serie de exactamente 4 numeros que aparezca.

    Ejemplo: "PGD GRANDE FRANCIA 0426 ETA 4/7/25 ESCALA: MANIFIESTO: 198401"
             -> "PGD GRANDE FRANCIA 0426"
    """
    if not texto:
        return ""
    match = re.search(r"\b\d{4}\b", str(texto))
    if not match:
        return str(texto).strip()
    return str(texto)[: match.end()].strip()


def sanitizar_nombre_archivo(nombre: str) -> str:
    """Elimina caracteres invalidos para nombres de archivo en Windows."""
    nombre = (nombre or "SIN_BL").strip()
    nombre = re.sub(r'[\\/*?:"<>|]', "_", nombre)
    nombre = re.sub(r"\s+", " ", nombre).strip()
    return nombre or "SIN_BL"


# ---------------------------------------------------------------------------
# Deteccion de columnas del Excel (tolerante a variaciones de encabezado)
# ---------------------------------------------------------------------------

# Cada entrada: nombre logico -> lista de palabras clave que deben estar
# TODAS presentes (normalizadas) en el encabezado de la columna.
_COLUMNAS_BUSCADAS: dict[str, list[str]] = {
    "bl": ["BL"],
    "tp": ["TP"],
    "qty": ["QT"],  # cubre QTY y QTI (variantes vistas en las planillas)
    "descripcion": ["DESCRIPCION"],
    "chasis": ["CHASIS"],
    "cnee": ["CNEE"],
    "deposito": ["DEPOSITO"],
    "fecha_vencimiento": ["RETIRO", "DEVOLU"],
}

# Como respaldo, si no se encuentra por nombre, se usa la letra de columna
# de Excel indicada por el usuario (0 = A, 1 = B, ...).
_COLUMNAS_RESPALDO_LETRA: dict[str, int] = {
    "bl": 2,          # columna C
    "deposito": 9,    # columna J
    "fecha_vencimiento": 12,  # columna M
}


def _detectar_columnas(columnas_excel: list) -> dict[str, str]:
    """Devuelve un dict nombre_logico -> nombre real de columna del DataFrame."""
    normalizadas = {col: _normalizar(col) for col in columnas_excel}

    resultado: dict[str, str] = {}
    for logico, palabras_clave in _COLUMNAS_BUSCADAS.items():
        encontrada = None
        for col_original, col_norm in normalizadas.items():
            if all(palabra in col_norm for palabra in palabras_clave):
                encontrada = col_original
                break
        if encontrada is None and logico in _COLUMNAS_RESPALDO_LETRA:
            idx = _COLUMNAS_RESPALDO_LETRA[logico]
            if idx < len(columnas_excel):
                encontrada = columnas_excel[idx]
        if encontrada is not None:
            resultado[logico] = encontrada

    faltantes = [c for c in ("deposito", "bl", "cnee") if c not in resultado]
    if faltantes:
        raise LibreGeneratorError(
            "No se pudieron identificar en el Excel las columnas: "
            + ", ".join(faltantes)
            + ". Revisa que la fila 2 tenga los titulos correctos "
            "(por ejemplo: BL, DEPOSITO, CNEE)."
        )

    return resultado


# ---------------------------------------------------------------------------
# Lectura del Excel y armado de los "libres"
# ---------------------------------------------------------------------------

def _item_desde_fila(fila, columnas: dict[str, str]) -> ItemLibre:
    cnee = _valor_celda(fila.get(columnas.get("cnee", ""), ""))
    return ItemLibre(
        tp=_valor_celda(fila.get(columnas.get("tp", ""), "")),
        qty=_valor_celda(fila.get(columnas.get("qty", ""), "")),
        descripcion=_valor_celda(fila.get(columnas.get("descripcion", ""), "")),
        chasis=_valor_celda(fila.get(columnas.get("chasis", ""), "")),
        cnee=cnee,
    )


def _libre_desde_filas(
    filas: list,
    extra1: str,
    columnas: dict[str, str],
    deposito: str,
) -> Libre:
    primera = filas[0]
    bl = _valor_celda(primera.get(columnas.get("bl", ""), ""))
    cnee = _valor_celda(primera.get(columnas.get("cnee", ""), ""))
    fecha_venc = _formatear_fecha(
        primera.get(columnas.get("fecha_vencimiento", ""), "")
    )
    return Libre(
        bl=bl,
        extra1=extra1,
        extra2=bl,
        extra3=cnee,
        fecha_vencimiento=fecha_venc,
        deposito=deposito,
        items=[_item_desde_fila(fila, columnas) for fila in filas],
    )


def _agrupar_filas_por_bl(filas: list, columnas: dict[str, str]) -> list[list]:
    """Agrupa filas por BL respetando el orden de primera aparicion."""
    grupos: dict[str, list] = {}
    orden: list[str] = []
    for fila in filas:
        bl = _valor_celda(fila.get(columnas.get("bl", ""), ""))
        if bl not in grupos:
            grupos[bl] = []
            orden.append(bl)
        grupos[bl].append(fila)
    return [grupos[bl] for bl in orden]


def _libres_agrupados_por_bl(
    df,
    extra1: str,
    columnas: dict[str, str],
    deposito: str,
) -> list[Libre]:
    """Arma un libre por cada BL, juntando todas las filas de ese BL."""
    filas = [fila for _, fila in df.iterrows()]
    return [
        _libre_desde_filas(grupo, extra1, columnas, deposito)
        for grupo in _agrupar_filas_por_bl(filas, columnas)
    ]


def leer_libres_desde_excel(ruta_excel: str | Path) -> list[Libre]:
    """Lee el Excel y arma los libres segun el deposito.

    Tanto 1716 (retiro TMM) como 1714 (contenedor) generan un PDF por BL:
    si hay varias filas con el mismo numero, van juntas en la misma tabla.
    """

    ruta_excel = Path(ruta_excel)
    if not ruta_excel.exists():
        raise LibreGeneratorError(f"No se encontro el archivo: {ruta_excel}")

    try:
        encabezado_general_df = pd.read_excel(
            ruta_excel, header=None, nrows=1, engine="openpyxl"
        )
        texto_fila1 = _valor_celda(encabezado_general_df.iloc[0, 0])
    except Exception as exc:
        raise LibreGeneratorError(f"No se pudo leer la fila 1 del Excel: {exc}") from exc

    extra1 = extraer_encabezado_hasta_cuatro_digitos(texto_fila1)

    try:
        df = pd.read_excel(
            ruta_excel, header=FILA_ENCABEZADOS_EXCEL - 1, engine="openpyxl"
        )
    except Exception as exc:
        raise LibreGeneratorError(f"No se pudo leer el Excel: {exc}") from exc

    df = df.dropna(how="all")
    if df.empty:
        raise LibreGeneratorError("El Excel no tiene filas de datos.")

    columnas = _detectar_columnas(list(df.columns))

    col_deposito = columnas["deposito"]
    depositos = df[col_deposito].apply(lambda v: _normalizar(_valor_celda(v)))
    df_1716 = df[depositos == DEPOSITO_RETIRO_TMM]
    df_1714 = df[depositos == DEPOSITO_CONTENEDOR]

    if df_1716.empty and df_1714.empty:
        raise LibreGeneratorError(
            "No se encontraron lineas con DEPOSITO = 1716 ni 1714 en el Excel."
        )

    libres: list[Libre] = []
    libres.extend(_libres_agrupados_por_bl(df_1716, extra1, columnas, DEPOSITO_RETIRO_TMM))
    libres.extend(_libres_agrupados_por_bl(df_1714, extra1, columnas, DEPOSITO_CONTENEDOR))
    return libres


# ---------------------------------------------------------------------------
# wkhtmltopdf: deteccion de la instalacion
# ---------------------------------------------------------------------------

def encontrar_wkhtmltopdf(ruta_configurada: Optional[str] = None) -> Optional[str]:
    """Busca el ejecutable de wkhtmltopdf en la ruta configurada, en el
    PATH del sistema y en las ubicaciones habituales de instalacion."""

    candidatos = []
    if ruta_configurada:
        candidatos.append(ruta_configurada)

    en_path = shutil.which("wkhtmltopdf")
    if en_path:
        candidatos.append(en_path)

    candidatos.extend(RUTAS_WKHTMLTOPDF_HABITUALES)

    for candidato in candidatos:
        if candidato and Path(candidato).is_file():
            return candidato
    return None


# ---------------------------------------------------------------------------
# Generacion de PDFs
# ---------------------------------------------------------------------------

def _construir_config_pdfkit(ruta_wkhtmltopdf: str):
    if pdfkit is None:
        raise LibreGeneratorError(
            "Falta instalar la libreria 'pdfkit' (pip install pdfkit)."
        )
    return pdfkit.configuration(wkhtmltopdf=ruta_wkhtmltopdf)


_logo_data_uri_cache: Optional[str] = None


def _logo_data_uri() -> str:
    """Devuelve el logo de KMA como data URI (base64) para incrustarlo
    directamente en el HTML, sin depender de rutas de archivo externas
    (esto es importante para que funcione igual como script y como .exe)."""
    global _logo_data_uri_cache
    if _logo_data_uri_cache is None:
        try:
            datos = RUTA_LOGO.read_bytes()
            b64 = base64.b64encode(datos).decode("ascii")
            _logo_data_uri_cache = f"data:image/png;base64,{b64}"
        except Exception:
            _logo_data_uri_cache = ""
    return _logo_data_uri_cache


def renderizar_html_libre(
    libre: Libre,
    fecha_llegada: str,
    incluir_fechas: bool = True,
) -> str:
    env = Environment(loader=FileSystemLoader(str(TEMPLATE_DIR)))
    template = env.get_template(TEMPLATE_NAME)
    return template.render(
        titulo="TECNORDI SA",
        extra1=libre.extra1,
        extra2=libre.extra2,
        extra3=libre.extra3,
        fecha_llegada=fecha_llegada,
        fecha_vencimiento=libre.fecha_vencimiento,
        items=libre.items,
        es_contenedor=libre.es_contenedor,
        incluir_fechas=incluir_fechas,
        logo_data_uri=_logo_data_uri(),
    )


def _nombre_archivo_disponible(carpeta: Path, nombre_base: str) -> Path:
    """Evita pisar archivos si dos libres terminaran con el mismo nombre
    (por ejemplo, dos lineas con el mismo numero de BL)."""
    nombre_base = sanitizar_nombre_archivo(nombre_base)
    destino = carpeta / f"{nombre_base}.pdf"
    contador = 2
    while destino.exists():
        destino = carpeta / f"{nombre_base}_{contador}.pdf"
        contador += 1
    return destino


def generar_pdfs(
    libres: list[Libre],
    fecha_llegada: str,
    carpeta_destino: str | Path,
    ruta_wkhtmltopdf: str,
    incluir_fechas: bool = True,
    on_progreso: Optional[Callable[[int, int, str], None]] = None,
) -> list[Path]:
    """Genera un PDF por cada libre en la carpeta indicada.

    on_progreso(indice_actual, total, nombre_archivo) se llama despues de
    generar cada PDF, util para actualizar una barra de progreso en la GUI.
    """
    carpeta_destino = Path(carpeta_destino)
    carpeta_destino.mkdir(parents=True, exist_ok=True)

    config = _construir_config_pdfkit(ruta_wkhtmltopdf)
    opciones_pdf = {
        "encoding": "UTF-8",
        "quiet": "",
        "page-size": "A4",
        "margin-top": "20mm",
        "margin-bottom": "20mm",
        "margin-left": "22mm",
        "margin-right": "22mm",
    }

    generados: list[Path] = []
    total = len(libres)
    for idx, libre in enumerate(libres, start=1):
        html = renderizar_html_libre(libre, fecha_llegada, incluir_fechas=incluir_fechas)
        destino = _nombre_archivo_disponible(carpeta_destino, libre.nombre_base_archivo)
        try:
            pdfkit.from_string(html, str(destino), configuration=config, options=opciones_pdf)
        except Exception as exc:
            raise LibreGeneratorError(
                f"No se pudo generar el PDF para el BL '{libre.bl}': {exc}"
            ) from exc
        generados.append(destino)
        if on_progreso:
            on_progreso(idx, total, destino.name)

    return generados


def carpeta_descargas() -> Path:
    """Devuelve la carpeta Descargas del usuario actual de Windows."""
    descargas = Path.home() / "Downloads"
    if descargas.exists():
        return descargas
    return Path.home()
