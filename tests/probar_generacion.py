"""Prueba de punta a punta sin interfaz grafica: lee el Excel de prueba,
filtra deposito 1716 y genera los PDF en una carpeta de salida local
(en vez de la carpeta Descargas real, para no ensuciarla durante pruebas)."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from libre_core import (  # noqa: E402
    encontrar_wkhtmltopdf,
    generar_pdfs,
    leer_libres_desde_excel,
)

RUTA_EXCEL = Path(__file__).resolve().parent / "manifiesto_prueba.xlsx"
CARPETA_SALIDA = Path(__file__).resolve().parent / "salida_prueba"


def main() -> None:
    ruta_wk = encontrar_wkhtmltopdf()
    print(f"wkhtmltopdf detectado en: {ruta_wk}")
    if not ruta_wk:
        print("No se encontro wkhtmltopdf, abortando prueba.")
        return

    libres = leer_libres_desde_excel(RUTA_EXCEL)
    print(f"Lineas con deposito 1716 encontradas: {len(libres)}")
    for libre in libres:
        print(
            f"  BL={libre.bl} | extra1={libre.extra1!r} | extra3={libre.extra3!r} | "
            f"vencimiento={libre.fecha_vencimiento} | item={libre.items[0]}"
        )

    generados = generar_pdfs(
        libres,
        fecha_llegada="13/06/2026",
        carpeta_destino=CARPETA_SALIDA,
        ruta_wkhtmltopdf=ruta_wk,
        on_progreso=lambda i, t, n: print(f"  [{i}/{t}] {n}"),
    )
    print(f"PDFs generados: {len(generados)} en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
