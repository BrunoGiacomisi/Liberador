# Generador de Libres - TECNORDI SA

Aplicación de escritorio que genera automáticamente los PDF de "libre" a
partir de un Excel de manifiesto, para todas las líneas cuyo **DEPOSITO**
sea **1716**.

Cada libre generado se descarga como PDF en la carpeta **Descargas** del
usuario, nombrado con el número de **BL** de la línea correspondiente
(ejemplo: `S329568298.pdf`).

## Para el usuario final (no técnico)

1. **Instalar wkhtmltopdf** (una sola vez, es gratis):
   Descargalo desde <https://wkhtmltopdf.org/downloads.html> e instalalo con
   las opciones por defecto. La aplicación lo detecta automáticamente si
   quedó instalado en `C:\Program Files\wkhtmltopdf`. Si quedó en otro
   lugar, usa el botón **"Configurar wkhtmltopdf..."** dentro del programa
   para indicar dónde está `wkhtmltopdf.exe`.

2. **Abrir `GeneradorDeLibres.exe`** (no hace falta instalar nada más, ni
   tener Python).

3. Dentro del programa:
   - Paso 1: elegí el archivo Excel del manifiesto.
   - Paso 2: escribí la fecha de llegada del barco (DD/MM/AAAA).
   - Paso 3: apretá **"Generar libres"**.

4. Los PDF quedan listos en tu carpeta **Descargas**.

## Cómo interpreta el Excel

- La **fila 1** del Excel tiene el texto libre (ej: `PGD GRANDE FRANCIA 0426
  ETA 4/7/25 ESCALA: MANIFIESTO: 198401`). El programa toma ese texto hasta
  la primera serie de 4 números que encuentra (`0426`), y ese es el dato
  que va al lado de "ENTREGUESE:" en cada libre.
- La **fila 2** tiene los títulos de columna (BL, TP, QTY, DESCRIPCION,
  CHASIS, CNEE, DEPOSITO, FECHA LIBRE PARA RETIRO Y/O DEVOLUCIÓN, etc.).
- Se procesan solo las filas donde la columna **DEPOSITO** es **1716**.
- Cada línea filtrada genera un PDF independiente con:
  - Título: `TECNORDI SA`
  - `ENTREGUESE: <dato de la fila 1>`
  - Número de BL de la línea
  - CNEE de la línea
  - Una tabla con TP / QTY / DESCRIPCION / CHASIS / CNEE de esa línea
  - `Retiro de la mercadería en TMM`
  - `Llegada del barco: <fecha ingresada por el usuario>`
  - `Fecha de vencimiento: <columna "FECHA LIBRE PARA RETIRO Y/O
    DEVOLUCIÓN">`
  - El texto legal fijo de TECNORDI SA

Si dos líneas 1716 tienen el mismo número de BL, el segundo archivo se
guarda como `BL_2.pdf` para no pisar al primero.

La detección de columnas es flexible: no importa si hay pequeñas
variaciones de tildes/mayúsculas en los títulos de columna (por ejemplo
"QTY" o "QTI"), igual las reconoce. Si por algún motivo no encuentra una
columna por nombre, usa como respaldo la letra de columna estándar de la
planilla (C = BL, J = DEPOSITO, M = fecha de vencimiento).

## Para desarrollo / generar el .exe

Requisitos: Python 3.10+ instalado, y wkhtmltopdf instalado en el sistema
(<https://wkhtmltopdf.org/downloads.html>).

```bash
pip install -r requirements.txt
python main.py          # corre la app en modo desarrollo
```

Para generar el ejecutable distribuible (queda en `dist\GeneradorDeLibres.exe`):

```bash
build.bat
```

El `.exe` es autocontenido (incluye Python, pandas, Jinja2, etc.) y se
puede copiar/enviar a cualquier PC con Windows. Lo único que cada PC
necesita tener instalado aparte es **wkhtmltopdf** (no se puede
empaquetar dentro del .exe porque es un programa externo, no una
librería de Python).

## Estructura del proyecto

```
main.py                Punto de entrada (arranca la interfaz gráfica)
gui.py                  Interfaz gráfica (Tkinter)
libre_core.py           Lógica: leer Excel, filtrar 1716, generar PDF
templates/
  libre_template.html   Plantilla del PDF (Jinja2)
tests/
  generar_excel_prueba.py   Crea un Excel de ejemplo para probar
  probar_generacion.py      Prueba de punta a punta sin interfaz gráfica
build.bat                Script para generar el .exe con PyInstaller
requirements.txt
```
