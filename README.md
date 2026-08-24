# Generador de Libres - TECNORDI SA

Aplicación de escritorio que genera automáticamente los PDF de "libre" a
partir de un Excel de manifiesto, para las líneas cuyo **DEPOSITO** sea
**1716** (retiro TMM) o **1714** (libre de contenedor).

Los PDF se descargan en la carpeta **Descargas** del usuario:

- Depósito **1716**: el archivo se nombra con el número de BL
  (ejemplo: `S329568298.pdf`).
- Depósito **1714**: el archivo se nombra con el BL más `CNT`
  (ejemplo: `S329701796 CNT.pdf`).

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
   - Paso 2: escribí la fecha de llegada del barco (DD/MM/AAAA). Si
     desmarcás **Incluir fechas en el PDF**, el documento sale sin
     llegada del barco ni fecha de vencimiento.
   - Paso 3: apretá **"Generar libres"**.

4. Los PDF quedan listos en tu carpeta **Descargas**.

## Cómo interpreta el Excel

- La **fila 1** del Excel tiene el texto libre (ej: `PGD GRANDE FRANCIA 0426
  ETA 4/7/25 ESCALA: MANIFIESTO: 198401`). El programa toma ese texto hasta
  la primera serie de 4 números que encuentra (`0426`), y ese es el dato
  que va al lado de "ENTREGUESE:" en cada libre.
- La **fila 2** tiene los títulos de columna (BL, TP, QTY, DESCRIPCION,
  CHASIS, CNEE, DEPOSITO, FECHA LIBRE PARA RETIRO Y/O DEVOLUCIÓN, etc.).
- Se procesan las filas donde la columna **DEPOSITO** es **1716** o **1714**.

### Depósito 1716 (retiro TMM)

Todas las líneas 1716 que compartan el mismo BL van juntas en **un solo
PDF**, con una fila de tabla por cada línea. El documento incluye:

- Título: `TECNORDI SA`
- `ENTREGUESE: <dato de la fila 1>`
- Número de BL
- CNEE
- Una tabla con TP / QTY / DESCRIPCION / CHASIS / CNEE (una fila por
  cada línea del mismo BL)
- `Retiro de la mercadería en TMM`
- `Llegada del barco: <fecha ingresada por el usuario>`
- `Fecha de vencimiento: <columna "FECHA LIBRE PARA RETIRO Y/O
  DEVOLUCIÓN">`
- El texto legal fijo de TECNORDI SA

### Depósito 1714 (libre de contenedor)

Igual que 1716: todas las líneas 1714 del mismo BL van juntas en **un
solo PDF**. El documento es igual, con estas diferencias:

- Pie en rojo subrayado: `Devolución Murchison`
- Primero `Fecha de vencimiento` y después `Llegada del barco`
- Nombre de archivo: `BL CNT.pdf` (ejemplo: `S329701796 CNT.pdf`)

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
libre_core.py           Lógica: leer Excel, filtrar 1716/1714, generar PDF
templates/
  libre_template.html   Plantilla del PDF (Jinja2)
tests/
  generar_excel_prueba.py   Crea un Excel de ejemplo para probar
  probar_generacion.py      Prueba de punta a punta sin interfaz gráfica
build.bat                Script para generar el .exe con PyInstaller
requirements.txt
```
