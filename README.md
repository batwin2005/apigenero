# API Convertidor Excel → XML / CSV / PDF

> **No requiere Excel instalado.** Toda la conversión se ejecuta dentro del mismo API
> con librerías Python (`openpyxl`, `xlrd`, `pyxlsb`, `odfpy`, `reportlab`).
> Funciona en Windows, Linux o Docker sin Office. Ver `Dockerfile`.

Sube **cualquier Excel** (`.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.ods`, `.csv`) y lo transforma
al **formato que le envíes**: `xml`, `csv` o `pdf`. Luego lo **descargas en la ruta que tú le digas**.

> Los `.xlsm` con macros se leen como datos (openpyxl). El código VBA no se ejecuta.

## 1. Instalación

```powershell
py -m pip install -r requirements.txt
py -m pip install requests httpx   # solo para usar cliente.py
```

## 2. Arrancar el API

```powershell
py -m uvicorn app.main:app --reload --port 8000
```

- Docs interactivas: http://localhost:8000/docs
- Info: http://localhost:8000/
- Formatos: http://localhost:8000/api/v1/formatos

## 3. Cómo se envía el formato y se elige la ruta

El endpoint principal:

```
POST /api/v1/convertir?formato=xml|csv|pdf
```

| Parámetro      | Dónde | Qué hace |
|----------------|-------|----------|
| `archivo`      | body (multipart) | El Excel a convertir |
| `formato`      | query **(obligatorio)** | Formato al que se va a transformar: `xml`, `csv` o `pdf` |
| `hoja`         | query (opcional) | Solo esa hoja. Si se omite, todas |
| `nombre_salida`| query (opcional) | Nombre del archivo descargado. Ej: `reporte.xml` |
| `ruta_destino` | query (opcional) | Ruta ABSOLUTA del **servidor** donde guardar. Si se envía, responde JSON en vez del archivo |
| `delimiter`    | query (opcional) | Delimitador CSV de una hoja (por defecto `,`) |

Hay **2 formas de elegir dónde queda el archivo**:

### Opción A — Tú decides la ruta de descarga (recomendada)

No envíes `ruta_destino`. El API devuelve el archivo y **tú lo guardas donde quieras**:

```powershell
# Con cliente.py (lo recomendado: --salida es la ruta que tú le digas)
py cliente.py --archivo ejemplos/ejemplo.xlsm --formato xml --salida C:/salidas/reporte.xml
py cliente.py --archivo ejemplos/ejemplo.xlsm --formato pdf --salida ./salidas/reporte.pdf
py cliente.py --archivo ejemplos/ejemplo.xlsm --formato csv --salida ./salidas/personas.csv --hoja Personas

# Con curl (-o = ruta de descarga que tú eliges)
curl -X POST "http://localhost:8000/api/v1/convertir?formato=xml&nombre_salida=reporte.xml" -F "archivo=@ejemplos/ejemplo.xlsm" -o C:/salidas/reporte.xml
curl -X POST "http://localhost:8000/api/v1/convertir?formato=pdf" -F "archivo=@ejemplos/ejemplo.xlsm" -o ./salidas/reporte.pdf
```

### Opción B — El servidor guarda directo en la ruta que le envíes

Solo funciona si el API corre en tu misma máquina/red con acceso a esa ruta.
Envía `ruta_destino` (absoluta) y el API responde un JSON:

```powershell
curl -X POST "http://localhost:8000/api/v1/convertir?formato=pdf&ruta_destino=C:/salidas/desde_servidor.pdf" -F "archivo=@ejemplos/ejemplo.xlsm"

# con cliente.py
py cliente.py --archivo ejemplos/ejemplo.xlsm --formato pdf --ruta-destino C:/salidas/desde_servidor.pdf
```

Respuesta:

```json
{
  "ok": true,
  "mensaje": "Archivo convertido a pdf y guardado en el servidor.",
  "formato": "pdf",
  "ruta_guardada": "C:/salidas/desde_servidor.pdf",
  "bytes": 2752
}
```

## 4. Notas

- **XML**: por defecto es **Spreadsheet 2003** (`Workbook > Worksheet > Table > Row > Cell > Data`
  con tipos `String`/`Number`/`DateTime`/`Boolean`) — se abre con doble clic en Excel 2003 y posteriores.
  Sí: un `.xlsm` se convierte a este XML sin problema (las macros no se ejecutan, solo se leen los datos).
  Con `xml_estilo=simple` obtienes un XML genérico `<workbook><sheet><row>`.
  Ver ejemplo: `ejemplos/ejemplo_convertido.xml`.

- CSV con **una hoja** → devuelve `.csv`. Con **varias hojas** → devuelve `.zip` con un CSV por hoja.
- PDF incluye una tabla por hoja (máx. 500 filas por hoja).
- Límite de subida: 25 MB (ajustable en `app/main.py` → `MAX_FILE_MB`).
- Aclaración importante: por seguridad del protocolo HTTP, **el servidor no puede escribir
  en tu PC**. Por eso la ruta de descarga la elige el cliente (`--salida` / `-o`),
  y `ruta_destino` solo escribe en disco del servidor.
