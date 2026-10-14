"""API Convertidor Excel (XLS/XLSM/XLSX/...) -> XML / CSV / PDF.

Ejecución:
    pip install -r requirements.txt
    uvicorn app.main:app --reload --port 8000
Docs:
    http://localhost:8000/docs
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response

from .converter import (
    SUPPORTED_INPUT_EXTS,
    SUPPORTED_OUTPUTS,
    read_excel_sheets,
    sheet_to_csv_bytes,
    sheets_to_csv_zip,
    sheets_to_pdf,
    sheets_to_xml,
)

MAX_FILE_MB = 25
MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024

app = FastAPI(
    title="API Convertidor Excel → XML/CSV/PDF",
    description=(
        "Sube **cualquier Excel** (.xls, .xlsx, .xlsm, .xlsb, .ods, .csv) "
        "y conviértelo a **XML, CSV o PDF**. Soporta macros (.xlsm) — "
        "se leen los datos/valores, no se ejecuta código VBA."
    ),
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse, tags=["General"])
def home():
    return """
    <html><body style="font-family:sans-serif;max-width:720px;margin:40px auto">
    <h1>📄 API Convertidor Excel → XML / CSV / PDF</h1>
    <p>Convierte <b>.xls .xlsx .xlsm</b> (y más) a <b>XML, CSV, PDF</b>.</p>
    <ul>
      <li><a href="/docs">/docs</a> — Swagger UI (probar el API)</li>
      <li><a href="/api/v1/formatos">/api/v1/formatos</a> — formatos soportados</li>
      <li><a href="/health">/health</a> — estado</li>
    </ul>
    <h3>Ejemplo rápido (curl)</h3>
    <pre>curl -X POST "http://localhost:8000/api/v1/convertir?formato=xml" -F "archivo=@ejemplo.xlsm" -o salida.xml</pre>
    </body></html>
    """


@app.get("/health", tags=["General"])
def health():
    return {"status": "ok"}


@app.get("/api/v1/formatos", tags=["General"])
def formatos():
    return {
        "entrada": sorted(SUPPORTED_INPUT_EXTS),
        "salida": sorted(SUPPORTED_OUTPUTS),
        "nota_xlsm": "Los .xlsm con macros se leen como datos (openpyxl). El VBA no se ejecuta.",
        "nota_xml": "El XML es Spreadsheet 2003 por defecto (abre con Excel 2003 y posteriores). Usa xml_estilo=simple para XML genérico.",
        "tamaño_max_mb": MAX_FILE_MB,
    }


@app.post("/api/v1/convertir", tags=["Conversión"])
async def convertir(
    archivo: UploadFile = File(..., description="Archivo Excel (.xls, .xlsx, .xlsm, ...)"),
    formato: Literal["xml", "csv", "pdf"] = Query(..., description="Formato al que se va a transformar: xml, csv o pdf"),
    hoja: str | None = Query(None, description="Nombre de hoja a convertir. Si se omite, se convierten todas."),
    delimiter: str = Query(",", description="Delimitador CSV (solo aplica a CSV de una hoja)"),
    xml_estilo: Literal["2003", "simple"] = Query(
        "2003",
        description="Estilo del XML: '2003' (Spreadsheet 2003, abre en Excel) o 'simple' (genérico <workbook>).",
    ),
    nombre_salida: str | None = Query(
        None,
        description="Nombre del archivo de descarga. Ej: reporte.xml. Si se omite se usa el nombre del Excel.",
    ),
    ruta_destino: str | None = Query(
        None,
        description="Ruta ABSOLUTA del servidor donde guardar el resultado. "
        "Solo funciona si el API corre en tu misma máquina/red con acceso a esa ruta. "
        "Ej: C:/salidas/reporte.pdf o /tmp/reporte.xml. "
        "Si se envía, el API guarda ahí y responde JSON en vez de descargar el binario.",
    ),
):
    """Convierte un Excel al formato pedido.

    - **formato**: el formato al que se va a transformar (xml, csv, pdf).
    - **Descarga normal**: no envíes `ruta_destino` y el API te devuelve el archivo
      para que TÚ lo guardes en la ruta que quieras (ver `cliente.py --salida RUTA` o `curl -o RUTA`).
    - **Guardado en servidor**: envía `ruta_destino` y el API guarda el archivo
      directamente en esa ruta del servidor y responde un JSON con la ruta.
    """
    if not archivo.filename:
        raise HTTPException(400, "Debes enviar un archivo.")

    ext = Path(archivo.filename).suffix.lower()
    if ext not in SUPPORTED_INPUT_EXTS:
        raise HTTPException(
            400,
            f"Extensión '{ext}' no soportada. Soportadas: {sorted(SUPPORTED_INPUT_EXTS)}",
        )

    content = await archivo.read()
    if not content:
        raise HTTPException(400, "El archivo está vacío.")
    if len(content) > MAX_FILE_BYTES:
        raise HTTPException(400, f"Archivo supera el máximo de {MAX_FILE_MB} MB.")

    try:
        sheets = read_excel_sheets(content, archivo.filename)
    except ValueError as e:
        raise HTTPException(400, str(e))

    # Filtrar a una hoja si se pide
    if hoja:
        if hoja not in sheets:
            raise HTTPException(
                400, f"Hoja '{hoja}' no encontrada. Hojas: {list(sheets.keys())}"
            )
        sheets = {hoja: sheets[hoja]}

    base = Path(archivo.filename).stem or "archivo"

    payload: bytes
    media_type: str
    default_filename: str

    # ---------- XML (Spreadsheet 2003 por defecto) ----------
    if formato == "xml":
        payload = sheets_to_xml(sheets, estilo=xml_estilo)
        media_type = "application/xml"
        default_filename = f"{base}.xml"

    # ---------- CSV ----------
    elif formato == "csv":
        if len(sheets) == 1:
            df = next(iter(sheets.values()))
            buf = io.StringIO()
            # delimiter personalizado
            df.to_csv(buf, index=False, sep=delimiter if len(delimiter) == 1 else ",")
            payload = "\ufeff".encode("utf-8") + buf.getvalue().encode("utf-8")
            media_type = "text/csv; charset=utf-8"
            default_filename = f"{base}.csv"
        else:
            # varias hojas -> ZIP con un CSV por hoja
            payload = sheets_to_csv_zip(sheets)
            media_type = "application/zip"
            default_filename = f"{base}_csvs.zip"

    # ---------- PDF ----------
    elif formato == "pdf":
        payload = sheets_to_pdf(sheets, title=f"Reporte: {archivo.filename}")
        media_type = "application/pdf"
        default_filename = f"{base}.pdf"

    else:
        raise HTTPException(400, "Formato inválido. Usa xml, csv o pdf.")

    # Nombre de descarga: el que pida el usuario o el de por defecto
    download_name = (nombre_salida.strip() if nombre_salida else "") or default_filename
    # Evitar path traversal en el header: solo el nombre, sin carpetas
    download_name = Path(download_name).name or default_filename

    # ---- Caso 2: guardar directo en ruta del SERVIDOR ----
    if ruta_destino:
        try:
            dest = Path(ruta_destino).expanduser()
            if not dest.is_absolute():
                raise HTTPException(
                    400,
                    "ruta_destino debe ser ABSOLUTA. Ej: C:/salidas/reporte.pdf o /tmp/reporte.xml",
                )
            # Si pasan una carpeta, completar con el nombre de descarga
            if dest.suffix == "" or ruta_destino.endswith(("/", "\\")):
                dest = dest / download_name
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(payload)
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(400, f"No se pudo guardar en ruta_destino: {e}")
        return {
            "ok": True,
            "mensaje": f"Archivo convertido a {formato} y guardado en el servidor.",
            "formato": formato,
            "archivo_origen": archivo.filename,
            "hojas": list(sheets.keys()),
            "ruta_guardada": str(dest),
            "bytes": len(payload),
            "descarga_sugerida": download_name,
        }

    # ---- Caso 1 (normal): devolver binario para que el CLIENTE lo guarde donde quiera ----
    return Response(
        content=payload,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )


# Alias en inglés por comodidad
@app.post("/api/v1/convert", tags=["Conversión"])
async def convertir_en(
    archivo: UploadFile = File(...),
    formato: Literal["xml", "csv", "pdf"] = Query(...),
    hoja: str | None = Query(None),
    xml_estilo: Literal["2003", "simple"] = Query("2003"),
    nombre_salida: str | None = Query(None),
    ruta_destino: str | None = Query(None),
):
    return await convertir(
        archivo=archivo, formato=formato, hoja=hoja, xml_estilo=xml_estilo,
        nombre_salida=nombre_salida, ruta_destino=ruta_destino,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
