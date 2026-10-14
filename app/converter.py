"""Lógica de lectura Excel y conversión a XML / CSV / PDF."""
from __future__ import annotations

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import Dict

import pandas as pd

# Extensiones de entrada soportadas -> engine de pandas
SUPPORTED_INPUT_EXTS = {".xls", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".ods", ".csv"}
SUPPORTED_OUTPUTS = {"xml", "csv", "pdf"}

MAX_PDF_ROWS_PER_SHEET = 500  # evita PDFs gigantes


def _engine_for_ext(ext: str) -> str | None:
    ext = ext.lower()
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return "openpyxl"
    if ext == ".xls":
        return "xlrd"
    if ext == ".xlsb":
        return "pyxlsb"
    if ext == ".ods":
        return "odf"
    return None


def read_excel_sheets(file_bytes: bytes, filename: str) -> Dict[str, pd.DataFrame]:
    """Lee cualquier Excel soportado y devuelve {nombre_hoja: DataFrame}."""
    ext = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in SUPPORTED_INPUT_EXTS:
        raise ValueError(f"Extensión '{ext}' no soportada. Usa: {sorted(SUPPORTED_INPUT_EXTS)}")

    bio = io.BytesIO(file_bytes)

    # CSV de entrada se trata como una sola hoja
    if ext == ".csv":
        df = pd.read_csv(bio)
        return {"Sheet1": df}

    engine = _engine_for_ext(ext)
    try:
        # sheet_name=None -> dict de DataFrames (todas las hojas)
        sheets = pd.read_excel(bio, sheet_name=None, engine=engine)
    except Exception as e:
        raise ValueError(f"No se pudo leer el Excel ({filename}): {e}") from e

    if not sheets:
        raise ValueError("El archivo Excel no contiene hojas legibles.")
    return sheets


def clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value)


def sanitize_tag(name: str) -> str:
    """Convierte nombre de columna en tag XML válido."""
    name = str(name).strip() or "columna"
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"[^A-Za-z0-9_.-]", "", name)
    if not re.match(r"^[A-Za-z_]", name):
        name = "col_" + name
    return name[:60] or "columna"


# ---------------- XML ----------------
# Por defecto: XML Spreadsheet 2003 (el .xml que abre Excel 2003).
# Opción "simple": XML genérico <workbook><sheet><row>...

NS_MAIN = "urn:schemas-microsoft-com:office:spreadsheet"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_X = "urn:schemas-microsoft-com:office:excel"
NS_SS = "urn:schemas-microsoft-com:office:spreadsheet"
NS_HTML = "http://www.w3.org/TR/REC-html40"

for _prefix, _uri in (("o", NS_O), ("x", NS_X), ("ss", NS_SS), ("html", NS_HTML)):
    ET.register_namespace(_prefix, _uri)
ET.register_namespace("", NS_MAIN)


# Atributos que Excel 2003 exige con prefijo ss: (mismo namespace que el
# default, pero ElementTree los serializa sin prefijo; se corrigen abajo).
_SS_ATTRS = ("Name", "ExpandedColumnCount", "ExpandedRowCount", "Type")
_SS_ATTR_RE = re.compile(r'<[^>]+>')
_SS_ATTR_IN_TAG_RE = re.compile(r'\s(Name|ExpandedColumnCount|ExpandedRowCount|Type)="')

def _prefix_ss_attributes(xml_text: str) -> str:
    """Restaura el prefijo ss: en los atributos dentro de los tags (no toca el texto)."""
    def _fix_tag(m: re.Match) -> str:
        return _SS_ATTR_IN_TAG_RE.sub(r' ss:\1="', m.group(0))
    return _SS_ATTR_RE.sub(_fix_tag, xml_text)


def _clean_sheet_name(name: str) -> str:
    # Excel 2003: máx 31 caracteres, sin : \ / ? * [ ]
    name = re.sub(r'[:\\/?*\[\]]', "_", str(name)).strip() or "Hoja"
    return name[:31]


def _cell_2003(parent_row: ET.Element, value) -> None:
    """Añade <Cell><Data ss:Type=...> al Row según el tipo del valor."""
    import datetime as _dt

    if value is None or (not isinstance(value, (_dt.datetime, _dt.date)) and pd.isna(value)):
        ET.SubElement(parent_row, f"{{{NS_SS}}}Cell")
        return
    cell = ET.SubElement(parent_row, f"{{{NS_SS}}}Cell")
    if isinstance(value, bool):
        data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "Boolean"})
        data.text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        # NaN ya filtrado arriba; float('inf') Excel no lo soporta -> string
        import math
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "String"})
            data.text = ""
        else:
            data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "Number"})
            data.text = str(value)
    elif isinstance(value, _dt.datetime):
        data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "DateTime"})
        data.text = value.strftime("%Y-%m-%dT%H:%M:%S")
    elif isinstance(value, _dt.date):
        data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "DateTime"})
        data.text = value.strftime("%Y-%m-%dT00:00:00")
    elif isinstance(value, pd.Timestamp):
        data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "DateTime"})
        data.text = value.to_pydatetime().strftime("%Y-%m-%dT%H:%M:%S")
    else:
        data = ET.SubElement(cell, f"{{{NS_SS}}}Data", {f"{{{NS_SS}}}Type": "String"})
        data.text = str(value)


def sheets_to_xml_2003(sheets: Dict[str, pd.DataFrame]) -> bytes:
    """Genera XML Spreadsheet 2003 (compatible con Excel 2003, se abre con doble clic)."""
    root = ET.Element(f"{{{NS_MAIN}}}Workbook", {
        f"{{{NS_O}}}dummy": "x",  # se elimina abajo; fuerza declaración xmlns:o
    })
    # Eliminar el atributo dummy y declarar namespaces extra que Excel espera
    del root.attrib[f"{{{NS_O}}}dummy"]
    root.set("xmlns:o", NS_O)
    root.set("xmlns:x", NS_X)
    root.set("xmlns:ss", NS_SS)
    root.set("xmlns:html", NS_HTML)

    for sheet_name, df in sheets.items():
        ws = ET.SubElement(root, f"{{{NS_SS}}}Worksheet",
                           {f"{{{NS_SS}}}Name": _clean_sheet_name(sheet_name)})
        n_cols = max(len(df.columns), 1)
        n_rows = len(df) + 1  # +1 encabezado
        table = ET.SubElement(ws, f"{{{NS_SS}}}Table", {
            f"{{{NS_SS}}}ExpandedColumnCount": str(n_cols),
            f"{{{NS_SS}}}ExpandedRowCount": str(max(n_rows, 1)),
        })
        # Fila de encabezados
        header_row = ET.SubElement(table, f"{{{NS_SS}}}Row")
        for h in df.columns.astype(str):
            _cell_2003(header_row, str(h))
        # Filas de datos
        for _, row in df.iterrows():
            xml_row = ET.SubElement(table, f"{{{NS_SS}}}Row")
            for val in row:
                _cell_2003(xml_row, val)

    body = _prefix_ss_attributes(ET.tostring(root, encoding="utf-8").decode("utf-8"))
    header = ('<?xml version="1.0" encoding="UTF-8"?>\n'
              '<?mso-application progid="Excel.Sheet"?>\n')
    return (header + body).encode("utf-8")


def sheets_to_simple_xml(sheets: Dict[str, pd.DataFrame]) -> bytes:
    root = ET.Element("workbook")
    for sheet_name, df in sheets.items():
        sheet_el = ET.SubElement(root, "sheet", name=str(sheet_name))
        # headers
        headers_el = ET.SubElement(sheet_el, "headers")
        col_tags = [sanitize_tag(c) for c in df.columns.astype(str)]
        for h in df.columns.astype(str):
            col_el = ET.SubElement(headers_el, "column")
            col_el.text = str(h)
        # rows
        rows_el = ET.SubElement(sheet_el, "rows")
        for i, row in df.iterrows():
            row_el = ET.SubElement(rows_el, "row", index=str(int(i) + 1))
            for tag, val in zip(col_tags, row):
                cell_el = ET.SubElement(row_el, tag)
                cell_el.text = clean_cell(val)
    rough = ET.tostring(root, encoding="utf-8")
    pretty = minidom.parseString(rough).toprettyxml(indent="  ", encoding="utf-8")
    return pretty


def sheets_to_xml(sheets: Dict[str, pd.DataFrame], estilo: str = "2003") -> bytes:
    """estilo='2003' (por defecto, compatible Excel 2003) o 'simple' (genérico)."""
    if estilo == "simple":
        return sheets_to_simple_xml(sheets)
    return sheets_to_xml_2003(sheets)


# ---------------- CSV ----------------
def sheet_to_csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False, encoding="utf-8")
    return "\ufeff".encode("utf-8") + buf.getvalue().encode("utf-8")  # BOM para Excel


def sheets_to_csv_zip(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for sheet_name, df in sheets.items():
            safe = re.sub(r'[\\/*?:"<>|]', "_", str(sheet_name))[:50] or "hoja"
            z.writestr(f"{safe}.csv", sheet_to_csv_bytes(df))
    buf.seek(0)
    return buf.read()


# ---------------- PDF ----------------
def sheets_to_pdf(sheets: Dict[str, pd.DataFrame], title: str = "Reporte Excel") -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=30, rightMargin=30, topMargin=40, bottomMargin=30,
                            title=title)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(title, styles["Title"]))
    story.append(Spacer(1, 12))

    total_sheets = len(sheets)
    for idx, (sheet_name, df) in enumerate(sheets.items()):
        story.append(Paragraph(f"Hoja: {sheet_name} ({len(df)} filas x {len(df.columns)} columnas)",
                               styles["Heading2"]))
        story.append(Spacer(1, 6))

        df_cut = df.head(MAX_PDF_ROWS_PER_SHEET).copy()
        # Convertir todo a string corto para que quepa
        data = [[Paragraph(f"<b>{str(c)[:40]}</b>", styles["Normal"]) for c in df_cut.columns.astype(str)]]
        for _, row in df_cut.iterrows():
            data.append([Paragraph(clean_cell(v)[:60], styles["Normal"]) for v in row])

        if len(data) == 1:
            data.append([Paragraph("(hoja vacía)", styles["Normal"])] * max(len(df_cut.columns), 1))

        col_count = len(data[0])
        # Ancho aproximado repartido
        page_width = landscape(A4)[0] - 60
        col_width = max(60, min(150, page_width / max(col_count, 1)))
        table = Table(data, colWidths=[col_width] * col_count, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eaf1f8")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(table)

        if len(df) > MAX_PDF_ROWS_PER_SHEET:
            story.append(Spacer(1, 6))
            story.append(Paragraph(
                f"<i>Nota: se muestran las primeras {MAX_PDF_ROWS_PER_SHEET} de {len(df)} filas.</i>",
                styles["Normal"]))

        if idx < total_sheets - 1:
            story.append(PageBreak())

    doc.build(story)
    buf.seek(0)
    return buf.read()
