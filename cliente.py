"""Cliente del API Convertidor.

Envías el FORMATO al que quieres transformar (xml/csv/pdf)
y le dices en qué RUTA guardar la descarga.

Uso:
    pip install requests
    py cliente.py --archivo ejemplos/ejemplo.xlsm --formato xml --salida C:/salidas/reporte.xml
    py cliente.py --archivo datos.xlsx --formato pdf --salida ./salidas/reporte.pdf
    py cliente.py --archivo datos.xlsx --formato csv --salida ./salidas/datos.csv --hoja Hoja1

    # Guardar directo en disco del SERVIDOR (solo si el API corre en tu máquina):
    py cliente.py --archivo datos.xlsm --formato pdf --ruta-destino C:/salidas/desde_servidor.pdf
"""
from __future__ import annotations

import argparse
from pathlib import Path

import requests


def main() -> None:
    p = argparse.ArgumentParser(description="Convierte un Excel con el API y lo guarda en la ruta que le digas.")
    p.add_argument("--url", default="http://localhost:8000", help="URL base del API")
    p.add_argument("--archivo", required=True, help="Ruta del Excel a convertir (.xls, .xlsx, .xlsm, ...)")
    p.add_argument("--formato", required=True, choices=["xml", "csv", "pdf"],
                   help="Formato al que se va a transformar")
    p.add_argument("--salida", default=None,
                   help="RUTA local donde guardar la descarga. Ej: C:/salidas/reporte.pdf")
    p.add_argument("--hoja", default=None, help="Nombre de hoja (opcional, por defecto todas)")
    p.add_argument("--xml-estilo", default="2003", choices=["2003", "simple"],
                   help="Estilo XML: 2003 (Spreadsheet 2003, abre en Excel) o simple (generico)")
    p.add_argument("--nombre-salida", default=None, help="Nombre sugerido de descarga en el servidor")
    p.add_argument("--ruta-destino", default=None,
                   help="Ruta ABSOLUTA del SERVIDOR donde guardar (en vez de descargar). "
                        "Solo si el API corre en tu máquina.")
    args = p.parse_args()

    origen = Path(args.archivo)
    if not origen.is_file():
        raise SystemExit(f"No existe el archivo: {origen}")

    params = {"formato": args.formato}
    if args.hoja:
        params["hoja"] = args.hoja
    if args.formato == "xml":
        params["xml_estilo"] = args.xml_estilo
    if args.nombre_salida:
        params["nombre_salida"] = args.nombre_salida
    if args.ruta_destino:
        params["ruta_destino"] = args.ruta_destino

    url = args.url.rstrip("/") + "/api/v1/convertir"
    print(f"Enviando {origen.name} -> formato={args.formato} ...")
    with open(origen, "rb") as f:
        r = requests.post(url, params=params, files={"archivo": (origen.name, f)}, timeout=120)

    if r.status_code != 200:
        print(f"Error {r.status_code}: {r.text[:1000]}")
        raise SystemExit(1)

    # Caso guardado en servidor: responde JSON
    if args.ruta_destino:
        print("Guardado en servidor:", r.json())
        return

    # Caso descarga: guardar en la RUTA que el usuario diga
    if args.salida:
        destino = Path(args.salida).expanduser()
    else:
        # por defecto: mismo nombre junto al origen con nueva extensión
        destino = origen.with_suffix("." + args.formato if args.formato != "csv" else ".csv")

    # Si dan una carpeta, completar con nombre automático
    if str(args.salida or "").endswith(("/", "\\")) or (destino.suffix == "" and args.salida):
        destino = destino / f"{origen.stem}.{args.formato}"

    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(r.content)
    print(f"OK {len(r.content)} bytes descargados en: {destino.resolve()}")


if __name__ == "__main__":
    main()
