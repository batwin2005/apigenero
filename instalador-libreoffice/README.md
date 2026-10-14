# Instalador de LibreOffice por consola (JAR)

JAR autocontenido, **sin dependencias** (compilado con JDK 11, `javac` directo).
Descarga LibreOffice y lo deja en la ruta que le indiques, por consola.

## Uso en Linux

```bash
# Recomendado: AppImage oficial (un solo archivo, sin root, sin instalación)
java -jar LibreOfficeInstaller.jar \
  --url https://descargas.ejemplo/LibreOffice.AppImage \
  --dest /opt/libreoffice

# El archivo queda en /opt/libreoffice/LibreOffice.AppImage, con permiso +x
# y el JAR verifica el arranque con --version.

# Si el archivo llegó por FTP: bájenlo con lftp y apunten al archivo local
lftp -u usuario,clave servidor -e "get LibreOffice.AppImage -o /tmp/; bye"
java -jar LibreOfficeInstaller.jar \
  --url file:///tmp/LibreOffice.AppImage \
  --dest /opt/libreoffice
```

## Opciones

| Opción | Descripción |
|---|---|
| `--url URL` | `http(s)://...` o `file:///...` del AppImage/instalador (obligatorio) |
| `--dest RUTA` | Carpeta destino (se crea si no existe) (obligatorio) |
| `--name NOMBRE` | Nombre del archivo destino (por defecto: el de la URL) |
| `--solo-descargar` | Solo descarga, sin permiso de ejecución ni verificación |
| `--no-verificar` | Descarga + permiso de ejecución, omite la prueba de arranque |
| `--timeout SEG` | Timeout de lectura (defecto 300) |

## Recompilar

```bash
./build.sh        # Linux (requiere JDK)
build.bat         # Windows (requiere JDK)
```

## Notas

- En servidores sin FUSE, el AppImage se usa así:
  `/opt/libreoffice/LibreOffice.AppImage --appimage-extract-and-run --headless --convert-to pdf archivo.xlsm`
- Este instalador es **opcional**: el API convierte XLSM→XML/CSV/PDF sin LibreOffice.
  Solo instálalo si quieres PDFs con fidelidad total del formato original.
