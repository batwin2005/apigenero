import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.PosixFilePermission;
import java.util.EnumSet;
import java.util.Locale;
import java.util.Set;
import java.util.concurrent.TimeUnit;

/**
 * Instalador de LibreOffice por consola, empaquetado como JAR sin dependencias.
 *
 * Descarga LibreOffice (recomendado: el AppImage oficial, un solo archivo que
 * corre sin root ni instalación) y lo deja en la ruta que se le indique.
 *
 * Uso en Linux:
 *   java -jar LibreOfficeInstaller.jar --url &lt;http|https|file&gt; --dest /opt/libreoffice
 *
 * Ejemplo (AppImage 64 bits, ~250 MB):
 *   java -jar LibreOfficeInstaller.jar \
 *     --url https://www.libreoffice.org/download/download/libreoffice-stable/ \
 *     --dest /opt/libreoffice --name LibreOffice.AppImage
 *
 * Nota: si el archivo llega por FTP, primero bájenlo con lftp/curl y luego
 * apunten --url a file:///ruta/local/del/archivo (o pásenlo directo con --url
 * file://...). El JAR también acepta esa forma.
 *
 * Opciones:
 *   --url URL          http(s)://... o file:///... del AppImage/instalador (obligatorio)
 *   --dest RUTA        carpeta destino donde quedará el archivo (obligatorio)
 *   --name NOMBRE      nombre del archivo destino (por defecto: último tramo de la URL)
 *   --solo-descargar   solo descarga, no da permiso de ejecución ni verifica
 *   --no-verificar     descarga + permiso de ejecución, omite la prueba de arranque
 *   --timeout SEG      timeout de lectura en segundos (defecto 300)
 */
public class Installer {

    public static void main(String[] args) throws Exception {
        String url = opt(args, "--url", null);
        String dest = opt(args, "--dest", null);
        String name = opt(args, "--name", null);
        boolean soloDescargar = flag(args, "--solo-descargar");
        boolean noVerificar = flag(args, "--no-verificar");
        int timeout = Integer.parseInt(opt(args, "--timeout", "300"));

        if (url == null || dest == null) {
            System.err.println("Faltan parámetros obligatorios.");
            System.err.println("Uso: java -jar LibreOfficeInstaller.jar --url URL --dest RUTA [--name NOMBRE] [--solo-descargar] [--no-verificar]");
            System.exit(2);
        }
        if (name == null || name.isEmpty()) {
            name = url.substring(url.lastIndexOf('/') + 1);
            if (name.isEmpty()) {
                name = "LibreOffice.AppImage";
            }
        }

        Path destDir = Paths.get(dest);
        Files.createDirectories(destDir);
        Path destino = destDir.resolve(name);
        Path temporal = destDir.resolve(name + ".part");

        System.out.println("Descargando:");
        System.out.println("  Origen : " + url);
        System.out.println("  Destino: " + destino);
        descargar(url, temporal, timeout);
        Files.move(temporal, destino, StandardCopyOption.REPLACE_EXISTING);
        long bytes = Files.size(destino);
        System.out.println("OK descarga completa: " + bytes + " bytes en " + destino);

        if (soloDescargar) {
            return;
        }

        boolean esWindows = System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("win");
        if (!esWindows) {
            darPermisoEjecucion(destino);
        } else {
            System.out.println("Windows detectado: se omite chmod (no aplica).");
        }

        if (noVerificar) {
            return;
        }
        verificar(destino, esWindows);
    }

    private static void descargar(String urlStr, Path temporal, int timeoutSeg) throws IOException {
        URL url = new URL(urlStr);
        String protocol = url.getProtocol().toLowerCase(Locale.ROOT);
        if (!protocol.equals("http") && !protocol.equals("https") && !protocol.equals("file")) {
            throw new IOException("Protocolo no soportado: " + protocol
                    + ". Usa http(s):// o file:// (si lo trajiste por FTP, apunta al archivo local).");
        }

        long total = -1;
        try (InputStream in = abrirStream(url, timeoutSeg);
             OutputStream out = Files.newOutputStream(temporal)) {
            if (protocol.equals("file")) {
                try {
                    total = Files.size(Paths.get(url.toURI()));
                } catch (Exception ignored) {
                    total = -1;
                }
            }
            byte[] buf = new byte[65536];
            long copiados = 0;
            long ultimoReporte = System.currentTimeMillis();
            int n;
            while ((n = in.read(buf)) != -1) {
                out.write(buf, 0, n);
                copiados += n;
                long ahora = System.currentTimeMillis();
                if (ahora - ultimoReporte > 2000) {
                    ultimoReporte = ahora;
                    if (total > 0) {
                        System.out.printf("  ... %d MB de %d MB (%d%%)%n",
                                copiados / 1048576, total / 1048576, copiados * 100 / total);
                    } else {
                        System.out.printf("  ... %d MB descargados%n", copiados / 1048576);
                    }
                }
            }
        }
        if (Files.size(temporal) == 0) {
            throw new IOException("La descarga quedó vacía, revisa la URL.");
        }
    }

    private static InputStream abrirStream(URL url, int timeoutSeg) throws IOException {
        String protocol = url.getProtocol().toLowerCase(Locale.ROOT);
        if (protocol.equals("file")) {
            try {
                return Files.newInputStream(Paths.get(url.toURI()));
            } catch (Exception e) {
                throw new FileNotFoundException("No se pudo abrir el archivo local: " + url + " (" + e + ")");
            }
        }
        // http/https con redirecciones y User-Agent (algunos mirrors lo exigen)
        URL actual = url;
        for (int i = 0; i < 5; i++) {
            HttpURLConnection con = (HttpURLConnection) actual.openConnection();
            con.setConnectTimeout(30000);
            con.setReadTimeout(timeoutSeg * 1000);
            con.setInstanceFollowRedirects(false);
            con.setRequestProperty("User-Agent", "LibreOfficeInstaller/1.0");
            int codigo = con.getResponseCode();
            if (codigo == 301 || codigo == 302 || codigo == 303 || codigo == 307 || codigo == 308) {
                String loc = con.getHeaderField("Location");
                con.disconnect();
                if (loc == null) {
                    throw new IOException("Redirección sin Location desde " + actual);
                }
                actual = new URL(actual, loc);
                continue;
            }
            if (codigo != 200) {
                con.disconnect();
                throw new IOException("HTTP " + codigo + " al descargar " + actual);
            }
            final HttpURLConnection c = con;
            InputStream in = con.getInputStream();
            return new InputStream() {
                @Override public int read() throws IOException { return in.read(); }
                @Override public int read(byte[] b, int off, int len) throws IOException { return in.read(b, off, len); }
                @Override public void close() throws IOException { try { in.close(); } finally { c.disconnect(); } }
            };
        }
        throw new IOException("Demasiadas redirecciones para " + url);
    }

    private static void darPermisoEjecucion(Path destino) {
        try {
            Set<PosixFilePermission> perms = EnumSet.of(
                    PosixFilePermission.OWNER_READ, PosixFilePermission.OWNER_WRITE, PosixFilePermission.OWNER_EXECUTE,
                    PosixFilePermission.GROUP_READ, PosixFilePermission.GROUP_EXECUTE,
                    PosixFilePermission.OTHERS_READ, PosixFilePermission.OTHERS_EXECUTE);
            Files.setPosixFilePermissions(destino, perms);
            System.out.println("OK permiso de ejecución asignado.");
        } catch (UnsupportedOperationException e) {
            // Sistema sin POSIX (raro en Linux): respaldo con chmod
            try {
                new ProcessBuilder("chmod", "+x", destino.toString()).inheritIO().start().waitFor();
                System.out.println("OK permiso de ejecución asignado (chmod).");
            } catch (Exception ex) {
                System.out.println("AVISO: no se pudo dar permiso de ejecución: " + ex);
            }
        } catch (Exception e) {
            System.out.println("AVISO: no se pudo dar permiso de ejecución: " + e);
        }
    }

    private static void verificar(Path destino, boolean esWindows) throws Exception {
        System.out.println("Verificando arranque...");
        String nombre = destino.getFileName().toString().toLowerCase(Locale.ROOT);
        ProcessBuilder pb;
        if (esWindows && (nombre.endsWith(".msi") || nombre.endsWith(".exe"))) {
            System.out.println("Instalador de Windows detectado; este JAR solo lo descarga.");
            System.out.println("Para instalar en silencio ejecuta: msiexec /i \"" + destino + "\" /qn");
            return;
        }
        if (nombre.endsWith(".msi") || nombre.endsWith(".exe")
                || nombre.endsWith(".deb") || nombre.endsWith(".rpm")
                || nombre.endsWith(".tar.gz") || nombre.endsWith(".tgz")) {
            System.out.println("Archivo " + destino.getFileName() + " descargado en " + destino);
            System.out.println("Instálalo con el gestor de paquetes (dpkg -i / dnf install) o extrae el .tar.gz.");
            return;
        }
        // Se asume AppImage o binario: prueba de arranque
        pb = new ProcessBuilder(destino.toString(), "--version");
        pb.redirectErrorStream(true);
        Process p = pb.start();
        boolean termino = p.waitFor(90, TimeUnit.SECONDS);
        String salida = new String(p.getInputStream().readAllBytes());
        if (termino && p.exitValue() == 0) {
            System.out.println("OK LibreOffice responde: " + salida.trim());
        } else {
            System.out.println("AVISO: no se pudo verificar el arranque.");
            System.out.println("Salida: " + salida.trim());
            System.out.println("Si es AppImage en servidor sin FUSE, úsalo así:");
            System.out.println("  " + destino + " --appimage-extract-and-run --headless --version");
        }
    }

    private static String opt(String[] args, String clave, String defecto) {
        for (int i = 0; i < args.length - 1; i++) {
            if (args[i].equals(clave)) {
                return args[i + 1];
            }
        }
        return defecto;
    }

    private static boolean flag(String[] args, String clave) {
        for (String a : args) {
            if (a.equals(clave)) {
                return true;
            }
        }
        return false;
    }
}
