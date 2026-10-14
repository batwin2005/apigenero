#!/usr/bin/env bash
# Recompila el JAR (solo requiere JDK, sin Maven ni dependencias).
set -e
cd "$(dirname "$0")"
javac Installer.java
echo "Main-Class: Installer" > MANIFEST.MF
jar cfm LibreOfficeInstaller.jar MANIFEST.MF Installer.class 'Installer$1.class'
rm -f Installer.class 'Installer$1.class'
ls -la LibreOfficeInstaller.jar
