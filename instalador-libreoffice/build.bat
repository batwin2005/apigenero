@echo off
REM Recompila el JAR en Windows (solo requiere JDK).
cd /d %~dp0
javac Installer.java
echo Main-Class: Installer> MANIFEST.MF
jar cfm LibreOfficeInstaller.jar MANIFEST.MF Installer.class "Installer$1.class"
del Installer.class "Installer$1.class"
dir LibreOfficeInstaller.jar
