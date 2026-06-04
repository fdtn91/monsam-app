@echo off
chcp 65001 > nul
cd /d "%~dp0"

:: Verificar que node_modules exista
if not exist "node_modules" (
    echo Instalando dependencias, espera un momento...
    npm install --silent
)

:: Lanzar Electron ocultando la consola inmediatamente
start "" /B npx electron . --no-console
exit
