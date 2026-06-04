@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo  ============================================
echo    MONSAN Accesorios - Gestor 3D v1.0.0
echo  ============================================
echo.

:: Verificar que node_modules exista
if not exist "node_modules" (
    echo  Primera ejecucion detectada.
    echo  Instalando dependencias, espera un momento...
    echo.
    npm install
    echo.
    echo  Instalacion completada.
    echo.
)

echo  Iniciando la aplicacion...
echo.
npx electron .
