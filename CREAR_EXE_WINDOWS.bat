@echo off
setlocal
cd /d "%~dp0"

echo Preparando Estructurador para Windows...
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -m venv .venv-windows
) else (
    python -m venv .venv-windows
)
if errorlevel 1 goto error_python

".venv-windows\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto error_build

".venv-windows\Scripts\python.exe" -m unittest -q
if errorlevel 1 goto error_build

".venv-windows\Scripts\python.exe" compilar_windows.py
if errorlevel 1 goto error_build

echo.
echo Listo: dist\Estructurador.exe
echo Puedes copiar SOLO ese archivo al equipo donde quieras usarlo.
pause
exit /b 0

:error_python
echo.
echo No se pudo preparar Python. Instala Python 3.14 de 64 bits para Windows
echo desde https://www.python.org/downloads/windows/ y vuelve a ejecutar este archivo.
pause
exit /b 1

:error_build
echo.
echo No se pudo generar el ejecutable. Revisa el error que aparece arriba.
pause
exit /b 1
