@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==============================================================
echo    ETIQUETADOR DE COLORES   (local, sin API, coste cero)
echo ==============================================================
echo.

rem Buscar Python. Primero "python", y si no esta, el lanzador "py".
set PY=
where python >nul 2>nul && set PY=python
if "%PY%"=="" (where py >nul 2>nul && set PY=py)

if "%PY%"=="" (
  echo No encuentro Python en este ordenador.
  echo Instalalo desde https://www.python.org/downloads/ y marca
  echo la casilla "Add Python to PATH" durante la instalacion.
  echo.
  pause
  exit /b 1
)

if "%~1"=="" (
  echo No has arrastrado ningun fichero: se abrira un selector.
  echo.
)

%PY% "%~dp0etiquetar_colores_local.py" %*
set CODIGO=%ERRORLEVEL%

echo.
if not "%CODIGO%"=="0" (
  echo ---------------------------------------------------------------
  echo Ha terminado con algun problema. Lee el mensaje de arriba.
  echo ---------------------------------------------------------------
)
echo Pulsa una tecla para cerrar esta ventana.
pause >nul
endlocal
