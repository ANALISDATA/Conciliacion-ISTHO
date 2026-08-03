@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem Git se instalo para este usuario; si aun no esta en el PATH se agrega aqui.
set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;C:\Program Files\Git\cmd;%PATH%"

echo ==========================================================
echo   SUBIR LA APLICACION A GITHUB
echo ==========================================================
echo.
echo  La PRIMERA vez se abrira una ventana del navegador para
echo  que inicies sesion en GitHub y autorices el acceso.
echo  Despues de eso ya no la volvera a pedir.
echo.
echo ----------------------------------------------------------
echo.

git push -u origin main

echo.
echo ----------------------------------------------------------
if %ERRORLEVEL% EQU 0 (
    echo   LISTO. El codigo quedo subido a GitHub.
) else (
    echo   Algo fallo. Copia el mensaje de arriba y enviamelo.
)
echo ----------------------------------------------------------
echo.
pause
