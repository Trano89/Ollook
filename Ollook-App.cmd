@echo off
REM Lance Ollook. L'executable compile d'abord : il porte l'icone de
REM l'application et ne depend d'aucune installation de Python.
setlocal
cd /d "%~dp0"

if exist "dist\Ollook.exe" (
    start "" "dist\Ollook.exe" %*
    goto :end
)

echo dist\Ollook.exe est absent - repli sur le script Python.
echo Pour construire l'executable :  python build_exe.py
echo.

where pythonw >nul 2>&1
if %ERRORLEVEL%==0 (
    start "" pythonw ollook_app.py %*
    goto :end
)

where python >nul 2>&1
if %ERRORLEVEL%==0 (
    start "" python ollook_app.py %*
    goto :end
)

echo Python 3 est introuvable.
echo Installez-le depuis https://www.python.org/downloads/
pause

:end
