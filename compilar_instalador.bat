@echo off
cd /d "%~dp0"

echo ============================================
echo   HM Consulting - Compilador v1.0.2
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 ( echo ERROR: Python no instalado. & pause & exit /b 1 )

echo [1/4] Instalando dependencias...
python -m pip install pandas openpyxl requests pillow pyinstaller cryptography --quiet
echo     OK

echo.
echo [2/4] Compilando app principal...
python -m PyInstaller --onefile --windowed ^
    --name "CargadorHorasRedmine" ^
    --icon "HM_Icono.ico" ^
    --add-data "logo_app.png;." ^
    --add-data "logo_instalador.png;." ^
    --add-data "HM_Icono.ico;." ^
    --add-data "icono_acerca.png;." ^
    --add-data "Carga_Horas_-_c_ID_Ticket.xlsx;." ^
    --add-data "Carga_Horas_-_sin_ID_Ticket.xlsx;." ^
    app.py
if errorlevel 1 ( echo ERROR compilando app. & pause & exit /b 1 )
echo     OK

echo.
echo [3/4] Compilando instalador...
python -m PyInstaller --onefile --windowed ^
    --name "Setup_CargadorHoras" ^
    --icon "HM_Icono.ico" ^
    --add-data "logo_instalador.png;." ^
    --add-data "logo_app.png;." ^
    --add-data "HM_Icono.ico;." ^
    --add-data "dist\CargadorHorasRedmine.exe;." ^
    instalador.py
if errorlevel 1 ( echo ERROR compilando instalador. & pause & exit /b 1 )
echo     OK

echo.
echo [4/4] Armando paquete final...
if not exist "Paquete_Instalador" mkdir "Paquete_Instalador"
copy dist\Setup_CargadorHoras.exe Paquete_Instalador\
echo     OK

echo.
echo ============================================
echo   Listo! Compartir: Paquete_Instalador\
echo   Setup_CargadorHoras.exe
echo ============================================
pause
