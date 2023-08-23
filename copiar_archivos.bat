@echo off
rem Copiar directorios y archivos desde /app a /dist/SAGOT_Vendedor/app
xcopy /E /I /Y "app\ui" "dist\SAGOT_Vendedor\app\ui"
xcopy /E /I /Y "app\icono_imagen" "dist\SAGOT_Vendedor\app\icono_imagen"
xcopy /E /I /Y "app\formatos" "dist\SAGOT_Vendedor\app\formatos"
xcopy /Y "app\config.json" "dist\SAGOT_Vendedor\app\config.json"

rem Crear directorios dentro de /dist/main/app
mkdir "dist\SAGOT_Vendedor\app\ordenes"
mkdir "dist\SAGOT_Vendedor\app\reingresos"
mkdir "dist\SAGOT_Vendedor\app\informes"

pause