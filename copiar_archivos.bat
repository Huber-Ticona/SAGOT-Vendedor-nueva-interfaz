@echo off
rem Copiar directorios y archivos desde /app a /dist/main/app
xcopy /E /I /Y "app\ui" "dist\main\app\ui"
xcopy /E /I /Y "app\icono_imagen" "dist\main\app\icono_imagen"
xcopy /E /I /Y "app\formatos" "dist\main\app\formatos"
xcopy /Y "app\config.json" "dist\main\app\config.json"

rem Crear directorios dentro de /dist/main/app
mkdir "dist\main\app\ordenes"
mkdir "dist\main\app\reingresos"
mkdir "dist\main\app\informes"

pause
