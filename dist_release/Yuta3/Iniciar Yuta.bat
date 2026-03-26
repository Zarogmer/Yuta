@echo off
setlocal
set "APP_DIR=%~dp0"
set "EXE_PATH="
if exist "%APP_DIR%Yuta.exe" set "EXE_PATH=%APP_DIR%Yuta.exe"
if not defined EXE_PATH (
    for /f "delims=" %%I in ('dir /b /s "%APP_DIR%Yuta.exe" 2^>nul') do (
        set "EXE_PATH=%%I"
        goto :found
    )
)
if not defined EXE_PATH (
    for /f "delims=" %%I in ('dir /b /s "%APP_DIR%..\Yuta.exe" 2^>nul') do (
        set "EXE_PATH=%%I"
        goto :found
    )
)
if not defined EXE_PATH (
    for /f "delims=" %%I in ('dir /b /s "%APP_DIR%..\..\Yuta.exe" 2^>nul') do (
        set "EXE_PATH=%%I"
        goto :found
    )
)
:found
if not defined EXE_PATH (
    echo [ERRO] Yuta.exe nao encontrado em %APP_DIR% nem em subpastas.
    echo [DICA] Copie a pasta completa do build e mantenha a estrutura do dist_release.
    pause
    exit /b 1
)
for %%I in ("%EXE_PATH%") do set "EXE_DIR=%%~dpI"
pushd "%EXE_DIR%"
start "Yuta" "%EXE_PATH%"
popd
endlocal
