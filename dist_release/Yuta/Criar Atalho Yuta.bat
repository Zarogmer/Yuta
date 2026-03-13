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
    echo [ERRO] Nao foi encontrado o executavel Yuta.exe em %APP_DIR% nem em subpastas.
    echo [DICA] Copie a pasta completa do build e tente novamente.
    pause
    exit /b 1
)
for %%I in ("%EXE_PATH%") do set "EXE_DIR=%%~dpI"
set "EXE_PATH_PS=%EXE_PATH:\=\\%"
set "EXE_DIR_PS=%EXE_DIR:\=\\%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$exe='%EXE_PATH_PS%'; $work='%EXE_DIR_PS%'; $lnk=Join-Path $env:USERPROFILE 'Desktop\\Yuta.lnk'; $w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut($lnk); $s.TargetPath=$exe; $s.WorkingDirectory=$work; $s.IconLocation=$exe + ',0'; $s.Save()"
echo [OK] Atalho criado na Area de Trabalho: %USERPROFILE%\Desktop\Yuta.lnk
endlocal
