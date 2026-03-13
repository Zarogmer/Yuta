@echo off
setlocal EnableDelayedExpansion
set PYI=.venv\Scripts\pyinstaller.exe
if not exist "%PYI%" set PYI=pyinstaller
set DISTDIR=dist_release
set WORKDIR=build_release
set OCR_DATA=

call :try_remove_dir "%DISTDIR%\Yuta"
if exist "%DISTDIR%\Yuta" (
	set "DISTDIR=dist_release_alt"
	echo [AVISO] A pasta "dist_release\Yuta" esta em uso.
	echo [AVISO] Build sera gerado em "!DISTDIR!\Yuta".
)

call :try_remove_dir "%WORKDIR%\Yuta"
if exist "%WORKDIR%\Yuta" (
	set "WORKDIR=build_release_alt"
	echo [AVISO] A pasta "build_release\Yuta" esta em uso.
	echo [AVISO] Workdir alternativo: "!WORKDIR!".
)

if exist "tesseract" (
	set OCR_DATA=%OCR_DATA% --add-data "tesseract;tesseract"
) else (
	echo [AVISO] Pasta "tesseract" nao encontrada na raiz do projeto.
	echo [AVISO] OCR no cliente vai falhar sem tesseract.exe e tessdata.
)

if exist "poppler" (
	set OCR_DATA=%OCR_DATA% --add-data "poppler;poppler"
) else (
	echo [AVISO] Pasta "poppler" nao encontrada na raiz do projeto.
	echo [AVISO] Leitura de PDF imagem no cliente vai falhar sem pdftoppm.exe/pdfinfo.exe.
)

"%PYI%" --clean --noconfirm --onedir --windowed --name Yuta ^
--contents-directory . ^
--icon assets/icons/gear.ico ^
--paths . ^
--collect-submodules backend ^
--collect-all holidays ^
--distpath "%DISTDIR%" ^
--workpath "%WORKDIR%" ^
--add-data "assets;assets" ^
--add-data "data;data" ^
--add-data "config;config" ^
%OCR_DATA% ^
backend/app/main.py

if errorlevel 1 goto :eof

set "ATALHO_SCRIPT=%DISTDIR%\Yuta\Criar Atalho Yuta.bat"
(
echo @echo off
echo setlocal
echo set "APP_DIR=%%~dp0"
echo set "EXE_PATH="
echo if exist "%%APP_DIR%%Yuta.exe" set "EXE_PATH=%%APP_DIR%%Yuta.exe"
echo if not defined EXE_PATH ^(
echo ^    for /f "delims=" %%%%I in ^('dir /b /s "%%APP_DIR%%Yuta.exe" 2^^^>nul'^) do ^(
echo ^        set "EXE_PATH=%%%%I"
echo ^        goto :found
echo ^    ^)
echo ^)
echo if not defined EXE_PATH ^(
echo ^    for /f "delims=" %%%%I in ^('dir /b /s "%%APP_DIR%%..\Yuta.exe" 2^^^>nul'^) do ^(
echo ^        set "EXE_PATH=%%%%I"
echo ^        goto :found
echo ^    ^)
echo ^)
echo if not defined EXE_PATH ^(
echo ^    for /f "delims=" %%%%I in ^('dir /b /s "%%APP_DIR%%..\..\Yuta.exe" 2^^^>nul'^) do ^(
echo ^        set "EXE_PATH=%%%%I"
echo ^        goto :found
echo ^    ^)
echo ^)
echo :found
echo if not defined EXE_PATH ^(
echo ^    echo [ERRO] Nao foi encontrado o executavel Yuta.exe em %%APP_DIR%% nem em subpastas.
echo ^    echo [DICA] Copie a pasta completa do build e tente novamente.
echo ^    pause
echo ^    exit /b 1
echo ^)
echo for %%%%I in ^("%%EXE_PATH%%"^) do set "EXE_DIR=%%%%~dpI"
echo set "EXE_PATH_PS=%%EXE_PATH:\=\\%%"
echo set "EXE_DIR_PS=%%EXE_DIR:\=\\%%"
echo powershell -NoProfile -ExecutionPolicy Bypass -Command "$exe='%%EXE_PATH_PS%%'; $work='%%EXE_DIR_PS%%'; $lnk=Join-Path $env:USERPROFILE 'Desktop\\Yuta.lnk'; $w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut($lnk); $s.TargetPath=$exe; $s.WorkingDirectory=$work; $s.IconLocation=$exe + ',0'; $s.Save()"
echo echo [OK] Atalho criado na Area de Trabalho: %%USERPROFILE%%\Desktop\Yuta.lnk
echo endlocal
) > "%ATALHO_SCRIPT%"

echo [OK] Script de atalho gerado em: %ATALHO_SCRIPT%

set "START_SCRIPT=%DISTDIR%\Yuta\Iniciar Yuta.bat"
(
echo @echo off
echo setlocal
echo set "APP_DIR=%%~dp0"
echo set "EXE_PATH="
echo if exist "%%APP_DIR%%Yuta.exe" set "EXE_PATH=%%APP_DIR%%Yuta.exe"
echo if not defined EXE_PATH ^(
echo ^    for /f "delims=" %%%%I in ^('dir /b /s "%%APP_DIR%%Yuta.exe" 2^^^>nul'^) do ^(
echo ^        set "EXE_PATH=%%%%I"
echo ^        goto :found
echo ^    ^)
echo ^)
echo if not defined EXE_PATH ^(
echo ^    for /f "delims=" %%%%I in ^('dir /b /s "%%APP_DIR%%..\Yuta.exe" 2^^^>nul'^) do ^(
echo ^        set "EXE_PATH=%%%%I"
echo ^        goto :found
echo ^    ^)
echo ^)
echo if not defined EXE_PATH ^(
echo ^    for /f "delims=" %%%%I in ^('dir /b /s "%%APP_DIR%%..\..\Yuta.exe" 2^^^>nul'^) do ^(
echo ^        set "EXE_PATH=%%%%I"
echo ^        goto :found
echo ^    ^)
echo ^)
echo :found
echo if not defined EXE_PATH ^(
echo ^    echo [ERRO] Yuta.exe nao encontrado em %%APP_DIR%% nem em subpastas.
echo ^    echo [DICA] Copie a pasta completa do build e mantenha a estrutura do dist_release.
echo ^    pause
echo ^    exit /b 1
echo ^)
echo for %%%%I in ^("%%EXE_PATH%%"^) do set "EXE_DIR=%%%%~dpI"
echo pushd "%%EXE_DIR%%"
echo start "Yuta" "%%EXE_PATH%%"
echo popd
echo endlocal
) > "%START_SCRIPT%"

echo [OK] Script de inicializacao gerado em: %START_SCRIPT%

goto :eof

:try_remove_dir
set "TARGET_DIR=%~1"
if not exist "%TARGET_DIR%" goto :eof

rmdir /s /q "%TARGET_DIR%" >nul 2>&1
if not exist "%TARGET_DIR%" goto :eof

for /l %%R in (1,1,3) do (
	timeout /t 1 /nobreak >nul
	rmdir /s /q "%TARGET_DIR%" >nul 2>&1
	if not exist "%TARGET_DIR%" goto :eof
)

goto :eof
