@echo off
set PYI=.venv\Scripts\pyinstaller.exe
if not exist "%PYI%" set PYI=pyinstaller
set DISTDIR=dist_release
set WORKDIR=build_release

if exist "%DISTDIR%\Yuta" rmdir /s /q "%DISTDIR%\Yuta"
if exist "%WORKDIR%\Yuta" rmdir /s /q "%WORKDIR%\Yuta"

"%PYI%" --clean --noconfirm --onedir --windowed --name Yuta ^
--icon assets/icons/gear.ico ^
--paths . ^
--collect-submodules backend ^
--collect-all holidays ^
--distpath "%DISTDIR%" ^
--workpath "%WORKDIR%" ^
--add-data "assets;assets" ^
--add-data "data;data" ^
--add-data "config;config" ^
backend/app/main.py
