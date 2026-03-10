@echo off
set PYI=.venv\Scripts\pyinstaller.exe
if not exist "%PYI%" set PYI=pyinstaller
"%PYI%" --clean --noconfirm --onedir --windowed --name Yuta ^
--icon assets/icons/gear.ico ^
--add-data "assets;assets" ^
--add-data "data;data" ^
--add-data "config;config" ^
backend/app/main.py
