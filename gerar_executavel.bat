@echo off
rem Gera o executavel do Agente GO (pasta dist\Agente GO).
rem Depois rode: py empacotar.py   (gera o .zip para enviar aos colegas)
cd /d "%~dp0"
py -m PyInstaller --noconfirm --clean --windowed --name "Agente GO" ^
  --icon "src\icone.ico" --add-data "src\icone.ico;." ^
  --paths src --collect-all customtkinter ^
  --hidden-import keyring.backends.Windows ^
  --exclude-module matplotlib --exclude-module numpy --exclude-module pandas ^
  src\app.py
if errorlevel 1 (echo FALHOU & exit /b 1)
echo OK: dist\Agente GO\Agente GO.exe
