@echo off
chcp 65001 >nul
REM ==== 一键打包为 Windows exe ====
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name SkykiwiScraper ^
  --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module lxml --exclude-module PIL ^
  --exclude-module scipy --exclude-module IPython --exclude-module pytest ^
  skykiwi_scraper.py
echo.
echo 打包完成: %~dp0dist\SkykiwiScraper.exe
pause
