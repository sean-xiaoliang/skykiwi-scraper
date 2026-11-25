@echo off
chcp 65001 >nul
REM ==== 一键打包为 Windows exe (先跑测试, 不通过则不打包) ====
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
python test_geo.py  || goto :failed
python test_core.py || goto :failed
python test_features.py || goto :failed
python test_gui.py  || goto :failed
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name SkykiwiScraper ^
  --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module lxml --exclude-module PIL ^
  --exclude-module scipy --exclude-module IPython --exclude-module pytest ^
  skykiwi_scraper.py
echo.
echo 打包完成: %~dp0dist\SkykiwiScraper.exe
pause
exit /b 0

:failed
echo.
echo 测试未通过, 已停止打包。
pause
exit /b 1
