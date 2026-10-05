@echo off
chcp 65001 >nul
cd /d "%~dp0\.."

python -m pip install -r requirements.txt -r requirements-dev.txt
if errorlevel 1 exit /b 1

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "答辩专用插件" ^
  --icon "assets\icon.ico" ^
  --add-data "assets\icon.ico;." ^
  --add-data "assets\icon.png;." ^
  --hidden-import=win32com.client ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  --hidden-import=win32timezone ^
  --hidden-import=docx ^
  --collect-submodules=docx ^
  notes_overlay.py

if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo.
echo Done: dist\答辩专用插件.exe
explorer dist
