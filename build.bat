@echo off
chcp 65001 > nul
echo ================================
echo  SmartStore Monitor - EXE Build
echo ================================
echo.

echo [Step 1] Checking Python...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found.
    echo Install from python.org - check "Add to PATH"!
    pause
    exit /b 1
)

echo.
echo [Step 2] Installing packages...
python -m pip install requests beautifulsoup4 openpyxl python-dateutil pyinstaller
if %errorlevel% neq 0 (
    echo ERROR: pip install failed.
    pause
    exit /b 1
)

echo.
echo [Step 3] Building EXE...
python -m PyInstaller --onefile --windowed --name "SmartStore_Monitor" main.py
if %errorlevel% neq 0 (
    echo ERROR: Build failed. Check error above.
    pause
    exit /b 1
)

echo.
echo [Step 4] Creating deploy folder...
if not exist "deploy" mkdir "deploy"
if exist "deploy\license.json" del /f /q "deploy\license.json"

if exist "dist\SmartStore_Monitor.exe" (
    copy "dist\SmartStore_Monitor.exe" "deploy\"
    echo.
    echo ================================
    echo  SUCCESS!
    echo  deploy\SmartStore_Monitor.exe
    echo ================================
) else (
    echo ERROR: EXE not found in dist folder.
)

pause
