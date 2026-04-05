@echo off
echo Office Document Copy Utility - EXE 빌드
echo ==========================================

pip install -r requirements.txt

pyinstaller ^
    --onefile ^
    --console ^
    --name OfficeCopy ^
    --clean ^
    office_copy.py

echo.
echo 빌드 완료: dist\OfficeCopy.exe
pause
