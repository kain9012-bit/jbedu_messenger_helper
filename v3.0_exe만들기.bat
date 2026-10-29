@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/2] 필요한 모듈 설치 확인
py -3 -m pip install -q pyinstaller pywinauto pywin32 openpyxl pillow 2>nul || python -m pip install -q pyinstaller pywinauto pywin32 openpyxl pillow
echo [2/2] exe 만들기
py -3 -m PyInstaller --noconfirm "jbedu_messenger_helper(v3.0).spec" 2>nul || python -m PyInstaller --noconfirm "jbedu_messenger_helper(v3.0).spec"
echo.
echo 완료: dist 폴더의 "전북특별자치도교육청_메신저 작업 도우미(v3.0).exe"
pause
