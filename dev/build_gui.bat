@echo off
chcp 65001 >nul
cd /d "%~dp0.."
echo 正在打包 GUI（TransportAgreement_GUI）...
echo.
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 请先安装 PyInstaller: pip install -r dev\requirements-dev.txt
    pause
    exit /b 1
)
pyinstaller --onedir --windowed --name "TransportAgreement_GUI" --clean --paths dev dev\gui.py
if errorlevel 1 (
    echo 打包失败。
    pause
    exit /b 1
)
echo.
echo 打包完成。输出目录: dist\TransportAgreement_GUI\
echo 将 EmailAddress.csv、smtp_config.ini、Signature.txt 等放在该目录下即可使用。
pause
