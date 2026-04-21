@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "文档比对Web.exe" (
    start "" "文档比对Web.exe"
    exit /b 0
)

where py >nul 2>&1
if %errorlevel% equ 0 (
    py -3 web_app.py
    pause
    exit /b %errorlevel%
)

where python >nul 2>&1
if %errorlevel% equ 0 (
    python web_app.py
    pause
    exit /b %errorlevel%
)

echo [错误] 未找到 Python。请安装 Python 3 并勾选“添加到 PATH”。
echo 依赖包：pip install flask openpyxl lxml
pause
exit /b 1
