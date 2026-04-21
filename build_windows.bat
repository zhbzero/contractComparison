@echo off
REM 在 Windows 电脑上双击本脚本：安装依赖并打包为单文件 exe。
REM 重要：必须在 Windows 上执行，生成的 exe 只能在 Windows 上运行。
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>&1
if %errorlevel% equ 0 (
    set "PIP=py -3 -m pip"
    set "PYI=py -3 -m PyInstaller"
    goto :install
)
where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PIP=python -m pip"
    set "PYI=python -m PyInstaller"
    goto :install
)

echo [错误] 未找到 Python。请先安装 Python 3 并勾选“添加到 PATH”。
pause
exit /b 1

:install
echo 正在安装/更新打包依赖...
%PIP% install -U pip
%PIP% install pyinstaller openpyxl lxml

echo.
echo 正在打包（单文件）...
%PYI% --noconfirm contract_compare.spec
%PYI% --noconfirm contract_web.spec

if %errorlevel% neq 0 (
    echo [错误] 打包失败，请查看上方日志。
    pause
    exit /b 1
)

echo.
echo 完成。可执行文件：
echo - dist\文档比对.exe
echo - dist\文档比对Web.exe
echo 离线比对：将「文档比对.exe」与两份 docx 放在同一文件夹运行。
echo 网页版：运行「文档比对Web.exe」后打开 http://127.0.0.1:8080
pause
