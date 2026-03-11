@echo off
chcp 65001 >nul
REM 默认使用本 bat 所在目录；若将 qct_app.bat 放到别处使用，请设置环境变量 QCT_TOOLS_DIR 指向 QCT_Tools 文件夹
if defined QCT_TOOLS_DIR (
    set "QCT_DIR=%QCT_TOOLS_DIR%"
) else (
    set "QCT_DIR=%~dp0"
)
if not "%QCT_DIR:~-1%"=="\" if not "%QCT_DIR:~-1%"=="/" set "QCT_DIR=%QCT_DIR%\"
cd /d "%QCT_DIR%"
python "%QCT_DIR%app_gui.py"
if errorlevel 1 pause
