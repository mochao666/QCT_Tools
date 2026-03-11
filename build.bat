@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PROJECT_DIR=%CD%"
set "DIST_DIR=%PROJECT_DIR%\dist\QCT_Tools"
echo 当前项目目录: %PROJECT_DIR%
echo 打包输出将到: %DIST_DIR%
echo.
echo 正在打包 QCT小工具，请确保已安装: pip install pyinstaller
echo.

pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 未检测到 PyInstaller，正在安装...
    pip install pyinstaller
)

REM 卸载与 PyInstaller 不兼容的 pathlib 旧版 backport（若存在）
pip show pathlib >nul 2>&1
if not errorlevel 1 (
    echo 正在卸载与 PyInstaller 不兼容的 pathlib 包...
    pip uninstall -y pathlib
    echo.
)

REM 每次打包前重新生成 exe 图标 icon.ico（放大镜+对勾，深蓝+浅灰）
echo 正在生成图标 icon.ico ...
pip show Pillow >nul 2>&1
if errorlevel 1 pip install Pillow
python make_icon.py
if not exist "icon.ico" (
    echo 警告: icon.ico 未生成，exe 将使用默认图标。
) else (
    echo 图标已更新。
)
echo.

pyinstaller --noconfirm --clean --distpath "%PROJECT_DIR%\dist" --workpath "%PROJECT_DIR%\build" QCT_Tools.spec
if errorlevel 1 (
    echo 打包失败，请查看上方错误信息。
    pause
    exit /b 1
)

if exist "%DIST_DIR%" (
    copy /Y "使用说明.txt" "%DIST_DIR%\使用说明.txt" >nul 2>&1
    echo.
    echo ========== 打包成功 ==========
    echo 可执行程序位置:
    echo   %DIST_DIR%\QCT_Tools.exe
    echo.
    echo 请将整个文件夹压缩后分发:
    echo   %DIST_DIR%
    echo.
    echo 若 exe 图标未更新：请关闭资源管理器后重新打开，或将 exe 复制到新文件夹再查看（Windows 会缓存图标）。
    echo.
    start "" "%DIST_DIR%"
) else (
    echo.
    echo 未找到 dist\QCT_Tools，打包可能未完成。当前 dist 目录内容:
    dir /b "%PROJECT_DIR%\dist" 2>nul || echo dist 目录不存在
    echo.
)
pause
