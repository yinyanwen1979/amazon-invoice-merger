@echo off
chcp 65001 >nul
echo ============================================
echo   Amazon Invoice 合并工具 - 打包脚本
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8 或以上版本。
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装依赖库...
pip install pandas openpyxl Pillow pyinstaller -q
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络连接。
    pause
    exit /b 1
)
echo       依赖安装完成。
echo.

echo [2/3] 开始打包（约需 1-3 分钟）...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Amazon合并工具" ^
    --add-data "mascot.png;." ^
    --clean ^
    main.py

if errorlevel 1 (
    echo [错误] 打包失败，请查看上方错误信息。
    pause
    exit /b 1
)
echo       打包完成。
echo.

echo [3/3] 清理临时文件...
if exist build   rmdir /s /q build
if exist *.spec  del /q *.spec
echo       清理完成。
echo.

echo ============================================
echo   成功！可执行文件位于：
echo   dist\Amazon合并工具.exe
echo ============================================
echo.

set /p run="是否立即运行程序测试？(y/n): "
if /i "%run%"=="y" (
    start "" "dist\Amazon合并工具.exe"
)

pause
