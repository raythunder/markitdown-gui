@echo off
chcp 65001 >nul
echo ====================================================
echo MarkItDown GUI - exe构建脚本
echo ====================================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.8或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo 检测到Python版本:
python --version

echo.
echo 开始构建MarkItDown GUI...
echo.

REM 运行Python构建脚本
python build_exe.py

echo.
echo 构建完成！
echo.

REM 检查是否成功生成exe文件
if exist "dist\MarkItDown-GUI.exe" (
    echo 成功生成exe文件: dist\MarkItDown-GUI.exe
    echo.
    echo 您现在可以运行该exe文件来使用MarkItDown GUI
    echo.
    set /p choice="是否立即运行程序？(y/n): "
    if /i "%choice%"=="y" (
        start "" "dist\MarkItDown-GUI.exe"
    )
) else (
    echo 构建失败，请检查错误信息
)

echo.
pause 