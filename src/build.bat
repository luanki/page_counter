@echo off
chcp 65001 >nul
echo ==========================================
echo   文档页数统计工具 - Windows 打包脚本
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.10 或更高版本。
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 创建虚拟环境...
python -m venv venv
if errorlevel 1 (
    echo [错误] 创建虚拟环境失败。
    pause
    exit /b 1
)

echo [2/3] 安装依赖...
venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 安装依赖失败。
    pause
    exit /b 1
)

echo [3/3] 打包可执行文件...
venv\Scripts\pyinstaller --windowed --onefile --name PageCounter page_counter.py
if errorlevel 1 (
    echo [错误] 打包失败。
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   打包完成！
echo ==========================================
echo 可执行文件位于: dist\PageCounter.exe
echo 直接双击即可运行，无需安装其他软件。
echo.
pause
