#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC_DIR="$SCRIPT_DIR/src"

# Try to find virtual environment in common locations
if [ -f "$SCRIPT_DIR/venv/bin/activate" ]; then
    VENV_DIR="$SCRIPT_DIR/venv"
elif [ -f "$SCRIPT_DIR/../venv/bin/activate" ]; then
    VENV_DIR="$SCRIPT_DIR/../venv"
else
    VENV_DIR=""
fi

echo "=========================================="
echo "  文档页数统计工具 - macOS 构建脚本"
echo "=========================================="

# Activate virtual environment
if [ -n "$VENV_DIR" ] && [ -f "$VENV_DIR/bin/activate" ]; then
    source "$VENV_DIR/bin/activate"
    echo "[1/4] 虚拟环境已激活 ($VENV_DIR)"
else
    echo "[错误] 找不到虚拟环境。"
    echo "提示: 在项目目录或其父目录下运行:"
    echo "  python3 -m venv venv"
    echo "  source venv/bin/activate"
    echo "  pip install -r src/requirements.txt"
    exit 1
fi

# Clean old build artifacts
echo "[2/4] 清理旧的构建文件..."
rm -rf "$SRC_DIR/build"
rm -rf "$SRC_DIR/dist"
rm -rf "$SCRIPT_DIR/dist"
mkdir -p "$SCRIPT_DIR/dist"

# Build with PyInstaller
echo "[3/4] 开始打包..."
cd "$SRC_DIR"
pyinstaller --windowed --name PageCounter --clean page_counter.py

# Zip the app
echo "[4/4] 压缩并分发..."
cd "$SRC_DIR/dist"
zip -rq PageCounter_macOS.zip PageCounter.app

# Copy to repo/dist
cp PageCounter_macOS.zip "$SCRIPT_DIR/dist/"

# Clean up temporary files in src/
rm -rf "$SRC_DIR/build"
rm -rf "$SRC_DIR/dist"

echo ""
echo "=========================================="
echo "  构建完成！"
echo "=========================================="
echo "产物位置:"
echo "  - dist/PageCounter_macOS.zip"
echo ""
