#!/bin/zsh

# 基于 PyInstaller 打包 macOS 应用
# 生成的 .app 位于 dist/Excel处理小工具.app

set -euo pipefail

SCRIPT_DIR=$(cd "$(dirname "$0")" && pwd)
cd "$SCRIPT_DIR"

APP_NAME="Excel处理小工具"

if ! command -v pyinstaller >/dev/null 2>&1; then
  echo "未检测到 pyinstaller，正在安装..."
  pip install pyinstaller
fi

rm -rf build dist "${APP_NAME}.spec"

pyinstaller \
  --noconfirm \
  --windowed \
  --name "$APP_NAME" \
  main.py

echo "打包完成：dist/${APP_NAME}.app"
echo "注意：首次在未签名情况下运行，需右键-打开绕过 Gatekeeper。"

