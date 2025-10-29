@echo off
setlocal enabledelayedexpansion

REM 基于 PyInstaller 打包 Windows 可执行程序（单文件）
REM 生成的 .exe 位于 dist\ExcelTool.exe

where pyinstaller >nul 2>nul
if errorlevel 1 (
  echo 未检测到 pyinstaller，正在安装...
  pip install pyinstaller
)

set APP_NAME=ExcelTool
if exist build rmdir /S /Q build
if exist dist rmdir /S /Q dist
if exist %APP_NAME%.spec del %APP_NAME%.spec

REM 使用 --onefile 生成单文件，--collect-submodules 收集 pandas 依赖
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name %APP_NAME% ^
  --collect-submodules pandas ^
  main.py

echo.
echo 打包完成：dist\%APP_NAME%.exe
echo 将 dist\%APP_NAME%.exe 拷贝到其他 Windows 电脑即可运行（无需安装 Python）。
endlocal

