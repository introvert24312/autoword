@echo off
chcp 65001 >nul
echo ========================================
echo      AutoWord vNext 性能测试启动器
echo ========================================
echo.

echo 🚀 正在启动AutoWord GUI界面...
echo.

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python
    echo 请安装Python 3.8+并添加到PATH
    pause
    exit /b 1
)

echo ✅ Python环境检查通过

REM 检查基础依赖
echo 📦 检查基础依赖...
python -c "import tkinter; import threading; import json" >nul 2>&1
if errorlevel 1 (
    echo ❌ 基础依赖缺失，请检查Python安装
    pause
    exit /b 1
)

echo ✅ 基础依赖检查完成

REM 启动简化版GUI
echo.
echo 🎯 启动AutoWord简化版GUI界面...
echo 💡 提示: 使用Tkinter界面，无需额外依赖
echo.

python 简化版GUI.py

if errorlevel 1 (
    echo.
    echo ❌ GUI启动失败
    echo.
    echo 🔧 故障排除:
    echo 1. 确保Microsoft Word已安装
    echo 2. 检查Python版本 ^>= 3.8
    echo 3. 运行: pip install -r requirements.txt
    echo.
    pause
) else (
    echo.
    echo ✅ AutoWord GUI已正常关闭
)

echo.
pause