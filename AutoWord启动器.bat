@echo off
chcp 65001 >nul
title AutoWord 文档自动化系统
color 0A

echo.
echo   ___        _        _    _               _ 
echo  / _ \      ^| ^|      ^| ^|  ^| ^|             ^| ^|
echo ^| ^|_^| ^|_   _^| ^|_ ___ ^| ^|  ^| ^| ___  _ __ __^| ^|
echo ^|  _  ^| ^| ^| ^| __/ _ \^| ^|/\^| ^|/ _ \^| '__/ _` ^|
echo ^| ^| ^| ^| ^|_^| ^| ^|^| (_) \  /\  / (_) ^| ^| ^| (_^| ^|
echo \_^| ^|_/\__,_^|\__\___/ \/  \/ \___/^|_^|  \__,_^|
echo.
echo 文档自动化系统 - 开箱即用版
echo.

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到Python，请安装Python 3.8+
    pause
    exit /b 1
)

REM 安装依赖
echo 📦 安装必要依赖...
pip install -q pydantic requests pywin32 >nul 2>&1

echo ✅ 系统就绪！
echo.
echo 🎯 功能特点:
echo   • 智能批注解析
echo   • 四重格式保护
echo   • 安全试运行模式
echo   • 支持GPT-5和Claude 3.7
echo.
echo 📄 目标文档: 郭宇睿-现代汉语言文学研究论文
echo.

REM 运行AutoWord
python autoword_final.py

pause