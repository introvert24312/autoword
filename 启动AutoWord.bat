@echo off
chcp 65001 >nul
title AutoWord 文档自动化工具

echo.
echo ========================================
echo      AutoWord 文档自动化工具
echo ========================================
echo.
echo 正在启动GUI界面...
echo.

REM 直接启动GUI，不进行复杂检查
python main_gui.py

REM 如果出错，显示简单提示
if errorlevel 1 (
    echo.
    echo 启动失败，请确保：
    echo 1. Python已安装
    echo 2. 已运行: pip install -r requirements.txt
    echo.
    pause
)