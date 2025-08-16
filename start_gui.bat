@echo off
echo ========================================
echo      AutoWord GUI 启动器
echo ========================================
echo.

echo 正在启动AutoWord图形界面...
echo.

REM 检查Python是否可用
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请确保Python已安装并添加到PATH
    echo.
    pause
    exit /b 1
)

REM 启动GUI应用程序
python main_gui.py

REM 如果程序异常退出，显示错误信息
if errorlevel 1 (
    echo.
    echo ========================================
    echo GUI启动失败，请检查上面的错误信息
    echo ========================================
    echo.
    echo 常见解决方案:
    echo 1. 运行: pip install -r requirements.txt
    echo 2. 确保Microsoft Word已安装
    echo 3. 检查Python版本是否为3.8+
    echo.
    pause
) else (
    echo.
    echo GUI已正常关闭
)