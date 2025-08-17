@echo off
chcp 65001 >nul
title AutoWord vNext - 一键启动
color 0A

echo.
echo     ╔══════════════════════════════════════╗
echo     ║        AutoWord vNext 一键启动        ║
echo     ║      智能文档处理系统 v1.0.0          ║
echo     ╚══════════════════════════════════════╝
echo.

echo 🚀 欢迎使用AutoWord vNext！
echo.

REM 检查是否首次运行
if not exist "vnext_config.json" (
    echo 🔧 检测到首次运行，启动配置向导...
    echo.
    python 快速配置.py
    echo.
    if errorlevel 1 (
        echo ❌ 配置失败，请手动运行配置向导
        pause
        exit /b 1
    )
)

echo 📋 选择启动方式:
echo.
echo 1. 🖥️  启动简化版GUI (推荐)
echo 2. 💻 启动命令行版本
echo 3. 🧪 运行性能测试
echo 4. ⚙️  重新配置
echo 5. 📚 查看文档
echo 6. ❌ 退出
echo.

set /p choice="请选择 (1-6): "

if "%choice%"=="1" goto start_gui
if "%choice%"=="2" goto start_cli
if "%choice%"=="3" goto run_test
if "%choice%"=="4" goto reconfig
if "%choice%"=="5" goto show_docs
if "%choice%"=="6" goto exit
goto invalid_choice

:start_gui
echo.
echo 🖥️  启动简化版GUI界面...
python "简化版GUI.py"
if errorlevel 1 (
    echo.
    echo ❌ GUI启动失败，尝试命令行版本...
    echo.
    python "命令行版本.py"
)
goto end

:start_cli
echo.
echo 💻 启动命令行版本...
python "命令行版本.py"
goto end

:run_test
echo.
echo 🧪 运行性能测试...
python "性能测试工具.py"
goto end

:reconfig
echo.
echo ⚙️  重新配置...
python "快速配置.py"
goto end

:show_docs
echo.
echo 📚 打开文档目录...
if exist "docs\README.md" (
    start "" "docs"
    echo ✅ 文档目录已打开
) else (
    echo ❌ 文档目录不存在
)
goto end

:invalid_choice
echo.
echo ❌ 无效选择，请输入1-6
echo.
pause
goto start

:exit
echo.
echo 👋 感谢使用AutoWord vNext！
goto end

:end
echo.
pause