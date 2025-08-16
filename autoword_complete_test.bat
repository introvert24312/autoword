@echo off
chcp 65001 >nul
echo ========================================
echo      AutoWord 完整测试套件
echo ========================================
echo.

echo 请选择要执行的操作:
echo.
echo 1. 测试API配置
echo 2. 启动GUI应用程序
echo 3. 完整测试 (API + GUI)
echo 4. 退出
echo.

set /p choice="请输入选择 (1-4): "

if "%choice%"=="1" goto test_api
if "%choice%"=="2" goto start_gui
if "%choice%"=="3" goto full_test
if "%choice%"=="4" goto exit
goto invalid_choice

:test_api
echo.
echo ========================================
echo      测试API配置
echo ========================================
echo.
python test_api_config.py
goto end

:start_gui
echo.
echo ========================================
echo      启动GUI应用程序
echo ========================================
echo.
python main_gui.py
goto end

:full_test
echo.
echo ========================================
echo      执行完整测试
echo ========================================
echo.
echo [1/2] 测试API配置...
python test_api_config.py
if errorlevel 1 (
    echo.
    echo ❌ API测试失败，请检查配置后再启动GUI
    goto end
)

echo.
echo [2/2] 启动GUI应用程序...
echo.
python main_gui.py
goto end

:invalid_choice
echo.
echo ❌ 无效选择，请输入1-4之间的数字
echo.
pause
goto :eof

:exit
echo.
echo 退出测试套件
goto :eof

:end
echo.
echo ========================================
echo      测试完成
echo ========================================
echo.
pause