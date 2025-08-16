@echo off
chcp 65001 >nul
echo ========================================
echo      AutoWord API 配置测试
echo ========================================
echo.

echo 正在测试API配置...
echo.

python test_api_config.py

echo.
if errorlevel 1 (
    echo ========================================
    echo      测试失败
    echo ========================================
    echo.
    echo 请检查:
    echo 1. API密钥是否正确
    echo 2. 网络连接是否正常
    echo 3. API服务是否可用
    echo.
) else (
    echo ========================================
    echo      测试成功
    echo ========================================
    echo.
    echo API配置正确，可以正常使用GUI！
    echo.
)

pause