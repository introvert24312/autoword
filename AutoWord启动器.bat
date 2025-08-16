@echo off
chcp 65001 >nul
title AutoWord 文档自动化工具

cls
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                    AutoWord 文档自动化工具                    ║
echo ║                                                              ║
echo ║  • 支持 Claude 3.7 Sonnet 和 GPT-5 模型                     ║
echo ║  • API密钥已预配置，开箱即用                                  ║
echo ║  • 智能文档处理，批注驱动的自动化                             ║
echo ║  • 智能输出文件命名 (.process.docx)                          ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

echo [启动中] 正在初始化AutoWord GUI...
echo.

REM 启动GUI应用程序
python main_gui.py

REM 检查退出状态
if errorlevel 1 (
    echo.
    echo ╔══════════════════════════════════════════════════════════════╗
    echo ║                        启动失败                              ║
    echo ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo 可能的解决方案:
    echo.
    echo 1. 确保Python已安装 (推荐Python 3.8+)
    echo    检查命令: python --version
    echo.
    echo 2. 安装必要的依赖包:
    echo    pip install -r requirements.txt
    echo.
    echo 3. 确保Microsoft Word已安装
    echo.
    echo 4. 检查网络连接 (API调用需要网络)
    echo.
    echo 按任意键退出...
    pause >nul
) else (
    echo.
    echo [完成] AutoWord已正常关闭
)