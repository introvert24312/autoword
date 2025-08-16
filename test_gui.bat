@echo off
chcp 65001 >nul
echo ========================================
echo      AutoWord GUI 测试工具
echo ========================================
echo.

echo [1/4] 检查Python环境...
python --version
if errorlevel 1 (
    echo ❌ Python未安装或未添加到PATH
    goto :error
)
echo ✅ Python环境正常
echo.

echo [2/4] 检查GUI依赖包...
python -c "import PySide6; print('✅ PySide6:', PySide6.__version__)" 2>nul || (
    echo ❌ PySide6未安装
    echo 正在安装PySide6...
    pip install PySide6
)

python -c "import qdarkstyle; print('✅ qdarkstyle:', qdarkstyle.__version__)" 2>nul || (
    echo ❌ qdarkstyle未安装
    echo 正在安装qdarkstyle...
    pip install qdarkstyle
)

python -c "import cryptography; print('✅ cryptography:', cryptography.__version__)" 2>nul || (
    echo ❌ cryptography未安装
    echo 正在安装cryptography...
    pip install cryptography
)
echo.

echo [3/4] 检查Word环境...
python -c "import win32com.client; win32com.client.Dispatch('Word.Application').Quit(); print('✅ Microsoft Word可用')" 2>nul || (
    echo ⚠️  Microsoft Word未安装或不可用
    echo 注意: 这可能影响文档处理功能
)
echo.

echo [4/4] 启动GUI应用程序...
echo.
echo ========================================
echo      正在启动AutoWord GUI...
echo ========================================
echo.

python main_gui.py

echo.
if errorlevel 1 (
    goto :error
) else (
    echo ========================================
    echo      GUI测试完成
    echo ========================================
    goto :end
)

:error
echo.
echo ========================================
echo      发生错误
echo ========================================
echo.
echo 请检查上面的错误信息并尝试以下解决方案:
echo.
echo 1. 安装所有依赖:
echo    pip install -r requirements.txt
echo.
echo 2. 确保Microsoft Word已安装
echo.
echo 3. 检查Python版本 (需要3.8+):
echo    python --version
echo.
echo 4. 如果仍有问题，请查看详细日志
echo.

:end
pause