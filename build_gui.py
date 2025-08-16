#!/usr/bin/env python3
"""
AutoWord GUI 构建脚本
使用 PyInstaller 打包 GUI 应用程序
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path


def check_dependencies():
    """检查构建依赖"""
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__}")
    except ImportError:
        print("✗ PyInstaller 未安装")
        print("请运行: pip install pyinstaller")
        return False
    
    # 检查其他依赖
    required_packages = [
        'PySide6', 'qframelesswindow', 'qdarktheme', 
        'cryptography', 'pywin32'
    ]
    
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"✓ {package}")
        except ImportError:
            missing_packages.append(package)
            print(f"✗ {package}")
    
    if missing_packages:
        print(f"\n缺少依赖包: {', '.join(missing_packages)}")
        print("请运行: pip install -r requirements.txt")
        return False
    
    return True


def clean_build_dirs():
    """清理构建目录"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}")
            shutil.rmtree(dir_name)
    
    # 清理 .pyc 文件
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.pyc'):
                os.remove(os.path.join(root, file))


def build_application():
    """构建应用程序"""
    print("开始构建 AutoWord GUI...")
    
    # 构建命令
    cmd = [
        'pyinstaller',
        '--clean',
        '--noconfirm',
        'build_gui.spec'
    ]
    
    try:
        # 执行构建
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("构建成功!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"构建失败: {e}")
        print("错误输出:")
        print(e.stderr)
        return False


def create_installer_script():
    """创建安装脚本"""
    installer_script = """@echo off
echo AutoWord GUI Installer
echo =====================
echo.

REM 检查管理员权限
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
) else (
    echo This installer requires administrator privileges.
    echo Please run as administrator.
    pause
    exit /b 1
)

REM 创建程序目录
set INSTALL_DIR=C:\\Program Files\\AutoWord
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM 复制文件
echo Installing AutoWord GUI...
xcopy /E /I /Y "AutoWord\\*" "%INSTALL_DIR%\\"

REM 创建桌面快捷方式
set DESKTOP=%USERPROFILE%\\Desktop
echo Creating desktop shortcut...
powershell "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%DESKTOP%\\AutoWord.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\\AutoWord.exe'; $Shortcut.Save()"

REM 创建开始菜单快捷方式
set START_MENU=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs
if not exist "%START_MENU%\\AutoWord" mkdir "%START_MENU%\\AutoWord"
powershell "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%START_MENU%\\AutoWord\\AutoWord.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\\AutoWord.exe'; $Shortcut.Save()"

echo.
echo Installation completed successfully!
echo You can now run AutoWord from:
echo - Desktop shortcut
echo - Start Menu
echo - %INSTALL_DIR%\\AutoWord.exe
echo.
pause
"""
    
    with open('dist/install.bat', 'w', encoding='utf-8') as f:
        f.write(installer_script)
    
    print("安装脚本已创建: dist/install.bat")


def create_readme():
    """创建发布说明"""
    readme_content = """# AutoWord GUI - 发布版本

## 安装说明

### 方法1: 直接运行
1. 解压所有文件到任意目录
2. 双击 `AutoWord.exe` 运行程序

### 方法2: 系统安装
1. 以管理员身份运行 `install.bat`
2. 程序将安装到 `C:\\Program Files\\AutoWord`
3. 自动创建桌面和开始菜单快捷方式

## 系统要求

- Windows 10/11 (64位)
- Microsoft Word (已安装)
- 网络连接 (用于API调用)

## 使用说明

1. 启动程序
2. 选择AI模型 (Claude 3.7 或 GPT-4o)
3. 输入对应的API密钥
4. 选择要处理的Word文档
5. 点击"开始处理"

## 获取API密钥

- **Claude 3.7**: 访问 https://console.anthropic.com/
- **GPT-4o**: 访问 https://platform.openai.com/

## 技术支持

如遇问题，请查看程序内的错误提示和日志信息。

## 版本信息

- 版本: 1.0.0
- 构建日期: {build_date}
- Python版本: {python_version}
""".format(
        build_date=__import__('datetime').datetime.now().strftime('%Y-%m-%d'),
        python_version=sys.version.split()[0]
    )
    
    with open('dist/README.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print("发布说明已创建: dist/README.txt")


def main():
    """主函数"""
    print("AutoWord GUI 构建工具")
    print("=" * 30)
    
    # 检查依赖
    print("检查构建依赖...")
    if not check_dependencies():
        return 1
    
    # 清理构建目录
    print("\n清理构建目录...")
    clean_build_dirs()
    
    # 构建应用程序
    print("\n构建应用程序...")
    if not build_application():
        return 1
    
    # 创建附加文件
    print("\n创建附加文件...")
    create_installer_script()
    create_readme()
    
    print("\n构建完成!")
    print("输出目录: dist/AutoWord/")
    print("可执行文件: dist/AutoWord/AutoWord.exe")
    print("安装脚本: dist/install.bat")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())