#!/usr/bin/env python3
"""
AutoWord 构建脚本
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path
import argparse


def run_command(cmd, cwd=None):
    """运行命令"""
    print(f"运行命令: {cmd}")
    result = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"命令执行失败: {result.stderr}")
        sys.exit(1)
    
    return result.stdout


def clean_build():
    """清理构建目录"""
    print("清理构建目录...")
    
    dirs_to_clean = ['build', 'dist', '*.egg-info']
    
    for pattern in dirs_to_clean:
        for path in Path('.').glob(pattern):
            if path.is_dir():
                print(f"删除目录: {path}")
                shutil.rmtree(path)
            else:
                print(f"删除文件: {path}")
                path.unlink()


def install_dependencies():
    """安装依赖"""
    print("安装依赖...")
    
    # 安装基础依赖
    run_command("pip install -r requirements.txt")
    
    # 安装开发依赖
    run_command("pip install -e .[dev]")


def run_tests():
    """运行测试"""
    print("运行测试...")
    
    # 运行测试
    run_command("python -m pytest tests/ -v --tb=short")
    
    # 生成覆盖率报告
    run_command("python -m pytest tests/ --cov=autoword --cov-report=html --cov-report=term")


def build_package():
    """构建 Python 包"""
    print("构建 Python 包...")
    
    # 构建源码包和轮子包
    run_command("python setup.py sdist bdist_wheel")


def build_executable():
    """构建可执行文件"""
    print("构建可执行文件...")
    
    # 安装 PyInstaller
    run_command("pip install pyinstaller")
    
    # 使用 spec 文件构建
    run_command("pyinstaller build.spec --clean")


def create_installer():
    """创建安装程序"""
    print("创建安装程序...")
    
    # 这里可以集成 NSIS 或其他安装程序创建工具
    # 暂时只是复制文件到发布目录
    
    release_dir = Path("release")
    release_dir.mkdir(exist_ok=True)
    
    # 复制可执行文件
    if Path("dist/AutoWord").exists():
        shutil.copytree("dist/AutoWord", release_dir / "AutoWord", dirs_exist_ok=True)
    
    # 复制文档
    docs_to_copy = ["README.md", "LICENSE", "CHANGELOG.md"]
    for doc in docs_to_copy:
        if Path(doc).exists():
            shutil.copy2(doc, release_dir)
    
    print(f"发布文件已创建在: {release_dir}")


def validate_environment():
    """验证构建环境"""
    print("验证构建环境...")
    
    # 检查 Python 版本
    if sys.version_info < (3, 10):
        print("错误: 需要 Python 3.10 或更高版本")
        sys.exit(1)
    
    # 检查 Windows 环境
    if os.name != 'nt':
        print("警告: AutoWord 主要为 Windows 设计")
    
    # 检查 Word COM 可用性
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        print("✓ Word COM 可用")
    except Exception as e:
        print(f"警告: Word COM 不可用: {e}")
    
    print("环境验证完成")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="AutoWord 构建脚本")
    parser.add_argument("--clean", action="store_true", help="清理构建目录")
    parser.add_argument("--deps", action="store_true", help="安装依赖")
    parser.add_argument("--test", action="store_true", help="运行测试")
    parser.add_argument("--package", action="store_true", help="构建 Python 包")
    parser.add_argument("--exe", action="store_true", help="构建可执行文件")
    parser.add_argument("--installer", action="store_true", help="创建安装程序")
    parser.add_argument("--all", action="store_true", help="执行完整构建流程")
    parser.add_argument("--validate", action="store_true", help="验证环境")
    
    args = parser.parse_args()
    
    # 如果没有指定参数，显示帮助
    if not any(vars(args).values()):
        parser.print_help()
        return
    
    try:
        if args.validate or args.all:
            validate_environment()
        
        if args.clean or args.all:
            clean_build()
        
        if args.deps or args.all:
            install_dependencies()
        
        if args.test or args.all:
            run_tests()
        
        if args.package or args.all:
            build_package()
        
        if args.exe or args.all:
            build_executable()
        
        if args.installer or args.all:
            create_installer()
        
        print("构建完成!")
        
    except KeyboardInterrupt:
        print("\n构建被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"构建失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()