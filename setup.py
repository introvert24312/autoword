"""
AutoWord Setup Script
AutoWord 安装脚本
"""

from setuptools import setup, find_packages
import os

# 读取 README 文件
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

# 读取版本信息
version = "1.0.0"

# 读取依赖
with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="autoword",
    version=version,
    author="AutoWord Team",
    author_email="support@autoword.com",
    description="AutoWord - 智能 Word 文档自动化工具",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/autoword/autoword",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Text Processing :: Markup",
    ],
    python_requires=">=3.10",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=22.0.0",
            "flake8>=5.0.0",
            "mypy>=1.0.0",
        ],
        "gui": [
            "PySide6>=6.5.0",
            "qfluentwidgets>=1.4.0",
            "qframelesswindow>=0.3.0",
        ]
    },
    entry_points={
        "console_scripts": [
            "autoword=autoword.cli.main:main",
        ],
        "gui_scripts": [
            "autoword-gui=autoword.gui.main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "autoword": [
            "schemas/*.json",
            "resources/*",
            "gui/resources/*",
        ],
    },
    zip_safe=False,
)