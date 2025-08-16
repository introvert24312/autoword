# -*- mode: python ; coding: utf-8 -*-
"""
AutoWord PyInstaller 构建配置
"""

import os
import sys
from pathlib import Path

# 项目根目录
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# 数据文件
datas = [
    (str(project_root / 'schemas' / '*.json'), 'schemas'),
    (str(project_root / 'autoword' / 'resources'), 'autoword/resources'),
]

# 隐藏导入
hiddenimports = [
    'win32timezone',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    'win32api',
    'win32con',
    'win32gui',
    'win32process',
    'orjson',
    'pydantic',
    'requests',
    'colorlog',
    'click',
]

# GUI 相关的隐藏导入（如果构建 GUI 版本）
gui_hiddenimports = [
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    'qfluentwidgets',
    'qframelesswindow',
]

# 排除的模块
excludes = [
    'tkinter',
    'matplotlib',
    'numpy',
    'scipy',
    'pandas',
    'jupyter',
    'IPython',
]

# 控制台版本
console_a = Analysis(
    ['autoword/cli/main.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

console_pyz = PYZ(console_a.pure, console_a.zipped_data, cipher=None)

console_exe = EXE(
    console_pyz,
    console_a.scripts,
    [],
    exclude_binaries=True,
    name='autoword',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(project_root / 'autoword' / 'resources' / 'icon.ico') if (project_root / 'autoword' / 'resources' / 'icon.ico').exists() else None,
)

# GUI 版本（如果存在 GUI 模块）
gui_exists = (project_root / 'autoword' / 'gui' / 'main.py').exists()

if gui_exists:
    gui_a = Analysis(
        ['autoword/gui/main.py'],
        pathex=[str(project_root)],
        binaries=[],
        datas=datas + [
            (str(project_root / 'autoword' / 'gui' / 'resources'), 'autoword/gui/resources'),
        ],
        hiddenimports=hiddenimports + gui_hiddenimports,
        hookspath=[],
        hooksconfig={},
        runtime_hooks=[],
        excludes=excludes,
        win_no_prefer_redirects=False,
        win_private_assemblies=False,
        cipher=None,
        noarchive=False,
    )

    gui_pyz = PYZ(gui_a.pure, gui_a.zipped_data, cipher=None)

    gui_exe = EXE(
        gui_pyz,
        gui_a.scripts,
        [],
        exclude_binaries=True,
        name='autoword-gui',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=str(project_root / 'autoword' / 'resources' / 'icon.ico') if (project_root / 'autoword' / 'resources' / 'icon.ico').exists() else None,
    )

    # 收集所有文件
    coll = COLLECT(
        console_exe,
        console_a.binaries,
        console_a.zipfiles,
        console_a.datas,
        gui_exe,
        gui_a.binaries,
        gui_a.zipfiles,
        gui_a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='AutoWord',
    )
else:
    # 只有控制台版本
    coll = COLLECT(
        console_exe,
        console_a.binaries,
        console_a.zipfiles,
        console_a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='AutoWord',
    )