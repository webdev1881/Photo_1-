# -*- mode: python ; coding: utf-8 -*-

import sys
import os

# Определяем дополнительные модули для включения
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog', 
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'requests',
    'urllib3',
    'pandas',
    'openpyxl',
    'bs4',
    'soupsieve',
    're',
    'json',
    'threading',
    'time',
    'random',
    'os',
    'sys'
]

# Добавляем PIL и OpenCV если доступны
try:
    import PIL
    hiddenimports.extend([
        'PIL',
        'PIL.Image',
        'PIL.ImageTk', 
        'PIL.ImageOps',
        'PIL.ImageDraw'
    ])
except ImportError:
    pass

try:
    import cv2
    import numpy
    hiddenimports.extend([
        'cv2',
        'numpy'
    ])
except ImportError:
    pass

# Файлы для включения в сборку
datas = []
if os.path.exists('icon.ico'):
    datas.append(('icon.ico', '.'))
if os.path.exists('target.xlsx'):
    datas.append(('target.xlsx', '.'))

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'PyQt5',
        'PyQt6', 
        'PySide2',
        'PySide6',
        'wx'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ImageProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Отключает консольное окно
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None
)