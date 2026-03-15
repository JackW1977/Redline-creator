# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for MS Word Redline Creator v1.0.

Produces a single-folder distribution under dist/RedlineCreator/
containing RedlineCreator.exe and all dependencies.
"""

import sys
from pathlib import Path

block_cipher = None

# All Python source modules that make up the application
source_modules = [
    'gui.py',
    'compare_revisions.py',
    'comment_extractor.py',
    'comment_inserter.py',
    'comment_mapper.py',
    'config.py',
    'font_preserver.py',
    'text_extractor.py',
    'word_compare.py',
]

# Collect data files needed at runtime
datas = []

# Hidden imports that PyInstaller may miss
hidden_imports = [
    'lxml',
    'lxml.etree',
    'lxml._elementpath',
]

# Conditionally add Windows-only packages
if sys.platform == 'win32':
    hidden_imports += [
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'win32api',
    ]

# Try to include tkinterdnd2 if available
try:
    import tkinterdnd2
    tkdnd_path = Path(tkinterdnd2.__file__).parent
    # Include the tkdnd shared libraries
    datas.append((str(tkdnd_path), 'tkinterdnd2'))
    hidden_imports.append('tkinterdnd2')
except ImportError:
    pass

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'PIL',
        'pytest', 'unittest', 'doctest',
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
    [],
    exclude_binaries=True,
    name='RedlineCreator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # No console window — GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',       # Uncomment if you add an icon file
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='RedlineCreator',
)
