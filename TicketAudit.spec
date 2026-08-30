# -*- mode: python ; coding: utf-8 -*-

# PyInstaller 6 spec. Do not reintroduce `block_cipher` / `cipher=` /
# `a.zipped_data` / `win_no_prefer_redirects` / `win_private_assemblies`:
# bytecode encryption was removed in PyInstaller 6.0 and those arguments now
# raise TypeError, which fails the build before it starts.

datas = [
    ('gui/assets', 'gui/assets'),
    ('core/assets', 'core/assets'),
    ('gui/app_pyside.py', 'gui'),
    ('banner.txt', '.'),
]

hiddenimports = [
    'PySide6',
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    'PySide6.QtCharts',
    'PySide6.QtSvg',
    'pandas',
    'numpy',
    'openpyxl',
    'xlrd',
    'lingua',
    'pyperclip',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'unittest'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TicketAudit',
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
    icon='gui/assets/ticketaudit.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TicketAudit',
)
