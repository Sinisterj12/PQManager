# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['bootstrap.py'],  # Use bootstrap instead of main.py
    pathex=['d:/!Projects/PQManager'],
    binaries=[],
    datas=[('assets/Logo.ico', 'assets')],
    hiddenimports=[
        'win32print',
        'win32api',
        'win32con',
        'win32timezone',  # Add this for timezone support
        'win32security',  # Additional win32 modules that might be needed
        'win32ts',
        'tkinter',
        'tkinter.ttk',
        'PIL._tkinter_finder',
        'pystray._win32',
        'email.mime.text',
        'email.mime.multipart',
        'email.mime.nonmultipart',
        'email.mime.base',
        'email.charset',
        'email.encoders',
        'email.errors',
        'email.utils',
        'email.header'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Add source files
a.datas += [('src/main.py', 'd:/!Projects/PQManager/src/main.py', 'DATA')]
a.datas += [('src/printer_manager.py', 'd:/!Projects/PQManager/src/printer_manager.py', 'DATA')]
a.datas += [('src/tray_manager.py', 'd:/!Projects/PQManager/src/tray_manager.py', 'DATA')]
a.datas += [('src/config_manager.py', 'd:/!Projects/PQManager/src/config_manager.py', 'DATA')]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PQManager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False to hide console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/Logo.ico',
    version='file_version_info.txt',
    uac_admin=True,
)
