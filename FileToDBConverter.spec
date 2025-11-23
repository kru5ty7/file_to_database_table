# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src/app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'customtkinter',
        'darkdetect',
        'pandas',
        'pyodbc',
        'openpyxl',
        'cryptography',
        'cryptography.fernet',
        'PIL',
        'PIL.ImageGrab',
        'src',
        'src.database',
        'src.file_processor',
        'src.utils',
        'src.gui_main',
        'src.dialogs',
        'src.dialogs.preview_dialog',
        'src.dialogs.connection_dialog',
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

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='FileToDBConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for windowed app (no console)
    disable_windowing_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon='app.ico' if you have an icon file
)
