# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('version.json', '.'), ('C:/Users/QC60/AppData/Local/Programs/Python/Python314/Lib/site-packages/customtkinter', 'customtkinter'), ('E:/CODE/compare MATERIAL/VUNG Folder/VUNG Folder/public', 'public')],
    hiddenimports=['pymupdf', 'fitz', 'openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'et_xmlfile', 'PIL', 'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox', 'customtkinter', 'windnd', 'bom_comparator', 'bom_extractor', 'excel_exporter', 'model_comparator', 'series_bom_comparator', 'series_bom_exporter', 'updater'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='VIPQC AI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['E:/CODE/compare MATERIAL/VUNG Folder/VUNG Folder/public/Logo.ico'],
)
