# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['run_flask.py'],
    pathex=[],
    binaries=[],
    datas=[('quote_system/templates', 'quote_system/templates'), ('quote_system/static', 'quote_system/static'), ('config', 'config'), ('模板', '模板')],
    hiddenimports=['quote_system.web_app', 'quote_system.generator', 'quote_system.memoq_html', 'quote_system.projects', 'quote_system.config', 'quote_system.save_path_config', 'quote_system.paths', 'quote_system.utils', 'quote_system.feishu_client', 'quote_system.auto_quote', 'quote_system.auto_quote_zhan_shuang', 'quote_system.auto_quote_zhan_shuang_feishu', 'quote_system.auto_quote_4399', 'quote_system.auto_fill_tk', 'settlement.generate_settlement', 'settlement.generate_settlement_bilibili', 'settlement.generate_settlement_zhan_shuang', 'settlement.generate_settlement_zhan_shuang_faxing', 'settlement.generate_settlement_4399', 'settlement.generate_settlement_diezhi', 'settlement.settlement_tracker', 'settlement._com_utils', 'lxml', 'openpyxl', 'docx', 'win32com.client', 'pythoncom', 'pywintypes', 'win32api'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'PyQt5', 'IPython', 'pytest'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='flask_server',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='flask_server',
)
