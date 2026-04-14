# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_all


runtime_datas = []
runtime_binaries = []
runtime_hiddenimports = []
project_root = Path.cwd()
app_datas = [
    ('assets', 'assets'),
    ('workspace/data', 'workspace/data'),
    ('certs', 'certs'),
]

if (project_root / '.env').exists():
    app_datas.append(('.env', '.'))
elif (project_root / '.env.example').exists():
    app_datas.append(('.env.example', '.'))

for package_name in ('psycopg', 'dotenv', 'openpyxl', 'reportlab', 'PIL'):
    package_datas, package_binaries, package_hiddenimports = collect_all(package_name)
    runtime_datas += package_datas
    runtime_binaries += package_binaries
    runtime_hiddenimports += package_hiddenimports

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=runtime_binaries,
    datas=[*app_datas, *runtime_datas],
    hiddenimports=runtime_hiddenimports,
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
    [],
    exclude_binaries=True,
    name='Liquimech ERP',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Liquimech ERP',
)
app = BUNDLE(
    coll,
    name='Liquimech ERP.app',
    icon=None,
    bundle_identifier=None,
)
