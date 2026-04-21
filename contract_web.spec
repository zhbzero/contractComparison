# -*- mode: python ; coding: utf-8 -*-
# 在 Windows 上执行：pyinstaller contract_web.spec
# 生成单文件 exe：dist\文档比对Web.exe

from PyInstaller.utils.hooks import collect_all

lxml_datas, lxml_binaries, lxml_hiddenimports = collect_all("lxml")

a = Analysis(
    ["web_app.py"],
    pathex=[],
    binaries=lxml_binaries,
    datas=lxml_datas + [("templates", "templates")],
    hiddenimports=lxml_hiddenimports + ["lxml.etree", "lxml._elementpath"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="文档比对Web",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
