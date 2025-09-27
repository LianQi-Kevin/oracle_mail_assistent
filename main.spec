# -*- mode: python ; coding: utf-8 -*-

import sys, pathlib
from PyInstaller.utils.hooks import collect_submodules  # 如需隐藏导入可用
# ------------------ ① 额外 DLL 搜索 ------------------
def find_extra_dlls():
    root = pathlib.Path(sys.base_prefix)   # venv / conda 根
    dll_candidates = [
        root / "Library" / "bin" / "libssl-3-x64.dll",
        root / "Library" / "bin" / "libcrypto-3-x64.dll",
        root / "Library" / "bin" / "libssl-1_1-x64.dll",
        root / "Library" / "bin" / "libcrypto-1_1-x64.dll",
        root / "Library" / "bin" / "libexpat.dll",
        root / "Library" / "bin" / "liblzma.dll",
        root / "Library" / "bin" / "LIBBZ2.dll",
        root / "DLLs" / "libssl-3-x64.dll",
        root / "DLLs" / "libcrypto-3-x64.dll",
        root / "DLLs" / "libssl-1_1-x64.dll",
        root / "DLLs" / "libcrypto-1_1-x64.dll",
        root / "DLLs" / "libexpat.dll",
        root / "DLLs" / "liblzma.dll",
        root / "DLLs" / "LIBBZ2.dll",
    ]
    # 只返回存在的文件，目标目录设为 '.'（dist 根）
    return [(str(p), '.') for p in dll_candidates if p.exists()]

extra_binaries = find_extra_dlls()

# ------------------ ② 生成分析对象 ------------------
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=extra_binaries,   # <-- 二元组列表 OK
    datas=[],
    hiddenimports=[],
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
    name='aconex_drawing_progress_updater.exe',
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
