# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 빌드 설정. `빌드.bat` 또는 `pyinstaller 조회처추출.spec` 로 실행한다.

config와 template 폴더는 exe 안에 넣고, exe 옆에도 함께 복사한다(dist 폴더).
exe 옆에 config가 있으면 그쪽을 먼저 쓴다 — 기관 사전을 고칠 수 있게 하기 위해서다.
"""

a = Analysis(
    ["src/app.py"],
    pathex=["src"],
    binaries=[],
    datas=[("config", "config"), ("template", "template")],
    hiddenimports=["win32com.client", "win32timezone"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["PIL", "numpy", "pandas", "matplotlib", "pytest", "setuptools"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="조회처추출",
    debug=False,
    strip=False,
    upx=False,
    console=False,          # 화면 실행 — 검은 창을 띄우지 않는다
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
