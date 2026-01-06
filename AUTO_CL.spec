# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# ✅ Mais estável do que Path.cwd()
PROJECT_ROOT = Path(__spec__.origin).resolve().parent

APP_NAME = "pacl_v_1_2_1"
ICON_PATH = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\python\auto_cl\frontend\media\BR.ico"

# -----------------------------
# Hidden imports
# -----------------------------
hiddenimports = []
hiddenimports += ["pythoncom", "pywintypes"]
hiddenimports += collect_submodules("win32com")

# ✅ Como você roda scripts via runpy, force módulos usados lá:
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("pandas")

# Seus pacotes
hiddenimports += collect_submodules("backend")
hiddenimports += collect_submodules("app")

# Eel (ok)
hiddenimports += collect_submodules("eel")

# -----------------------------
# Datas
# -----------------------------
datas = []

# ✅ openpyxl tem templates internos (recomendado)
datas += collect_data_files("openpyxl")

# user_data.csv dentro de app/services
csv_path = PROJECT_ROOT / "app" / "services" / "user_data.csv"
if csv_path.exists():
    datas.append((str(csv_path), "app/services"))

# frontend inteiro
frontend_dir = PROJECT_ROOT / "frontend"
if frontend_dir.exists():
    datas.append((str(frontend_dir), "frontend"))

# backend inteiro como arquivos (opcional, mas ajuda se você usa paths)
backend_dir = PROJECT_ROOT / "backend"
if backend_dir.exists():
    datas.append((str(backend_dir), "backend"))

a = Analysis(
    ["app\\main_app.py"],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    exclude_binaries=True,   # ✅ onedir
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=ICON_PATH,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name=APP_NAME,
)
