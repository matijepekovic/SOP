from pathlib import Path


project_root = Path(SPECPATH).parent
icon_path = project_root / "assets" / "tray_icon.ico"

hiddenimports = [
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    "win32timezone",
    "keyring.backends.Windows",
    "pystray._win32",
    "PIL._tkinter_finder",
    "tkinter",
    "tkinter.messagebox",
    "tkinter.ttk",
    # Imported lazily inside TrayApp.run(), so static analysis can miss it.
    "sop_reporter.gui.control_window",
]

datas = [
    (str(project_root / "config" / "app_config.default.yaml"), "config"),
    (str(project_root / "config" / "extraction_rules.default.yaml"), "config"),
    (str(icon_path), "assets"),
]

a = Analysis(
    [str(project_root / "sop_reporter" / "main.py")],
    pathex=[str(project_root)],
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
    a.binaries,
    a.datas,
    [],
    name="SOPReporter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_path),
)
