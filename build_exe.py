"""
Build script: Compile BOM Extractor into a standalone Windows .exe using PyInstaller.
Includes customtkinter hidden imports.
"""

import os
import sys
import subprocess
import shutil

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def build():
    print("=" * 60)
    print("STARTING PYINSTALLER BUILD: VIPQC AI.exe")
    print("=" * 60)

    # Clean previous build artifacts
    for d in ["build", "dist"]:
        if os.path.exists(d):
            print(f"Cleaning previous {d}/ directory...")
            shutil.rmtree(d, ignore_errors=True)

    # Locate customtkinter data directory for bundling
    try:
        import customtkinter
        ctk_path = os.path.dirname(customtkinter.__file__)
        ctk_data = f"{ctk_path}{os.pathsep}customtkinter"
    except ImportError:
        ctk_data = None

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", "VIPQC AI",
        "--hidden-import", "pymupdf",
        "--hidden-import", "fitz",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.styles",
        "--hidden-import", "openpyxl.utils",
        "--hidden-import", "et_xmlfile",
        "--hidden-import", "PIL",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--hidden-import", "customtkinter",
        "--hidden-import", "windnd",
        "--hidden-import", "bom_comparator",
        "--hidden-import", "bom_extractor",
        "--hidden-import", "excel_exporter",
        "--hidden-import", "model_comparator",
        "--hidden-import", "series_bom_comparator",
        "--hidden-import", "series_bom_exporter",
        "--hidden-import", "updater",
    ]

    if os.path.exists("version.json"):
        cmd.extend(["--add-data", f"version.json{os.pathsep}."])

    if ctk_data:
        cmd += ["--add-data", ctk_data]

    # Application Icon
    ico_path = os.path.abspath(os.path.join("public", "Logo.ico"))
    if os.path.exists(ico_path):
        cmd.extend(["--icon", ico_path])

    # Bundled Assets (public folder with Logo.png & Logo.ico)
    public_dir = os.path.abspath("public")
    if os.path.exists(public_dir):
        cmd.extend(["--add-data", f"{public_dir}{os.pathsep}public"])

    cmd += ["app.py"]

    print("Running command:", " ".join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True)

    print("\n--- Build Output (last 3000 chars) ---")
    print(proc.stdout[-3000:] if len(proc.stdout) > 3000 else proc.stdout)

    if proc.returncode != 0:
        print("\n--- Build Errors ---")
        print(proc.stderr[-2000:])
        sys.exit(1)

    exe_path = os.path.join("dist", "VIPQC AI.exe")
    if os.path.exists(exe_path):
        target_path = "VIPQC AI.exe"
        shutil.copy2(exe_path, target_path)
        shutil.rmtree("dist", ignore_errors=True)
        shutil.rmtree("build", ignore_errors=True)
        size_mb = os.path.getsize(target_path) / (1024 * 1024)
        # Notify Windows Shell to refresh icon cache
        try:
            import ctypes
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        except Exception:
            pass
        print("=" * 60)
        print(f"BUILD SUCCESSFUL!  ->  {target_path}  ({size_mb:.1f} MB)")
        print("=" * 60)
    else:
        print("ERROR: dist/VIPQC AI.exe not found!")
        sys.exit(1)


if __name__ == "__main__":
    build()
