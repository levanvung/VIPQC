import sys
sys.stdout.reconfigure(encoding="utf-8")
import shutil
import py_compile

# 1. Backup app.py
shutil.copyfile("app.py", "app.py.bak")
print("Backed up app.py to app.py.bak")

# 2. Copy scratch/test_updated_app.py to app.py
shutil.copyfile("scratch/test_updated_app.py", "app.py")
print("Replaced app.py with updated version.")

# 3. Verify compilation
try:
    py_compile.compile("app.py", doraise=True)
    print("Verification SUCCESS: app.py compiles cleanly!")
except Exception as e:
    print(f"Compilation ERROR: {e}")
    shutil.copyfile("app.py.bak", "app.py")
    sys.exit(1)
