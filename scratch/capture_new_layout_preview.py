import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
    if os.path.exists(p):
        try: os.add_dll_directory(p)
        except Exception: pass

ws_path = r"e:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder"
sys.path.insert(0, ws_path)
os.chdir(ws_path)

from PIL import ImageGrab
from app import BOMExtractorApp

app = BOMExtractorApp()
app.geometry("1400x900+50+50")
app._set_nav_active("model_comp")
app.update()

mc = app.view_model_compare
pdf_a = os.path.join(ws_path, "sample_manuals", "media_1788593034521.pdf")
pdf_b = os.path.join(ws_path, "sample_manuals", "media_1788593034536.pdf")

if os.path.exists(pdf_a) and os.path.exists(pdf_b):
    mc.pdf_a_path = pdf_a
    mc.pdf_b_path = pdf_b
    mc._run_comparison()
    app.update()
    time.sleep(0.5)

    # Capture main window screenshot
    x = app.winfo_rootx()
    y = app.winfo_rooty()
    w = app.winfo_width()
    h = app.winfo_height()
    
    bbox = (x, y, x + w, y + h)
    shot = ImageGrab.grab(bbox=bbox)
    artifact_dir = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765"
    out_path = os.path.join(artifact_dir, "new_layout_preview.png")
    shot.save(out_path)
    print(f"Saved layout screenshot to {out_path}")

app.destroy()
