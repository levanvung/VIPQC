import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')
for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
    if os.path.exists(p):
        try: os.add_dll_directory(p)
        except Exception: pass
from PIL import Image, ImageDraw, ImageTk
import fitz

ws_path = r"e:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder"
sys.path.insert(0, ws_path)
os.chdir(ws_path)

from model_comparator import render_curtain_drawing_pair

pdf_a = os.path.join(ws_path, "sample_manuals", "media_1788593034521.pdf")
pdf_b = os.path.join(ws_path, "sample_manuals", "media_1788593034536.pdf")

# 1. Measure current cold render time
t0 = time.time()
img1, meta1 = render_curtain_drawing_pair(pdf_a, pdf_b, 1, 1, zoom=1.5, split_ratio=0.5, active_ref="R652")
t_cold = time.time() - t0
print(f"Cold render from PDF took: {t_cold*1000:.1f} ms")

# 2. Test cached slice & composite time
# Suppose we have img_a and img_b pre-rendered
w, h = img1.size
img_a = Image.new("RGB", (w, h), (30, 40, 60))
img_b = Image.new("RGB", (w, h), (10, 20, 30))

times = []
for i in range(100):
    ratio = i / 100.0
    split_x = int(w * ratio)
    t1 = time.time()
    
    # Fast composite
    comp = Image.new("RGB", (w, h))
    comp.paste(img_a.crop((0, 0, split_x, h)), (0, 0))
    comp.paste(img_b.crop((split_x, 0, w, h)), (split_x, 0))
    d = ImageDraw.Draw(comp)
    d.line([(split_x, 0), (split_x, h)], fill="#00C9A7", width=3)
    
    times.append(time.time() - t1)

avg_ms = (sum(times) / len(times)) * 1000
print(f"Cached composite took: {avg_ms:.2f} ms (Speedup: {t_cold*1000 / avg_ms:.0f}x!)")
