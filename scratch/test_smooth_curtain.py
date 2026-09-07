import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')
for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
    if os.path.exists(p):
        try: os.add_dll_directory(p)
        except Exception: pass

ws_path = r"e:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder"
sys.path.insert(0, ws_path)
os.chdir(ws_path)

import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk, ImageDraw
import fitz

root = ctk.CTk()
root.geometry("800x600")

# Simulate base caching
img_a = Image.new("RGB", (1000, 800), (20, 60, 100))
img_b = Image.new("RGB", (1000, 800), (100, 30, 40))

canvas = tk.Canvas(root, width=800, height=600, bg="#0B0F1A")
canvas.pack(fill="both", expand=True)

photo = None
img_id = None

def render_curtain(ratio):
    global photo, img_id
    t0 = time.time()
    w, h = 1000, 800
    split_x = int(w * ratio)
    
    comp = img_b.copy()
    if split_x > 0:
        comp.paste(img_a.crop((0, 0, split_x, h)), (0, 0))
    d = ImageDraw.Draw(comp)
    d.line([(split_x, 0), (split_x, h)], fill="#22D3EE", width=3)
    
    photo = ImageTk.PhotoImage(comp)
    if img_id is None:
        img_id = canvas.create_image(0, 0, image=photo, anchor="nw")
    else:
        canvas.itemconfig(img_id, image=photo)
    return time.time() - t0

times = []
for i in range(50):
    t = render_curtain(i / 50.0)
    times.append(t)
    root.update()

avg_fps = 1.0 / (sum(times) / len(times))
print(f"Smooth curtain render: {avg_fps:.1f} FPS (Frame time: {(sum(times)/len(times))*1000:.2f} ms)")
root.destroy()
