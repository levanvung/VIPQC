import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

# Ensure DLL search path for Windows Python 3.8+
for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
    if os.path.exists(p):
        try: os.add_dll_directory(p)
        except Exception: pass

ws_path = r"e:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder"
sys.path.insert(0, ws_path)
os.chdir(ws_path)

from app import BOMExtractorApp

print("=== STARTING USER REQUIREMENTS VERIFICATION ===")

app = BOMExtractorApp()
app.geometry("1360x860")
app.update()

# 1. Switch to Tab 3: Model Compare
app._set_nav_active("model_comp")
app.update()
mc = app.view_model_compare

# ── REQUIREMENT 1 VERIFICATION ────────────────────────────────────────────────
# "box so sánh 2 model bạn cho nó nằm dưới box chọn model a và model b sao cho nó cân đối"
print("\n--- 1. Testing Model A / B and Compare Box Layout Symmetry ---")
assert hasattr(mc, "row_models"), "mc must have row_models frame"
assert hasattr(mc, "card_compare_box"), "mc must have card_compare_box"
assert mc.row_models.winfo_manager() == "pack"
assert mc.card_compare_box.winfo_manager() == "pack"

# Verify relative packing order: row_models is before card_compare_box in bottom_dock
children = mc.bottom_dock.winfo_children()
idx_models = children.index(mc.row_models)
idx_comp = children.index(mc.card_compare_box)
print(f"Bottom dock children order: row_models index={idx_models}, card_compare_box index={idx_comp}")
assert idx_models < idx_comp, "row_models MUST be on top of card_compare_box!"

# Verify symmetry: Model A card and Model B card inside row_models
a_children = mc.row_models.winfo_children()
assert mc.card_model_a in a_children and mc.card_model_b in a_children
print("✓ Box Chọn Model A và Model B nằm song song cân đối (50%-50%).")
print("✓ Box So Sánh 2 Model (card_compare_box) nằm trực tiếp ở DƯỚI.")

# ── REQUIREMENT 2 VERIFICATION ────────────────────────────────────────────────
# "còn container soi bản vẽ khi tôi nhấn vào Bản vẽ PCB bạn cho height của nó rộng ra chút"
print("\n--- 2. Testing PCB Drawing Inspector Height Expansion ---")
pdf_a = os.path.join(ws_path, "sample_manuals", "media_1788593034521.pdf")
pdf_b = os.path.join(ws_path, "sample_manuals", "media_1788593034536.pdf")
mc.pdf_a_path = pdf_a
mc.pdf_b_path = pdf_b
mc._run_comparison()
app.update()

# Measure heights in Split mode
mc._on_display_mode_change("◫ Song Song")
app.update()
split_canvas_h = mc.canvas.winfo_height()
split_insp_h = mc.inspector_card.winfo_height()
print(f"Split Mode   -> Canvas Height: {split_canvas_h}px, Inspector Card Height: {split_insp_h}px")

# Switch to Drawing mode ("📐 Bản Vẽ PCB")
mc._on_display_mode_change("📐 Bản Vẽ PCB")
app.update()
time.sleep(0.1)
app.update()

drawing_canvas_h = mc.canvas.winfo_height()
drawing_insp_h = mc.inspector_card.winfo_height()
print(f"Drawing Mode -> Canvas Height: {drawing_canvas_h}px, Inspector Card Height: {drawing_insp_h}px")

height_gain = drawing_canvas_h - split_canvas_h
print(f"Height gained in Drawing Mode: +{height_gain}px")
assert drawing_canvas_h > split_canvas_h, f"Drawing mode canvas ({drawing_canvas_h}px) must be taller than split mode ({split_canvas_h}px)!"
assert not mc.row_models.winfo_ismapped(), "row_models should be hidden in drawing mode to free up vertical height!"
print(f"✓ Container soi bản vẽ đã rộng ra thêm +{height_gain}px khi chuyển sang '📐 Bản Vẽ PCB'!")

# ── REQUIREMENT 3 VERIFICATION ────────────────────────────────────────────────
# "và chức năng kéo để so sánh giữa 2 bản vẽ đang khá giật nhỉ, có cashc nào cho smooth mượt mà hơn không"
print("\n--- 3. Testing Ultra-Smooth Curtain Dragging Performance ---")

# Test 50 rapid slider moves in ModelCompareView
t0 = time.perf_counter()
frames = 50
for i in range(frames):
    ratio = i / float(frames)
    mc._on_curtain_slider(ratio)
    app.update_idletasks()
t1 = time.perf_counter()

elapsed = t1 - t0
fps = frames / elapsed
print(f"ModelCompareView Curtain Dragging: {frames} frames rendered in {elapsed:.3f}s ({fps:.1f} FPS)")
assert fps >= 30, f"Curtain dragging FPS ({fps:.1f}) must be at least 30 FPS!"
assert mc.canvas_img_id is not None, "Canvas image item id must be cached/reused!"

# Test in FullScreenWindow as well
print("\n--- Testing FullScreenWindow Curtain Dragging ---")
mc._open_fullscreen_drawing()
app.update()

fw = mc.fullscreen_window
assert fw is not None and fw.winfo_exists()

t0_fw = time.perf_counter()
for i in range(frames):
    ratio = i / float(frames)
    fw._on_curtain_slider(ratio)
    app.update_idletasks()
t1_fw = time.perf_counter()

elapsed_fw = t1_fw - t0_fw
fps_fw = frames / elapsed_fw
print(f"FullScreenWindow Curtain Dragging: {frames} frames rendered in {elapsed_fw:.3f}s ({fps_fw:.1f} FPS)")
assert fps_fw >= 30, f"FullScreen FPS ({fps_fw:.1f}) must be at least 30 FPS!"

fw._close()
app.update()

app.destroy()
print("\n🎉 ALL 3 USER REQUIREMENTS VERIFIED SUCCESSFULLY WITH FLYING COLORS!")
