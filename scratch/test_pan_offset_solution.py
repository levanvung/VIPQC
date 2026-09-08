import sys, os, glob
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import tkinter as tk
from app import BOMExtractorApp, DrawingFullScreenWindow

class FakeEvent:
    def __init__(self, x, y):
        self.x = x
        self.y = y

def test_pan_solution():
    print("=" * 60)
    print("VERIFICATION: CANVAS PANNING & CURTAIN DRAG FIX")
    print("=" * 60)

    app = BOMExtractorApp()
    app.geometry("1400x900")
    app._set_nav_active("model_comp")
    app.deiconify()
    app.update_idletasks()
    app.update()

    v = app.view_model_compare

    # ── TEST 1: Split View Layout ─────────────────────────────────────────────
    print("\n--- TEST 1: Split View Layout Dimensions ---")
    tc_w = v.table_card.winfo_width()
    ic_w = v.inspector_card.winfo_width()
    cv_w = v.canvas.winfo_width()
    cv_h = v.canvas.winfo_height()
    print(f"table_card width: {tc_w} px")
    print(f"inspector_card width: {ic_w} px")
    print(f"canvas width: {cv_w} px, height: {cv_h} px")

    assert ic_w > 400, f"Inspector card is too narrow: {ic_w}px (should be ~550px)!"
    assert cv_w > 400, f"Canvas is too narrow: {cv_w}px!"
    print("[PASS] TEST 1 PASSED: Split View has balanced 50:50 layout with large canvas!")

    # ── TEST 2: Load PDFs and Initial Render ──────────────────────────────────
    print("\n--- TEST 2: Preload Drawing & Initial Coordinates ---")
    pdf_files = sorted(glob.glob("sample_manuals/*.pdf"))
    assert len(pdf_files) >= 2, "Need at least 2 sample PDF files"
    v.pdf_a_path = os.path.abspath(pdf_files[0])
    v.pdf_b_path = os.path.abspath(pdf_files[1])
    v.active_page_a = 1
    v.active_page_b = 1
    v.view_mode = "curtain"
    v.split_ratio = 0.5
    v.rotation_angle = 0
    v.zoom_level = 1.0

    v._render_drawing(center_ref=False)
    app.update()

    init_coords = v.canvas.coords(v.canvas_img_id)
    init_x = init_coords[0]
    init_y = init_coords[1]
    print(f"Initial image coords on canvas: X={init_x}, Y={init_y}")
    print(f"Meta drawing width: {v.current_meta.get('width')}, height: {v.current_meta.get('height')}")

    # ── TEST 3: Pan Drawing to the LEFT (Mouse Drag Left) ─────────────────────
    print("\n--- TEST 3: Pan Drawing to the LEFT ---")
    # Click somewhere away from curtain divider (e.g. at x = 100, y = 150)
    v._on_canvas_b1_press(FakeEvent(100, 150))
    assert not v.dragging_curtain, "Should be in canvas pan mode, not curtain drag!"

    # Drag left by 180px (from 100 to -80)
    v._on_canvas_b1_motion(FakeEvent(-80, 150))
    v._on_canvas_b1_release(FakeEvent(-80, 150))
    app.update()

    panned_left_coords = v.canvas.coords(v.canvas_img_id)
    new_x = panned_left_coords[0]
    print(f"After panning left 180px: X={new_x}, Y={panned_left_coords[1]}")
    expected_x = init_x - 180
    assert abs(new_x - expected_x) < 1e-3, f"Expected X={expected_x}, got {new_x}!"
    print(f"[PASS] TEST 3 PASSED: Drawing successfully moved {init_x - new_x}px to the LEFT!")

    # ── TEST 4: Pan Drawing to the RIGHT ──────────────────────────────────────
    print("\n--- TEST 4: Pan Drawing to the RIGHT ---")
    # Click away from curtain divider (at x=97), e.g. at x=300
    v._on_canvas_b1_press(FakeEvent(300, 150))
    assert not v.dragging_curtain, "Should be in canvas pan mode!"
    # Drag right by 250px (from 300 to 550)
    v._on_canvas_b1_motion(FakeEvent(550, 150))
    v._on_canvas_b1_release(FakeEvent(550, 150))
    app.update()

    panned_right_coords = v.canvas.coords(v.canvas_img_id)
    new_x2 = panned_right_coords[0]
    print(f"After panning right 250px: X={new_x2}")
    expected_x2 = expected_x + 250
    assert abs(new_x2 - expected_x2) < 1e-3, f"Expected X={expected_x2}, got {new_x2}!"
    print(f"[PASS] TEST 4 PASSED: Drawing successfully moved {new_x2 - expected_x}px to the RIGHT!")

    # ── TEST 5: Middle-Click (Button 2) Pan ────────────────────────────────────
    print("\n--- TEST 5: Button 2 Middle-Click Pan ---")
    v._on_canvas_b2_press(FakeEvent(200, 200))
    v._on_canvas_b2_motion(FakeEvent(100, 200)) # drag left 100px
    app.update()
    b2_coords = v.canvas.coords(v.canvas_img_id)
    print(f"After Button 2 drag left 100px: X={b2_coords[0]}")
    assert abs(b2_coords[0] - (new_x2 - 100)) < 1e-3
    print("[PASS] TEST 5 PASSED: Button 2 middle-click pan operates flawlessly!")

    # ── TEST 6: Curtain Dragging to 0% (Far Left Edge) ────────────────────────
    print("\n--- TEST 6: Curtain Wipe Dragging to 0% (Far Left) ---")
    # Reset zoom/fit to center
    v._zoom_fit()
    app.update()
    draw_x0 = v.draw_offset_x
    split_x = v.current_meta["split_x"]
    curtain_x = draw_x0 + v.pan_offset_x + split_x
    print(f"Curtain divider screen X is at: {curtain_x}")

    # Press directly on curtain divider
    v._on_canvas_b1_press(FakeEvent(curtain_x, 150))
    assert v.dragging_curtain, f"Should detect curtain divider click at {curtain_x}!"

    # Drag left into margin (e.g. x = draw_x0 - 50, past the left edge of drawing)
    v._on_canvas_b1_motion(FakeEvent(draw_x0 - 50, 150))
    print(f"Curtain ratio after dragging past left edge: {v.split_ratio}")
    assert v.split_ratio == 0.0, f"Expected split_ratio=0.0, got {v.split_ratio}!"

    # Release mouse
    v._on_canvas_b1_release(FakeEvent(draw_x0 - 50, 150))
    app.update()
    print("[PASS] TEST 6 PASSED: Curtain successfully drags all the way to 0.0 (far left)!")

    # ── TEST 7: Fullscreen Modal Panning ──────────────────────────────────────
    print("\n--- TEST 7: Fullscreen Window 2D Panning & Curtain ---")
    fs = DrawingFullScreenWindow(
        parent=app,
        app=app,
        pdf_a_path=v.pdf_a_path,
        pdf_b_path=v.pdf_b_path,
        comp_result={},
        drawing_catalog=v.drawing_catalog,
        active_item=None,
        current_stage_id="ALL",
        active_page_a=1,
        active_page_b=1,
        split_ratio=0.5,
        view_mode="curtain",
        zoom_level=1.0,
        rotation_angle=0
    )
    fs.deiconify()
    fs.update_idletasks()
    fs.update()

    fs_init_coords = fs.canvas.coords(fs.canvas_img_id)
    print(f"Fullscreen initial coords: X={fs_init_coords[0]}, Y={fs_init_coords[1]}")

    # Drag left in fullscreen
    fs._on_canvas_b1_press(FakeEvent(100, 200))
    fs._on_canvas_b1_motion(FakeEvent(-50, 200))
    fs._on_canvas_b1_release(FakeEvent(-50, 200))
    fs.update()

    fs_left_coords = fs.canvas.coords(fs.canvas_img_id)
    print(f"Fullscreen after drag left 150px: X={fs_left_coords[0]}")
    assert abs(fs_left_coords[0] - (fs_init_coords[0] - 150)) < 1e-3
    print("[PASS] TEST 7 PASSED: Fullscreen window pans left smoothly!")

    fs.destroy()
    app.destroy()

    print("\n" + "=" * 60)
    print("ALL 7 VERIFICATION TESTS PASSED 100%!")
    print("=" * 60)

if __name__ == "__main__":
    test_pan_solution()
