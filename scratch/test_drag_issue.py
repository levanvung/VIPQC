import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import tkinter as tk
import glob
from app import BOMExtractorApp, ModelCompareView

def test_drawing_drag():
    pdf_files = sorted(glob.glob("sample_manuals/*.pdf"))
    pdf_a = os.path.abspath(pdf_files[0])
    pdf_b = os.path.abspath(pdf_files[1])

    app = BOMExtractorApp()
    app.geometry("1400x900")
    app._set_nav_active("model_comp")
    app.update()
    
    view = app.view_model_compare
    root = app

    view.pdf_a_path = pdf_a
    view.pdf_b_path = pdf_b
    view.active_page_a = 1
    view.active_page_b = 1
    view.view_mode = "curtain"
    view.split_ratio = 0.5
    view.rotation_angle = 0
    view.zoom_level = 1.0

    print("Initial render...")
    view._render_drawing(center_ref=False)
    root.update()

    cw = view.canvas.winfo_width()
    ch = view.canvas.winfo_height()
    meta = view.current_meta
    print(f"Canvas size: cw={cw}, ch={ch}")
    print(f"Meta drawing size: w={meta['width']}, h={meta['height']}")
    print(f"Draw offset: x={view.draw_offset_x}, y={view.draw_offset_y}")
    print(f"Scrollregion: {view.canvas.cget('scrollregion')}")
    print(f"Initial xview: {view.canvas.xview()}")
    print(f"Initial yview: {view.canvas.yview()}")

    # TEST 1: User tries to pan drawing to the left
    # User clicks at canvas center and drags left by 150px
    class FakeEvent:
        def __init__(self, x, y):
            self.x = x
            self.y = y

    # Suppose user clicks at (cw // 2 - 100, ch // 2) (away from curtain split line)
    click_x = cw // 2 - 100
    click_y = ch // 2
    e_press = FakeEvent(click_x, click_y)
    view._on_canvas_b1_press(e_press)
    print(f"\nClicked at ({click_x}, {click_y})")
    print(f"dragging_curtain = {view.dragging_curtain}")

    # Now drag left by 150px (x becomes click_x - 150)
    e_motion = FakeEvent(click_x - 150, click_y)
    view._on_canvas_b1_motion(e_motion)
    print(f"After dragging mouse left by 150px:")
    print(f"xview = {view.canvas.xview()}")

    # Now drag right by 150px
    e_motion2 = FakeEvent(click_x + 150, click_y)
    view._on_canvas_b1_motion(e_motion2)
    print(f"After dragging mouse right by 150px:")
    print(f"xview = {view.canvas.xview()}")

    # TEST 2: User tries Button 2 (middle click) panning
    view._on_canvas_b2_press(e_press)
    view._on_canvas_b2_motion(e_motion)
    print(f"\nButton 2 drag left: xview = {view.canvas.xview()}")

    # TEST 3: User tries dragging curtain divider to the left
    split_x = meta["split_x"]
    draw_x0 = view.draw_offset_x
    curtain_screen_x = draw_x0 + split_x
    print(f"\nCurtain divider screen X is at: {curtain_screen_x}")
    e_curtain_press = FakeEvent(curtain_screen_x, ch // 2)
    view._on_canvas_b1_press(e_curtain_press)
    print(f"Clicked curtain divider: dragging_curtain = {view.dragging_curtain}")

    # Drag curtain 200px to the left
    e_curtain_drag_left = FakeEvent(curtain_screen_x - 200, ch // 2)
    view._on_canvas_b1_motion(e_curtain_drag_left)
    print(f"After dragging curtain left 200px: split_ratio = {view.split_ratio}")

    # Drag curtain completely to the left edge of canvas (e.g. x = 20)
    e_curtain_drag_far_left = FakeEvent(20, ch // 2)
    view._on_canvas_b1_motion(e_curtain_drag_far_left)
    print(f"After dragging curtain to far left (x=20): split_ratio = {view.split_ratio}")

    # Now test what happens when zoomed in: zoom = 2.5
    print("\n--- Testing with Zoom = 2.5 ---")
    view.zoom_level = 2.5
    view._render_drawing(center_ref=False)
    root.update()
    meta_z = view.current_meta
    print(f"Zoomed meta drawing size: w={meta_z['width']}, h={meta_z['height']}")
    print(f"Zoomed scrollregion: {view.canvas.cget('scrollregion')}")
    print(f"Zoomed initial xview: {view.canvas.xview()}")

    # Zoomed: user tries to drag drawing left (mouse moves from 500 to 300)
    e_zoom_press = FakeEvent(500, 300)
    view._on_canvas_b1_press(e_zoom_press)
    print(f"Zoomed click at (500, 300): dragging_curtain = {view.dragging_curtain}")
    e_zoom_drag = FakeEvent(300, 300)
    view._on_canvas_b1_motion(e_zoom_drag)
    print(f"Zoomed drag mouse left (500->300): xview = {view.canvas.xview()}")

    # Zoomed: user tries to drag mouse right (500->700) (wanting to pull drawing to the right to see left side)
    view._on_canvas_b1_press(e_zoom_press)
    e_zoom_drag_r = FakeEvent(700, 300)
    view._on_canvas_b1_motion(e_zoom_drag_r)
    print(f"Zoomed drag mouse right (500->700): xview = {view.canvas.xview()}")

    root.destroy()
    print("\nTest completed successfully.")

if __name__ == "__main__":
    test_drawing_drag()
