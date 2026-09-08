import sys, os, glob
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import BOMExtractorApp

app = BOMExtractorApp()
app.geometry('1400x900')
app._set_nav_active('model_comp')
v = app.view_model_compare
v._on_display_mode_change('bản vẽ')
app.deiconify()
app.update_idletasks()
app.update()

pdf_files = sorted(glob.glob('sample_manuals/*.pdf'))
v.pdf_a_path = os.path.abspath(pdf_files[0])
v.pdf_b_path = os.path.abspath(pdf_files[1])
v.active_page_a = 1
v.active_page_b = 1
v.view_mode = 'curtain'
v.split_ratio = 0.5
v.rotation_angle = 0
v.zoom_level = 1.0

v._render_drawing(center_ref=False)
app.update()

cw = v.canvas.winfo_width()
ch = v.canvas.winfo_height()
meta = v.current_meta
print(f"Canvas cw={cw}, ch={ch}")
print(f"Drawing w={meta['width']}, h={meta['height']}")
print(f"draw_offset_x = {v.draw_offset_x}")
print(f"scrollregion = {v.canvas.cget('scrollregion')}")
print(f"Initial xview = {v.canvas.xview()}")

class FakeEvent:
    def __init__(self, x, y):
        self.x = x
        self.y = y

# Test 1: User tries to drag the drawing to the left
# Clicks on drawing (e.g. at x=draw_offset_x + 100) and drags left by 100px
click_x = v.draw_offset_x + 100
v._on_canvas_b1_press(FakeEvent(click_x, 150))
print(f"Press at {click_x}: dragging_curtain = {v.dragging_curtain}")
v._on_canvas_b1_motion(FakeEvent(click_x - 100, 150))
print(f"After drag mouse left 100px: xview = {v.canvas.xview()}")
v._on_canvas_b1_motion(FakeEvent(click_x + 100, 150))
print(f"After drag mouse right 100px: xview = {v.canvas.xview()}")

# Test 2: What happens when zoomed in to 2.0?
print("\n--- Zoomed 2.0x ---")
v.zoom_level = 2.0
v._render_drawing(center_ref=False)
app.update()
meta2 = v.current_meta
print(f"Zoomed w={meta2['width']}, h={meta2['height']}")
print(f"draw_offset_x = {v.draw_offset_x}")
print(f"scrollregion = {v.canvas.cget('scrollregion')}")
print(f"Initial zoomed xview = {v.canvas.xview()}")

# Now drag left
v._on_canvas_b1_press(FakeEvent(500, 150))
print(f"Zoomed press: dragging_curtain = {v.dragging_curtain}")
v._on_canvas_b1_motion(FakeEvent(300, 150))
print(f"Zoomed drag mouse left (500->300): xview = {v.canvas.xview()}")
# Now drag right
v._on_canvas_b1_press(FakeEvent(500, 150))
v._on_canvas_b1_motion(FakeEvent(700, 150))
print(f"Zoomed drag mouse right (500->700): xview = {v.canvas.xview()}")

app.destroy()
print("Done")
