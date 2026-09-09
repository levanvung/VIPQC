import sys
sys.stdout.reconfigure(encoding="utf-8")
import os
import time

sys.path.insert(0, os.path.abspath("."))
import app as app_mod

f_a = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
f_b = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"

gui = app_mod.BOMExtractorApp()

def do_capture():
    gui._set_nav_active("series_bom")
    view = gui.view_series_bom_compare
    view.file_a = f_a
    view.file_b = f_b
    view.lbl_file_a_name.configure(text=os.path.basename(f_a))
    view.lbl_file_b_name.configure(text=os.path.basename(f_b))
    
    view._start_compare()

    def check_done(count=0):
        if not view.is_comparing and view.comparison_result:
            gui.update()
            time.sleep(0.5)
            # Take screenshot using PIL ImageGrab if available
            try:
                from PIL import ImageGrab
                x = gui.winfo_rootx()
                y = gui.winfo_rooty()
                w = gui.winfo_width()
                h = gui.winfo_height()
                img = ImageGrab.grab((x, y, x + w, y + h))
                out_path = "guide_images/08_so_sanh_2_bom_toan_man_hinh.png"
                img.save(out_path)
                print(f"Captured screenshot to {out_path}")
            except Exception as e:
                print(f"Screenshot notice: {e}")
            gui.destroy()
        else:
            gui.after(200, lambda: check_done(count + 1))

    gui.after(100, check_done)

gui.after(200, do_capture)
gui.mainloop()
