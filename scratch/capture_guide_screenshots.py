import os
import sys
import time
import tkinter as tk
from PIL import ImageGrab, Image

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app
from model_comparator import get_annotated_base_images, render_curtain_drawing_pair

def capture_screens():
    brain_dir = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\guide_images"
    local_dir = os.path.abspath("guide_images")
    os.makedirs(brain_dir, exist_ok=True)
    os.makedirs(local_dir, exist_ok=True)

    print("Launching VIPQC App for guide screenshots...")
    app_instance = app.BOMExtractorApp()
    app_instance.geometry("1380x880+30+30")
    app_instance.update()

    def grab_and_save(name):
        time.sleep(0.3)
        app_instance.update()
        x = app_instance.winfo_rootx()
        y = app_instance.winfo_rooty()
        w = app_instance.winfo_width()
        h = app_instance.winfo_height()
        bbox = (x, y, x + w, y + h)
        img = ImageGrab.grab(bbox=bbox)
        p_brain = os.path.join(brain_dir, f"{name}.png")
        p_local = os.path.join(local_dir, f"{name}.png")
        img.save(p_brain)
        img.save(p_local)
        print(f"Captured: {name} ({w}x{h}) -> {p_brain}")

    # 1. Screen 1: Tab 1 (Trích xuất BOM) with sample files
    sample_pdfs = [
        os.path.abspath("sample_manuals/media_1788593034506.pdf"),
        os.path.abspath("sample_manuals/media_1788593034536.pdf")
    ]
    app_instance._set_nav_active("extract")
    app_instance._on_drop_files([f for f in sample_pdfs if os.path.exists(f)])
    from bom_extractor import extract_full_bom
    if app_instance.selected_files:
        meta, flat, _ = extract_full_bom(app_instance.selected_files[0])
        app_instance.all_records = flat[:60]
        app_instance._populate_tree(app_instance.all_records)
        app_instance._rebuild_filter_tabs()
        app_instance._refresh_badges()
    app_instance._set_nav_active("extract")
    app_instance.update()
    grab_and_save("01_tab_trich_xuat_bom")

    # 2. Screen 2: Tab 2 (Đối Chiếu BOM)
    app_instance._set_nav_active("compare")
    app_instance.update()
    grab_and_save("02_tab_doi_chieu_bom")

    # 3. Screen 3: Tab 3 (So Sánh 2 Model & Bản Vẽ PCB)
    app_instance._set_nav_active("model_comp")
    tab3 = app_instance.view_model_compare
    f1 = os.path.abspath(r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788678922426.pdf")
    f2 = os.path.abspath(r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788678927420.pdf")
    if os.path.exists(f1) and os.path.exists(f2):
        tab3.pdf_a_path = f1
        tab3.pdf_b_path = f2
        tab3.lbl_model_a_name.configure(text="📄 CHP3178AF-1A MP.pdf (335.5 KB)")
        tab3.lbl_model_a_sub.configure(text="Model: QPWBCAF3178  |  Series: 1A")
        tab3.lbl_model_b_name.configure(text="📄 CHP3178AF-1B MP.pdf (338.2 KB)")
        tab3.lbl_model_b_sub.configure(text="Model: QPWBCAF3178  |  Series: 1B")
        # Pre-render curtain drawing
        tab3._on_display_mode_change("◫ Song Song")
        tab3.seg_display_mode.set("◫ Song Song")
        tab3.seg_view_mode.set("↔️ Kéo Màn")
        tab3.split_ratio = 0.5
        tab3._render_drawing(center_ref=False)
        app_instance.update()
        grab_and_save("03_tab_so_sanh_model_curtain")

        # 4. Screen 4: Tab 3 (Full Drawing View - 100% Bản Vẽ)
        tab3._on_display_mode_change("🖼️ Bản Vẽ")
        tab3.seg_display_mode.set("🖼️ Bản Vẽ")
        tab3._render_drawing(center_ref=False)
        app_instance.update()
        grab_and_save("04_che_do_xem_ban_ve_toan_dien")

    app_instance.destroy()
    print("All guide screenshots successfully captured!")

if __name__ == "__main__":
    capture_screens()
