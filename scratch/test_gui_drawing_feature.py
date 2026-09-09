"""
Test script for verifying Series BOM Drawing Integration & Border Highlighting
"""

import os
import sys
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk

sys.path.insert(0, os.path.abspath("."))
import series_bom_comparator as sbc

ctk.set_appearance_mode("dark")

print("Initializing test for drawing integration...")

f_a = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
f_b = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"
drawing_sample = r"sample_manuals\media_1788593034488.pdf"

bom_a = sbc.parse_any_bom(f_a)
bom_b = sbc.parse_any_bom(f_b)
res = sbc.compare_series_boms(bom_a, bom_b)

print(f"Comparison ready: {len(res['qc_focus_items'])} focus items.")

# Test rendering annotated drawing page
pages = sbc.get_drawing_pages_info(drawing_sample)
print(f"Drawing pages found: {len(pages)}")

img, coords = sbc.render_annotated_drawing_page(
    drawing_sample, 0, res['qc_focus_items'], active_loc="R7501", zoom=1.0, rotation=0
)

print(f"Rendered annotated drawing: size={img.size if img else None}, coords count={len(coords)}")
if img:
    img.save("scratch/test_series_drawing_annotated.png")
    print("Saved to scratch/test_series_drawing_annotated.png")

print("Test passed successfully!")
