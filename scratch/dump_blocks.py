import sys
import os

if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

import pymupdf

f1 = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf" # 1A
doc = pymupdf.open(f1)
page = doc[0]

# Extract words or blocks
blocks = page.get_text("blocks")
with open("scratch/inspect_table_rows.txt", "w", encoding="utf-8") as out:
    out.write(f"=== File 1A Page 1 Blocks ===\n")
    for b in blocks:
        out.write(f"BBox: ({b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}) -> {repr(b[4].strip())}\n")

print("Dumped blocks to scratch/inspect_table_rows.txt")
