import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

import pymupdf

f = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
doc = pymupdf.open(f)
page = doc[0]

words = page.get_text("words")
header_words = [w for w in words if 120 <= w[1] <= 160]
header_words.sort(key=lambda w: w[0])
print("Header words with x-coords:")
for w in header_words:
    print(f"  x0={w[0]:.1f}, x1={w[2]:.1f}, y0={w[1]:.1f} -> '{w[4]}'")

row_words = [w for w in words if 180 <= w[1] <= 320]
row_words.sort(key=lambda w: (w[1], w[0]))
with open("scratch/inspect_words.txt", "w", encoding="utf-8") as out:
    for w in row_words:
        out.write(f"y0={w[1]:.1f}, x0={w[0]:.1f}, x1={w[2]:.1f} -> '{w[4]}'\n")
print("Saved row words to scratch/inspect_words.txt")
