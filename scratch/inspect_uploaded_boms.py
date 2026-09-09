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

files = [
    r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf",
    r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
]

with open("scratch/inspect_uploaded_boms.txt", "w", encoding="utf-8") as out:
    for path in files:
        doc = pymupdf.open(path)
        out.write(f"\n{'='*50}\nFILE: {os.path.basename(path)} | Total Pages: {len(doc)}\n{'='*50}\n")
        for p in range(len(doc)):
            page = doc[p]
            text = page.get_text()
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            out.write(f"\n--- Page {p+1} (Rect: {page.rect}, Rotation: {page.rotation}) ---\n")
            out.write("\n".join(lines[:35]))
            out.write("\n...\n")

print("Inspection completed successfully. See scratch/inspect_uploaded_boms.txt")
