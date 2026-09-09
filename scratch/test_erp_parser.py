import sys
import os
import re

if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

import pymupdf

def extract_erp_bom(pdf_path):
    doc = pymupdf.open(pdf_path)
    main_part_no = ""
    items = []
    
    for page_idx, page in enumerate(doc):
        text = page.get_text()
        
        # Check main part no in header
        if not main_part_no:
            m = re.search(r"主件料號:\s*([A-Za-z0-9\-_]+)", text)
            if m:
                main_part_no = m.group(1)
                
        # Parse table lines
        # Let's extract words with coordinates
        words = page.get_text("words") # (x0, y0, x1, y1, word, block_no, line_no, word_no)
        # We can also parse blocks or lines
        
    return {
        "file": os.path.basename(pdf_path),
        "main_part_no": main_part_no,
        "total_pages": len(doc)
    }

print("File 1 (499):", extract_erp_bom(r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"))
print("File 2 (554):", extract_erp_bom(r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"))
