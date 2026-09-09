import sys
import os
import re

sys.stdout.reconfigure(encoding='utf-8')

if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

import pymupdf

def parse_full_erp_bom(pdf_path):
    doc = pymupdf.open(pdf_path)
    main_part_no = ""
    records = []
    
    for page_idx, page in enumerate(doc):
        text = page.get_text()
        if not main_part_no:
            m = re.search(r"主件料號:\s*([A-Za-z0-9\-_]+)", text)
            if m:
                main_part_no = m.group(1).strip()
                
        words = page.get_text("words")
        # Filter words within table body (y from 155 to 1080)
        table_words = [w for w in words if 155 <= w[1] <= 1075]
        
        # Group words into rows based on item_no or y coordinates
        # Notice: item numbers appear around x0: 50-80 with numeric values (10, 20, 30...)
        item_words = [w for w in table_words if 50 <= w[0] <= 85 and re.match(r"^\d+$", w[4])]
        item_words.sort(key=lambda w: w[1])
        
        for i, item_w in enumerate(item_words):
            y_start = item_w[1] - 4
            y_end = item_words[i+1][1] - 4 if i + 1 < len(item_words) else 1075
            
            row_words = [w for w in table_words if y_start <= w[1] < y_end]
            
            # Extract fields by x coordinates
            # Level: x < 50
            level = " ".join([w[4] for w in row_words if w[0] < 50]).strip()
            item_no = item_w[4]
            # Part number: 95 <= x0 <= 220
            part_no = " ".join([w[4] for w in row_words if 95 <= w[0] <= 220]).strip()
            # Old part: 220 < x0 <= 280
            old_part = " ".join([w[4] for w in row_words if 220 < w[0] <= 280]).strip()
            # Name & Spec: 280 < x0 <= 490
            spec = " ".join([w[4] for w in row_words if 280 < w[0] <= 490]).strip()
            # Source: 490 < x0 <= 525
            src = " ".join([w[4] for w in row_words if 490 < w[0] <= 525]).strip()
            # Unit: 525 < x0 <= 565
            unit = " ".join([w[4] for w in row_words if 525 < w[0] <= 565]).strip()
            # Qty: 565 < x0 <= 620
            qty = " ".join([w[4] for w in row_words if 565 < w[0] <= 620]).strip()
            # Loss: 620 < x0 <= 670
            loss = " ".join([w[4] for w in row_words if 620 < w[0] <= 670]).strip()
            # Locations: x0 >= 670
            loc_words = [w[4] for w in row_words if w[0] >= 670]
            loc_str = "".join(loc_words)
            # Parse individual locations: e.g. Q8230, Q8209...
            locations = [loc.strip() for loc in re.findall(r"[A-Za-z0-9\-_]+", loc_str) if loc.strip()]
            
            if part_no:
                records.append({
                    "item_no": item_no,
                    "level": level,
                    "part_no": part_no,
                    "spec": spec,
                    "qty": qty,
                    "locations": locations,
                    "page": page_idx + 1
                })
                
    return main_part_no, records

f_1a = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
f_1d = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"

m_1a, rec_1a = parse_full_erp_bom(f_1a)
m_1d, rec_1d = parse_full_erp_bom(f_1d)

print(f"Model 1A ({m_1a}): Total records = {len(rec_1a)}")
print(f"Model 1D ({m_1d}): Total records = {len(rec_1d)}")

# Map each location to part_no
loc_map_1a = {}
for r in rec_1a:
    for loc in r["locations"]:
        loc_map_1a[loc] = r

loc_map_1d = {}
for r in rec_1d:
    for loc in r["locations"]:
        loc_map_1d[loc] = r

all_locs = sorted(list(set(loc_map_1a.keys()) | set(loc_map_1d.keys())))
print(f"Total unique locations across both series: {len(all_locs)}")

# Compare differences!
added_in_1d = []
removed_in_1d = []
modified_parts = []
same_parts = []

for loc in all_locs:
    in_a = loc_map_1a.get(loc)
    in_d = loc_map_1d.get(loc)
    
    if in_a and not in_d:
        removed_in_1d.append((loc, in_a["part_no"], in_a["spec"]))
    elif not in_a and in_d:
        added_in_1d.append((loc, in_d["part_no"], in_d["spec"]))
    elif in_a["part_no"] != in_d["part_no"]:
        modified_parts.append((loc, in_a["part_no"], in_d["part_no"], in_a["spec"], in_d["spec"]))
    else:
        same_parts.append(loc)

print("\n--- COMPARISON SUMMARY ---")
print(f"⚪ Common / Same locations: {len(same_parts)}")
print(f"🟡 Modified / Changed Part Number at same Location: {len(modified_parts)}")
print(f"🟢 Added in Series 1D: {len(added_in_1d)}")
print(f"🔴 Removed in Series 1D (present in 1A): {len(removed_in_1d)}")

if modified_parts:
    print("\nSAMPLE MODIFIED LOCATIONS:")
    for loc, p_a, p_d, s_a, s_d in modified_parts[:10]:
        print(f"  [{loc}] 1A: {p_a} --> 1D: {p_d}")

if added_in_1d:
    print("\nSAMPLE ADDED LOCATIONS in 1D:")
    for loc, p, s in added_in_1d[:10]:
        print(f"  + [{loc}] Part: {p}")

if removed_in_1d:
    print("\nSAMPLE REMOVED LOCATIONS in 1D:")
    for loc, p, s in removed_in_1d[:10]:
        print(f"  - [{loc}] Part: {p}")
