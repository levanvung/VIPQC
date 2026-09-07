"""
BOM Extractor Engine for Electronics Working Manuals
Extracts structured BOM data from PDF working manuals (e.g. Sharp, New Advanced Electronics, etc.)
Supports varying page rotations, multi-column tables, multi-variant assemblies, and alternate components.
"""

import os
import sys
import re

if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

import pymupdf


def classify_header_span(text: str):
    """Classifies a table header span into a semantic column role."""
    t = text.strip().upper()
    if t in ["NO", "NO."]:
        return "NO"
    if "S.N / RATING" in t or "S.N/RATING" in t or "S.N / VALUE" in t:
        return "SN_RATING"
    if t in ["S.N.", "S.N", "SN", "S.N. "]:
        return "SN"
    if "PART NO" in t or "PART NO." in t or "PART NUMBER" in t or t == "PART":
        return "PART"
    if any(k in t for k in ["SNM CODE", "S&O CODE", "S & O CODE", "PART CODE", "SNM"]):
        return "CODE"
    if any(k in t for k in ["RATING", "VALUE", "SPEC", "SPECIFICATION"]):
        return "RATING"
    if t in ["STN", "STATION"]:
        return "STN"
    if t in ["QTY", "QTY."]:
        return "QTY"
    if "REMARK" in t:
        return "REMARK"
    return None


def is_rating_string(txt: str) -> bool:
    """Detects if a text string is likely an electrical rating/spec rather than a part number."""
    t = txt.strip()
    if re.search(r'\b(?:RS|RN|RD|RF|RC|RU|RK|RM)[0-9\/\.]*(?:W|[A-Z][0-9])[A-Z0-9\-]*', t, re.IGNORECASE):
        return True
    if re.search(r'\b\d+(?:\.\d+)?\s*(?:uF|pF|nF|ohm|OHM|k|K|M|V|DCV|VAC|mA|A)\b', t, re.IGNORECASE):
        return True
    if re.search(r'\b\d+\/\d+W\b', t, re.IGNORECASE):
        return True
    if re.match(r'^(?:2S[ACD]|KR[AC]|1S[S1-9]|1N[0-9]|RL1N)', t):
        return True
    if re.match(r'^(?:CK|CC|CE|ECQ)[0-9A-Z\-]+', t):
        return True
    return False


def is_sharp_part_code(txt: str) -> bool:
    """Detects if a text string is a Sharp Part Number / S&O Code."""
    t = txt.strip()
    if re.match(r'^(?:R[RCHMFDN]|Q[SWA]|V[SCDHKPR][A-Z0-9])[A-Z0-9\+\-\/\.]{5,}', t):
        if not is_rating_string(t):
            return True
        if re.search(r'(?:AWZZ|WJZZ|TAZZ|AF1|AMX|JY|KY|1Y|1L|1T|\+)$', t):
            return True
    return False


def sanitize_part_and_rating(part_no: str, rating: str):
    """
    Ensures that if rating and part number were concatenated (e.g. 'RS1WBJ-10 RR-SZA016AWZZ+'),
    they are cleanly split: part_no='RR-SZA016AWZZ+', rating='RS1WBJ-10'.
    """
    p = part_no.strip()
    r = rating.strip()
    tokens = p.split()
    if len(tokens) >= 2:
        code_tokens = []
        rating_tokens = []
        for tok in tokens:
            if is_sharp_part_code(tok):
                code_tokens.append(tok)
            elif is_rating_string(tok):
                rating_tokens.append(tok)
            else:
                code_tokens.append(tok)
        if rating_tokens and code_tokens:
            p = " ".join(code_tokens)
            added_r = " ".join(rating_tokens)
            r = f"{added_r} {r}".strip() if r else added_r
    return p, r


def get_visual_spans(page):
    """
    Extracts all text spans from a PDF page transformed into human visual reading coordinates,
    accounting for page rotation.
    """
    mat = page.rotation_matrix
    dict_data = page.get_text("dict")
    spans = []
    
    for b in dict_data.get("blocks", []):
        if "lines" in b:
            for l in b["lines"]:
                for s in l["spans"]:
                    text = s["text"].strip()
                    if not text:
                        continue
                    x0, y0, x1, y1 = s["bbox"]
                    pt0 = pymupdf.Point(x0, y0)
                    pt1 = pymupdf.Point(x1, y1)
                    v_rect = pymupdf.Rect(pt0 * mat, pt1 * mat).normalize()
                    spans.append({
                        "text": text,
                        "x0": v_rect.x0,
                        "y0": v_rect.y0,
                        "x1": v_rect.x1,
                        "y1": v_rect.y1,
                        "font": s.get("font", ""),
                        "size": s.get("size", 0)
                    })
    return spans


def extract_metadata(page_text):
    """
    Extracts metadata from the page text (PWB code, Process code, Model, etc.)
    """
    meta = {
        "pwb_code": "",
        "process_code": "",
        "model": "",
        "issue_date": "",
        "wmcr_no": "",
        "solder_type": ""
    }
    
    # PWB Part Code
    for line in page_text.splitlines():
        line_clean = line.strip()
        if "QPWB" in line_clean:
            m = re.search(r'((?:PV-)?QPWBC[A-Z0-9]+)', line_clean)
            if m:
                meta["pwb_code"] = m.group(1).strip()
                break
                
    # Process Code
    m_proc = re.search(r'PROCESS\s*CODE\s*([A-Z0-9\-]+)', page_text, re.IGNORECASE)
    if m_proc:
        meta["process_code"] = m_proc.group(1).strip()
    else:
        for line in page_text.splitlines():
            line_clean = line.strip()
            m = re.search(r'\b(CHA[0-9]{4}[A-Z0-9\-]*|RAD[0-9]{4}[A-Z0-9\-]*|SCP[0-9]{4}[A-Z0-9\-]*|CHP[0-9]{4}[A-Z0-9\-]*)\b', line_clean)
            if m:
                meta["process_code"] = m.group(1).strip()
                break

    # Model
    m_model = re.search(r'MODEL\s*\n\s*([A-Z0-9\s\/\-_]+)', page_text)
    if m_model:
        meta["model"] = m_model.group(1).strip()
        
    # Date
    m_date = re.search(r'(\d{2}\.\d{2}\.\d{2,4}|\d{2}\s+[A-Z]{3}\s+\d{4})', page_text)
    if m_date:
        meta["issue_date"] = m_date.group(1).strip()
        
    # WMCR No
    m_wmcr = re.search(r'WMCR\s*NO[\s\n]*([0-9A-Z]+)', page_text)
    if m_wmcr:
        meta["wmcr_no"] = m_wmcr.group(1).strip()
        
    return meta


def cluster_by_y(spans, tolerance=3.0):
    """
    Groups a list of spans into horizontal rows by their y0 coordinate.
    """
    rows = {}
    for s in sorted(spans, key=lambda s: s["y0"]):
        y = s["y0"]
        matched_y = None
        for ry in rows:
            if abs(ry - y) <= tolerance:
                matched_y = ry
                break
        if matched_y is None:
            matched_y = y
            rows[matched_y] = []
        rows[matched_y].append(s)
    return rows


def parse_page_bom(doc, pno):
    """
    Parses BOM tables from a single page of the PDF manual.
    Returns a tuple: (page_metadata, list_of_items)
    """
    page = doc[pno]
    page_text = page.get_text()
    meta = extract_metadata(page_text)
    spans = get_visual_spans(page)
    
    # 1. Look for table headers
    header_keywords = {"NO", "No", "S.N.", "S.N", "Rating", "VALUE", "SNM", "Code", "Part", "No.", "QTY", "Qty", "REMARK", "Remark", "STN"}
    h_candidate_spans = [s for s in spans if s["text"] in header_keywords or any(k in s["text"] for k in ["SNM Code", "S&O Code", "Part No", "S.N / Rating", "S.N / VALUE"])]
    if not h_candidate_spans:
        return meta, []
        
    # Group header candidates by y
    h_groups = cluster_by_y(h_candidate_spans, tolerance=3.5)
    
    # Find valid header lines (must have NO/S.N and QTY/REMARK)
    valid_header_ys = []
    for y, s_list in h_groups.items():
        texts = [s["text"] for s in s_list]
        has_id = any(t in ["NO", "No", "S.N.", "S.N", "Part"] or "Part" in t for t in texts)
        has_qty_rem = any(t in ["QTY", "Qty", "REMARK", "Remark"] for t in texts)
        if has_id and has_qty_rem:
            valid_header_ys.append(y)
            
    if not valid_header_ys:
        return meta, []
        
    page_items = []
    
    for hy in sorted(valid_header_ys):
        # Header spans at this y
        h_line_spans = [s for s in spans if abs(s["y0"] - hy) <= 4.5]
        
        # Check if there are model variants or headers in row right above hy (within 20 pt)
        above_spans = [s for s in spans if (hy - 22.0) <= s["y0"] < (hy - 0.5)]
        
        # Detect sub-tables (e.g. Left sub-table, Right sub-table)
        no_headers = [s for s in h_line_spans if s["text"] in ["NO", "No"]]
        no_headers.sort(key=lambda s: s["x0"])
        
        subtables_bounds = []
        if len(no_headers) >= 2:
            for i, nh in enumerate(no_headers):
                min_x = nh["x0"] - 12.0
                if i + 1 < len(no_headers):
                    max_x = no_headers[i+1]["x0"] - 6.0
                else:
                    max_x = min_x + 360.0
                subtables_bounds.append((min_x, max_x))
        elif len(no_headers) == 1:
            nh = no_headers[0]
            min_x = nh["x0"] - 12.0
            subtables_bounds.append((min_x, min_x + 360.0))
        else:
            # Fallback: find S.N or Part No header
            sn_headers = [s for s in h_line_spans if "Part" in s["text"] or "S.N" in s["text"]]
            sn_headers.sort(key=lambda s: s["x0"])
            if sn_headers:
                min_x = sn_headers[0]["x0"] - 40.0
                subtables_bounds.append((min_x, min_x + 360.0))
            else:
                continue

        # Process each sub-table
        for min_x, max_x in subtables_bounds:
            sub_h_spans = [s for s in h_line_spans if min_x <= s["x0"] <= max_x]
            sub_above = [s for s in above_spans if min_x <= s["x0"] <= max_x]
            
            # Identify variant names (e.g. CHA3072AR-2, CHA3072AR-2A, CHP3280TRM-1, etc.)
            variant_names = []
            for s in sorted(sub_above, key=lambda s: s["x0"]):
                txt = s["text"].strip()
                if re.search(r'([A-Z0-9]{3,}\-[A-Z0-9]+)', txt):
                    if txt not in variant_names:
                        variant_names.append(txt)
            
            # If variant names not in row above, check if in process code
            if not variant_names:
                if meta["process_code"]:
                    variant_names = [meta["process_code"]]
                else:
                    variant_names = ["Standard"]
                    
            # Process Code detection: prioritize CHA / RAD / SCP / CHP codes with longest specific suffix
            found_proc_codes = []
            for v in variant_names:
                m_code = re.search(r'\b(CHA[0-9]{4}[A-Z0-9\-]*|RAD[0-9]{4}[A-Z0-9\-]*|SCP[0-9]{4}[A-Z0-9\-]*|CHP[0-9]{4}[A-Z0-9\-]*)\b', v)
                if m_code:
                    found_proc_codes.append(m_code.group(1).strip())
            
            page_proc_codes = re.findall(r'\b(CHA[0-9]{4}[A-Z0-9\-]*|RAD[0-9]{4}[A-Z0-9\-]*|SCP[0-9]{4}[A-Z0-9\-]*|CHP[0-9]{4}[A-Z0-9\-]*)\b', page_text)
            prefix = found_proc_codes[0][:3] if found_proc_codes else ""
            for pc in page_proc_codes:
                if prefix and pc.startswith(prefix) and pc not in found_proc_codes:
                    found_proc_codes.append(pc)
                    
            if found_proc_codes:
                proc_candidate = sorted(found_proc_codes, key=lambda x: (len(x), x), reverse=True)[0]
            elif meta["process_code"] and meta["process_code"] != meta["pwb_code"]:
                proc_candidate = meta["process_code"]
            elif variant_names:
                proc_candidate = variant_names[0]
            else:
                proc_candidate = "Standard"

            # Merge adjacent header spans on hy if they belong together (e.g. SNM + Code, S&O + Code, Part + No.)
            sub_h_sorted = sorted(sub_h_spans, key=lambda s: s["x0"])
            merged_h_spans = []
            for s in sub_h_sorted:
                if merged_h_spans and (s["x0"] - merged_h_spans[-1]["x1"]) < 8.0:
                    combined_txt = f"{merged_h_spans[-1]['text']} {s['text']}".strip()
                    if combined_txt.upper() in ["SNM CODE", "S&O CODE", "S & O CODE", "PART NO", "PART NO.", "PART NUMBER", "S.N / RATING", "S.N / VALUE", "INTERNAL S.N", "INTERNAL S.N."]:
                        merged_h_spans[-1]["text"] = combined_txt
                        merged_h_spans[-1]["x1"] = s["x1"]
                        continue
                merged_h_spans.append(dict(s))

            col_defs = []
            for s in merged_h_spans:
                role = classify_header_span(s["text"])
                if role:
                    col_defs.append({
                        "role": role,
                        "x0": s["x0"],
                        "x1": s["x1"],
                        "text": s["text"]
                    })

            left_cols = [c for c in col_defs if c["role"] not in ["QTY", "REMARK", "STN"]]
            qty_cols = [c for c in col_defs if c["role"] == "QTY"]
            qty_cols.sort(key=lambda s: s["x0"])
            rem_cols = [c for c in col_defs if c["role"] == "REMARK"]
            rem_cols.sort(key=lambda s: s["x0"])
            
            first_qty_x = qty_cols[0]["x0"] if qty_cols else (min_x + 135.0)

            col_bounds = []
            for i in range(len(left_cols) - 1):
                mid = (left_cols[i]["x1"] + left_cols[i+1]["x0"]) / 2.0
                col_bounds.append((left_cols[i]["role"], left_cols[i+1]["role"], mid))
            
            # Align multiple QTY columns with variants
            variant_map = []
            num_v = max(len(variant_names), len(qty_cols), 1)
            for vi in range(num_v):
                vname = variant_names[vi] if vi < len(variant_names) else (variant_names[-1] if variant_names else f"Variant_{vi+1}")
                q_x = qty_cols[vi]["x0"] if vi < len(qty_cols) else None
                r_x = rem_cols[vi]["x0"] if vi < len(rem_cols) else None
                variant_map.append({
                    "variant_name": vname,
                    "qty_x": q_x,
                    "rem_x": r_x
                })
                
            # Identify data spans below header
            data_spans = [s for s in spans if (hy + 1.5) <= s["y0"] <= (hy + 480.0) and min_x <= s["x0"] <= max_x]
            if not data_spans:
                continue
                
            data_rows = cluster_by_y(data_spans, tolerance=2.6)
            
            current_item = None
            last_valid_y = hy
            
            stop_keywords = [
                "TOTAL QTY", "GRAND TOTAL", "CAUTION", "STICK LABEL", "INSPECTION",
                "CO-ORDINATE", "MACHINE", "NC-PROGRAM", "UPDATE", "CHECK UPDATE",
                "SIDE VIEW", "STENCIL", "WMCR", "DELTA", "DIP A-SIDE", "REFLOW PROCESS"
            ]
            
            for ry in sorted(data_rows.keys()):
                r_spans = sorted(data_rows[ry], key=lambda s: s["x0"])
                row_text = " ".join([s["text"] for s in r_spans])
                
                # Check for footer or non-table section
                if any(k in row_text.upper() for k in stop_keywords):
                    break
                    
                # If gap between rows is too large (> 18 pt), table has ended
                if (ry - last_valid_y) > 18.0 and not (first_span := r_spans[0]) or ((ry - last_valid_y) > 18.0 and not re.match(r'^\d+$', r_spans[0]["text"])):
                    break
                    
                # Check if this row begins a new item (has a number in NO column)
                first_span = r_spans[0]
                is_item_start = False
                item_number = None
                
                if first_span["x0"] < min_x + 35.0 and re.match(r'^\d+$', first_span["text"]):
                    is_item_start = True
                    item_number = int(first_span["text"])
                    
                if is_item_start:
                    if current_item:
                        p, r = sanitize_part_and_rating(current_item["part_no"], current_item["rating"])
                        current_item["part_no"] = p
                        current_item["rating"] = r
                        page_items.append(current_item)
                        
                    current_item = {
                        "page": pno + 1,
                        "pwb_code": meta["pwb_code"],
                        "process_code": proc_candidate,
                        "model": meta["model"],
                        "item_no": item_number,
                        "sn": "",
                        "part_no": "",
                        "rating": "",
                        "alternates": [],
                        "variants_data": {},
                        "raw_lines": [row_text]
                    }
                    last_valid_y = ry
                    
                    for v in variant_map:
                        current_item["variants_data"][v["variant_name"]] = {
                            "qty": 0,
                            "remarks": []
                        }
                    
                    # Parse cells on this initial row
                    remaining_spans = r_spans[1:]
                    parse_row_cells(remaining_spans, current_item, min_x, max_x, variant_map, col_bounds, left_cols, first_qty_x)
                else:
                    # Continuation line for current item
                    if current_item:
                        last_valid_y = ry
                        current_item["raw_lines"].append(row_text)
                        # Check if this is an alternate part definition
                        if " or" in row_text or "or " in row_text or any(re.match(r'^[0-9]{8,}[A-Z0-9]*', s["text"]) for s in r_spans):
                            alt_text = row_text.strip()
                            if alt_text not in current_item["alternates"]:
                                current_item["alternates"].append(alt_text)
                        # Also check if there are remarks or qtys on this continuation line
                        parse_row_cells(r_spans, current_item, min_x, max_x, variant_map, col_bounds, left_cols, first_qty_x, is_continuation=True)
                        
            if current_item:
                p, r = sanitize_part_and_rating(current_item["part_no"], current_item["rating"])
                current_item["part_no"] = p
                current_item["rating"] = r
                page_items.append(current_item)
                
    return meta, page_items


def parse_row_cells(spans, item, min_x, max_x, variant_map, col_bounds, left_cols, first_qty_x, is_continuation=False):
    """
    Assigns individual text spans to fields: S.N, Part No/SNM Code, Rating, and Variant Qty/Remarks.
    Uses column bounds determined from table header to prevent mixing Rating with Part Number.
    """
    for s in spans:
        sx = (s["x0"] + s["x1"]) / 2.0
        txt = s["text"].strip()
        if not txt:
            continue
        
        # 1. Left side of table: S.N., Part No, Rating (sx < first_qty_x - 3.0)
        if sx < first_qty_x - 3.0:
            if is_continuation:
                # On continuation rows, left side is usually alternate part or wrapped rating
                continue
                
            if col_bounds and left_cols:
                # Find matching column bucket based on boundaries
                assigned_role = left_cols[-1]["role"]
                for i, (r1, r2, bmid) in enumerate(col_bounds):
                    if sx < bmid:
                        assigned_role = left_cols[i]["role"]
                        break
                
                if assigned_role == "NO":
                    continue
                elif assigned_role == "SN":
                    item["sn"] = f"{item['sn']} {txt}".strip() if item["sn"] else txt
                elif assigned_role in ["RATING", "VALUE"]:
                    item["rating"] = f"{item['rating']} {txt}".strip() if item["rating"] else txt
                elif assigned_role in ["CODE", "PART"]:
                    # Never append lone operator / polarity symbols to part number
                    if txt not in ["+", "-", "or", "C"]:
                        item["part_no"] = f"{item['part_no']} {txt}".strip() if item["part_no"] else txt
                elif assigned_role == "SN_RATING":
                    # Combined column (e.g. media_488: S.N / Rating)
                    if re.match(r'^\d{7,}[A-Z0-9\-]*$', txt):
                        item["sn"] = f"{item['sn']} {txt}".strip() if item["sn"] else txt
                    else:
                        item["rating"] = f"{item['rating']} {txt}".strip() if item["rating"] else txt
            else:
                # Heuristic fallback if header columns could not be defined
                if is_rating_string(txt):
                    item["rating"] = f"{item['rating']} {txt}".strip() if item["rating"] else txt
                elif re.match(r'^\d{7,}[A-Z0-9\-]*$', txt):
                    item["sn"] = f"{item['sn']} {txt}".strip() if item["sn"] else txt
                elif is_sharp_part_code(txt) or re.match(r'^[VQRBL][A-Z0-9\+\-\/]{8,}', txt):
                    if not item["part_no"]:
                        item["part_no"] = txt
                    else:
                        # If current part_no is actually a rating, move it to rating
                        if is_rating_string(item["part_no"]):
                            item["rating"] = f"{item['part_no']} {item['rating']}".strip()
                            item["part_no"] = txt
                        else:
                            item["part_no"] = f"{item['part_no']} {txt}".strip()
                else:
                    item["rating"] = f"{item['rating']} {txt}".strip() if item["rating"] else txt
        else:
            # 2. Right side of table: Variant QTY and REMARK columns
            matched_v = None
            if len(variant_map) == 1:
                matched_v = variant_map[0]
            elif len(variant_map) >= 2:
                v1 = variant_map[0]
                v2 = variant_map[1]
                v2_qx = v2["qty_x"] if v2["qty_x"] else (first_qty_x + 60.0)
                
                if sx < v2_qx - 5.0:
                    matched_v = v1
                else:
                    matched_v = v2
                    
            if matched_v:
                vdata = item["variants_data"][matched_v["variant_name"]]
                q_x = matched_v["qty_x"]
                is_qty = False
                if q_x and abs(sx - q_x) <= 12.0 and re.match(r'^\d+$', txt):
                    if vdata["qty"] == 0:
                        vdata["qty"] = int(txt)
                        is_qty = True
                elif re.match(r'^\d+$', txt) and sx < (q_x + 8.0 if q_x else first_qty_x + 12.0) and vdata["qty"] == 0:
                    vdata["qty"] = int(txt)
                    is_qty = True
                    
                if not is_qty:
                    if txt not in ["+", "-"] and txt not in vdata["remarks"]:
                        vdata["remarks"].append(txt)


def extract_full_bom(pdf_path):
    """
    Extracts the complete BOM from all pages of a PDF manual.
    Returns:
      - doc_info: dict with summary metadata
      - consolidated_records: list of dicts (flat records ready for Excel / UI)
      - matrix_records: list of dicts (matrix view with side-by-side variants)
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"File not found: {pdf_path}")
        
    doc = pymupdf.open(pdf_path)
    filename = os.path.basename(pdf_path)
    
    all_items = []
    doc_metadata = {
        "filename": filename,
        "filepath": pdf_path,
        "total_pages": len(doc),
        "pwb_code": "",
        "model": "",
        "processes": set(),
        "total_items": 0,
        "total_parts_count": 0
    }
    
    for pno in range(len(doc)):
        page_meta, items = parse_page_bom(doc, pno)
        if page_meta["pwb_code"] and not doc_metadata["pwb_code"]:
            doc_metadata["pwb_code"] = page_meta["pwb_code"]
        if page_meta["model"] and not doc_metadata["model"]:
            doc_metadata["model"] = page_meta["model"]
            
        all_items.extend(items)
        
    # Collect all unique processes from extracted items
    item_procs = {item["process_code"] for item in all_items if item["process_code"]}
    doc_metadata["processes"] = sorted(list(item_procs)) if item_procs else []
    doc_metadata["total_items"] = len(all_items)
    
    # Generate flat consolidated records and matrix records
    consolidated_records = []
    matrix_records = []
    
    seq_no = 1
    for item in all_items:
        # PWB and process
        pwb = item["pwb_code"] or doc_metadata["pwb_code"]
        proc = item["process_code"]
        part_no, rating = sanitize_part_and_rating(item["part_no"], item["rating"])
        sn = item["sn"]
        alternates = " | ".join(item["alternates"]) if item["alternates"] else ""
        
        # Matrix record row
        mat_row = {
            "No": item["item_no"],
            "Page": item["page"],
            "Process": proc,
            "PWB_Code": pwb,
            "Part_No": part_no or sn,
            "SNM_Code": part_no if part_no else "",
            "Internal_SN": sn,
            "Rating": rating,
            "Alternate_Parts": alternates
        }
        
        for vname, vdata in item["variants_data"].items():
            qty = vdata["qty"]
            # If qty is 0 but remarks exist, qty is count of remarks
            if qty == 0 and vdata["remarks"]:
                qty = len(vdata["remarks"])
            rem_str = " ".join(vdata["remarks"])
            
            mat_row[f"Qty_{vname}"] = qty
            mat_row[f"Remarks_{vname}"] = rem_str
            
            # Add to flat consolidated records if qty > 0 or remarks exist
            if qty > 0 or rem_str:
                consolidated_records.append({
                    "Seq": seq_no,
                    "Item_No": item["item_no"],
                    "Process": proc,
                    "Model_Variant": vname,
                    "PWB_Code": pwb,
                    "Part_No": part_no or sn,
                    "SNM_Code": part_no,
                    "Internal_SN": sn,
                    "Rating_Spec": rating,
                    "Qty": qty,
                    "Remarks_Locations": rem_str,
                    "Alternate_Parts": alternates,
                    "Source_Page": item["page"],
                    "Filename": filename
                })
                seq_no += 1
                doc_metadata["total_parts_count"] += qty
                
        matrix_records.append(mat_row)
        
    return doc_metadata, consolidated_records, matrix_records


if __name__ == "__main__":
    import glob
    for f in sorted(glob.glob("sample_manuals/*.pdf")):
        meta, flat, mat = extract_full_bom(f)
        print(f"=== {meta['filename']} ===")
        print(f"  PWB Code: {meta['pwb_code']}, Processes: {meta['processes']}")
        print(f"  Total Items: {meta['total_items']}, Total Flat BOM Rows: {len(flat)}, Total Parts Qty: {meta['total_parts_count']}")
        if flat:
            print("  Sample row 1:", flat[0])
