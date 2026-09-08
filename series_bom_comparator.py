"""
Series BOM Comparator Engine
Compares multi-level BOMs (Bills of Materials) across different series of the same model.
Supports ERP Multi-level BOM PDF printouts (多階材料用量清單列印) and Excel BOMs (.xlsx, .xls).
Identifies critical differences (Added, Removed/DNP, Modified) and produces a QC Focus Checklist.
"""

import os
import sys
import re
from typing import Dict, List, Tuple, Any, Optional

# Ensure Windows DLL search directories for PyMuPDF
if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

import pymupdf


def extract_model_and_series(text_or_filename: str) -> Tuple[str, str]:
    """
    Extracts base model and series from part number, header, or filename.
    Examples:
      - 'CHA3259AF-1A' -> Model: 'CHA3259AF', Series: '1A'
      - 'CHA3259AF-1D' -> Model: 'CHA3259AF', Series: '1D'
      - '3275-2A' -> Model: '3275-2', Series: '2A' or 'A'
    """
    clean = text_or_filename.strip()
    
    # Check for pattern like CHA3259AF-1A or 3275-2A
    m = re.search(r"([A-Za-z0-9_]+)-([0-9]+[A-Za-z0-9]*)", clean)
    if m:
        return m.group(1), m.group(2)
        
    m2 = re.search(r"([A-Za-z0-9_]+)\s*([A-Za-z0-9]+)$", clean)
    if m2:
        return m2.group(1), m2.group(2)
        
    return clean, ""


def parse_erp_bom_pdf(pdf_path: str) -> Dict[str, Any]:
    """
    Parses ERP Multi-level BOM PDF report (e.g. 多階材料用量清單列印).
    Extracts Header metadata (Main Part No, Model, Series, Report Date)
    and all component items with their Ref Des locations.
    """
    doc = pymupdf.open(pdf_path)
    main_part_no = ""
    report_title = ""
    company_name = ""
    records = []
    
    for page_idx, page in enumerate(doc):
        text = page.get_text()
        
        if not company_name:
            if "COMPANY LIMITED" in text or "ELECTRONICS" in text:
                for line in text.split("\n")[:5]:
                    if "COMPANY" in line.upper() or "TECHNOLOGIES" in line.upper():
                        company_name = line.strip()
                        break
                        
        if not main_part_no:
            m = re.search(r"主件料號:\s*([A-Za-z0-9\-_]+)", text)
            if m:
                main_part_no = m.group(1).strip()
                
        if not report_title:
            if "多階材料用量清單" in text:
                report_title = "多階材料用量清單"
            elif "BOM" in text:
                report_title = "BOM List"
                
        words = page.get_text("words")
        # Body words between header (y ~ 155) and footer (y ~ 1075)
        table_words = [w for w in words if 155 <= w[1] <= 1075]
        
        # Identify item rows: numbers in x0 between 50 and 85
        item_words = [w for w in table_words if 50 <= w[0] <= 85 and re.match(r"^\d+$", w[4])]
        item_words.sort(key=lambda w: w[1])
        
        for i, item_w in enumerate(item_words):
            y_start = item_w[1] - 4
            y_end = item_words[i+1][1] - 4 if i + 1 < len(item_words) else 1075
            
            row_words = [w for w in table_words if y_start <= w[1] < y_end]
            
            level = " ".join([w[4] for w in row_words if w[0] < 50]).strip()
            item_no = item_w[4]
            part_no = " ".join([w[4] for w in row_words if 95 <= w[0] <= 220]).strip()
            old_part = " ".join([w[4] for w in row_words if 220 < w[0] <= 280]).strip()
            spec = " ".join([w[4] for w in row_words if 280 < w[0] <= 490]).strip()
            src = " ".join([w[4] for w in row_words if 490 < w[0] <= 525]).strip()
            unit = " ".join([w[4] for w in row_words if 525 < w[0] <= 565]).strip()
            qty = " ".join([w[4] for w in row_words if 565 < w[0] <= 620]).strip()
            loss = " ".join([w[4] for w in row_words if 620 < w[0] <= 670]).strip()
            
            loc_words = [w[4] for w in row_words if w[0] >= 670]
            loc_str = "".join(loc_words)
            locations = [loc.strip() for loc in re.findall(r"[A-Za-z0-9\-_]+", loc_str) if loc.strip()]
            
            if part_no:
                records.append({
                    "item_no": item_no,
                    "level": level,
                    "part_no": part_no,
                    "old_part": old_part,
                    "spec": spec,
                    "src": src,
                    "unit": unit,
                    "qty": qty,
                    "loss": loss,
                    "locations": locations,
                    "page": page_idx + 1
                })
                
    base_model, series = extract_model_and_series(main_part_no or os.path.basename(pdf_path))
    
    return {
        "file_path": pdf_path,
        "file_name": os.path.basename(pdf_path),
        "file_type": "pdf_erp",
        "company_name": company_name,
        "report_title": report_title or "ERP Multi-Level BOM",
        "main_part_no": main_part_no,
        "base_model": base_model,
        "series": series,
        "total_pages": len(doc),
        "records": records
    }


def parse_excel_bom(excel_path: str) -> Dict[str, Any]:
    """
    Parses an Excel BOM file (.xlsx / .xls).
    Attempts to identify columns for Part Number, Locations / Ref Des, Qty, Description / Spec.
    """
    import openpyxl
    
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {
            "file_path": excel_path,
            "file_name": os.path.basename(excel_path),
            "file_type": "excel",
            "base_model": "",
            "series": "",
            "main_part_no": "",
            "records": []
        }
        
    # Find header row
    header_idx = -1
    col_map = {}
    for r_idx, row in enumerate(rows[:20]):
        row_str = [str(c).strip().upper() if c is not None else "" for c in row]
        # Check for key header names
        has_part = any("PART" in c or "MÃ" in c or "料號" in c for c in row_str)
        has_loc = any("LOC" in c or "REF" in c or "VỊ TRÍ" in c or "位置" in c for c in row_str)
        if has_part or has_loc:
            header_idx = r_idx
            for c_idx, val in enumerate(row_str):
                if any(k in val for k in ["PART NO", "PART NUMBER", "MATERIAL", "MÃ LK", "MÃ VẬT TƯ", "元件料號"]):
                    col_map["part_no"] = c_idx
                elif any(k in val for k in ["LOCATION", "REF", "VỊ TRÍ", "插件位置", "DESIGNATOR"]):
                    col_map["locations"] = c_idx
                elif any(k in val for k in ["QTY", "QUANTITY", "SL", "SỐ LƯỢNG", "用量"]):
                    col_map["qty"] = c_idx
                elif any(k in val for k in ["DESC", "SPEC", "SPECIFICATION", "THÔNG SỐ", "规格", "品名"]):
                    col_map["spec"] = c_idx
                elif any(k in val for k in ["NO", "STT", "項次"]):
                    col_map["item_no"] = c_idx
            break
            
    records = []
    if header_idx != -1 and "part_no" in col_map:
        for r_idx in range(header_idx + 1, len(rows)):
            row = rows[r_idx]
            part_val = str(row[col_map["part_no"]]).strip() if col_map.get("part_no") is not None and row[col_map["part_no"]] is not None else ""
            if not part_val or part_val.lower() == "none":
                continue
                
            loc_val = str(row[col_map["locations"]]).strip() if col_map.get("locations") is not None and row[col_map["locations"]] is not None else ""
            locations = [l.strip() for l in re.findall(r"[A-Za-z0-9\-_]+", loc_val) if l.strip()]
            
            qty_val = str(row[col_map["qty"]]).strip() if col_map.get("qty") is not None and row[col_map["qty"]] is not None else "1"
            spec_val = str(row[col_map["spec"]]).strip() if col_map.get("spec") is not None and row[col_map["spec"]] is not None else ""
            item_no = str(row[col_map["item_no"]]).strip() if col_map.get("item_no") is not None and row[col_map["item_no"]] is not None else str(len(records) + 1)
            
            records.append({
                "item_no": item_no,
                "level": "1",
                "part_no": part_val,
                "old_part": "",
                "spec": spec_val,
                "src": "",
                "unit": "PCS",
                "qty": qty_val,
                "loss": "",
                "locations": locations,
                "page": 1
            })

    base_model, series = extract_model_and_series(os.path.basename(excel_path))

    return {
        "file_path": excel_path,
        "file_name": os.path.basename(excel_path),
        "file_type": "excel",
        "company_name": "",
        "report_title": "Excel BOM",
        "main_part_no": os.path.basename(excel_path),
        "base_model": base_model,
        "series": series,
        "total_pages": 1,
        "records": records
    }


def parse_any_bom(file_path: str) -> Dict[str, Any]:
    """Unified entry point to parse either PDF or Excel BOM."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return parse_erp_bom_pdf(file_path)
    elif ext in [".xlsx", ".xls", ".xlsm"]:
        return parse_excel_bom(file_path)
    else:
        raise ValueError(f"Định dạng file không được hỗ trợ: {ext}. Vui lòng chọn file .pdf hoặc .xlsx!")


def compare_series_boms(bom_a_data: Dict[str, Any], bom_b_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compares two parsed BOM datasets (Series A as baseline, Series B as target).
    Calculates delta, categorizes into Added, Removed, Modified, Matched,
    and builds the QC Focus Checklist with specific action guides.
    """
    # 1. Map each location to its component record
    loc_map_a: Dict[str, Dict[str, Any]] = {}
    for r in bom_a_data.get("records", []):
        for loc in r.get("locations", []):
            loc_map_a[loc] = r
            
    loc_map_b: Dict[str, Dict[str, Any]] = {}
    for r in bom_b_data.get("records", []):
        for loc in r.get("locations", []):
            loc_map_b[loc] = r
            
    all_locations = sorted(list(set(loc_map_a.keys()) | set(loc_map_b.keys())))
    
    # 2. Categorization
    added_items: List[Dict[str, Any]] = []
    removed_items: List[Dict[str, Any]] = []
    modified_items: List[Dict[str, Any]] = []
    matched_items: List[Dict[str, Any]] = []
    all_comparison_rows: List[Dict[str, Any]] = []
    
    for loc in all_locations:
        rec_a = loc_map_a.get(loc)
        rec_b = loc_map_b.get(loc)
        
        part_a = rec_a["part_no"] if rec_a else ""
        spec_a = rec_a["spec"] if rec_a else ""
        qty_a = rec_a["qty"] if rec_a else "0"
        
        part_b = rec_b["part_no"] if rec_b else ""
        spec_b = rec_b["spec"] if rec_b else ""
        qty_b = rec_b["qty"] if rec_b else "0"
        
        if not rec_a and rec_b:
            # ADDED in Series B
            status = "ADDED"
            status_vn = "THÊM MỚI"
            badge_color = "#00e676"  # Bright green
            action_guide = f"Kiểm tra máy SMT ĐÃ GẮN linh kiện mã mới '{part_b}'. Soi hàn/giá trị đúng quy cách."
            row_data = {
                "location": loc,
                "status": status,
                "status_vn": status_vn,
                "badge_color": badge_color,
                "part_a": "(Trống - NC)",
                "spec_a": "",
                "qty_a": "0",
                "part_b": part_b,
                "spec_b": spec_b,
                "qty_b": qty_b,
                "action_guide": action_guide,
                "is_critical": True,
                "qc_status": "PENDING"  # PENDING, OK, NG
            }
            added_items.append(row_data)
            all_comparison_rows.append(row_data)
            
        elif rec_a and not rec_b:
            # REMOVED / DNP in Series B
            status = "REMOVED"
            status_vn = "BỎ TRỐNG (DNP)"
            badge_color = "#ff5252"  # Bright red
            action_guide = f"Kiểm tra vị trí {loc} PHẢI BỎ TRỐNG (DNP/NC). TUYỆT ĐỐI KHÔNG HÀN linh kiện cũ '{part_a}'."
            row_data = {
                "location": loc,
                "status": status,
                "status_vn": status_vn,
                "badge_color": badge_color,
                "part_a": part_a,
                "spec_a": spec_a,
                "qty_a": qty_a,
                "part_b": "(Bỏ trống - DNP)",
                "spec_b": "",
                "qty_b": "0",
                "action_guide": action_guide,
                "is_critical": True,
                "qc_status": "PENDING"
            }
            removed_items.append(row_data)
            all_comparison_rows.append(row_data)
            
        elif rec_a and rec_b and part_a != part_b:
            # MODIFIED part number at same location
            status = "MODIFIED"
            status_vn = "ĐỔI MÃ VẬT TƯ"
            badge_color = "#ffab00"  # Bright amber
            action_guide = f"ĐỔI MÃ: Kiểm tra SMT đã đổi Feeder từ '{part_a}' sang '{part_b}'. Đo LCR/Soi thông số."
            row_data = {
                "location": loc,
                "status": status,
                "status_vn": status_vn,
                "badge_color": badge_color,
                "part_a": part_a,
                "spec_a": spec_a,
                "qty_a": qty_a,
                "part_b": part_b,
                "spec_b": spec_b,
                "qty_b": qty_b,
                "action_guide": action_guide,
                "is_critical": True,
                "qc_status": "PENDING"
            }
            modified_items.append(row_data)
            all_comparison_rows.append(row_data)
            
        else:
            # MATCHED / COMMON
            status = "MATCHED"
            status_vn = "TRÙNG KHỚP"
            badge_color = "#9e9e9e"  # Gray
            action_guide = "Linh kiện dùng chung giữa 2 Series. Không cần kiểm tra lại khi đổi model."
            row_data = {
                "location": loc,
                "status": status,
                "status_vn": status_vn,
                "badge_color": badge_color,
                "part_a": part_a,
                "spec_a": spec_a,
                "qty_a": qty_a,
                "part_b": part_b,
                "spec_b": spec_b,
                "qty_b": qty_b,
                "action_guide": action_guide,
                "is_critical": False,
                "qc_status": "OK"
            }
            matched_items.append(row_data)
            all_comparison_rows.append(row_data)

    # 3. QC Focus List combines all critical items (Added + Removed + Modified)
    qc_focus_items = added_items + removed_items + modified_items
    
    # 4. Model validation check
    model_a = bom_a_data.get("base_model", "")
    model_b = bom_b_data.get("base_model", "")
    series_a = bom_a_data.get("series", "")
    series_b = bom_b_data.get("series", "")
    
    model_match = bool(model_a and model_b and (model_a.upper() in model_b.upper() or model_b.upper() in model_a.upper()))
    series_differ = (series_a != series_b) if (series_a and series_b) else True
    
    validation_status = "VALID"
    validation_msg = "Hai BOM cùng Model, khác Series hợp lệ."
    if not model_match and (model_a and model_b):
        validation_status = "WARNING_DIFF_MODEL"
        validation_msg = f"Cảnh báo: Hai BOM dường như khác Model ({model_a} vs {model_b}). Hãy chắc chắn bạn chọn đúng file!"
    elif not series_differ and (series_a and series_b):
        validation_status = "WARNING_SAME_SERIES"
        validation_msg = f"Cảnh báo: Cả hai file đều là Series '{series_a}'. Bạn đang so sánh 2 file cùng Series!"

    summary = {
        "model_a": model_a,
        "series_a": series_a,
        "main_part_a": bom_a_data.get("main_part_no", ""),
        "file_a": bom_a_data.get("file_name", ""),
        
        "model_b": model_b,
        "series_b": series_b,
        "main_part_b": bom_b_data.get("main_part_no", ""),
        "file_b": bom_b_data.get("file_name", ""),
        
        "validation_status": validation_status,
        "validation_msg": validation_msg,
        
        "total_locations": len(all_locations),
        "matched_count": len(matched_items),
        "added_count": len(added_items),
        "removed_count": len(removed_items),
        "modified_count": len(modified_items),
        "focus_count": len(qc_focus_items),
        "match_percentage": round((len(matched_items) / len(all_locations) * 100) if all_locations else 0, 1)
    }

    return {
        "summary": summary,
        "qc_focus_items": qc_focus_items,
        "added_items": added_items,
        "removed_items": removed_items,
        "modified_items": modified_items,
        "matched_items": matched_items,
        "all_items": all_comparison_rows
    }


# ─── DRAWING INTEGRATION & BORDER HIGHLIGHT ENGINE ───────────────────────────
from PIL import Image, ImageDraw, ImageFont


def get_drawing_pages_info(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Returns list of page info for a drawing PDF to populate page selection dropdown.
    """
    if not os.path.exists(pdf_path):
        return []
    doc = pymupdf.open(pdf_path)
    pages = []
    for idx, page in enumerate(doc):
        text = page.get_text()
        first_line = ""
        for line in text.split("\n"):
            line = line.strip()
            if line and len(line) > 3:
                first_line = line[:40]
                break
        label = f"Trang {idx + 1}"
        if first_line:
            label += f" — {first_line}"
        pages.append({
            "index": idx,
            "label": label,
            "width": page.rect.width,
            "height": page.rect.height,
            "rotation": page.rotation
        })
    return pages


def locate_focus_items_on_drawing(doc: pymupdf.Document, page_idx: int, focus_items: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """
    Locates exact bounding boxes for all QC Focus items on a drawing page (unrotated coordinates).
    Returns mapping: loc -> {rects: [(x0, y0, x1, y1)...], item: dict}
    """
    if page_idx < 0 or page_idx >= len(doc):
        return {}
    page = doc[page_idx]

    orig_rot = page.rotation
    page.set_rotation(0)

    # Pre-fetch all words on page
    words = page.get_text("words")  # (x0, y0, x1, y1, word, block_no, line_no, word_no)
    word_map = {}
    for w in words:
        clean_w = re.sub(r"[,\s;:\-_/]", "", w[4]).upper()
        if clean_w:
            word_map.setdefault(clean_w, []).append((w[0], w[1], w[2], w[3]))

    located = {}
    for item in focus_items:
        loc = item.get("location", "").strip()
        if not loc:
            continue
        clean_loc = re.sub(r"[,\s;:\-_/]", "", loc).upper()
        rects = word_map.get(clean_loc, [])
        if not rects:
            # Fallback to search_for
            raw_rects = page.search_for(loc)
            rects = [(r.x0, r.y0, r.x1, r.y1) for r in raw_rects]

        if rects:
            located[loc] = {
                "rects": rects,
                "item": item
            }

    page.set_rotation(orig_rot)
    return located


def _transform_rect(rect: Tuple[float, float, float, float], page_w: float, page_h: float, zoom: float, rotation: int) -> Tuple[int, int, int, int]:
    """
    Transforms 72 DPI unrotated PDF rect to zoomed & rotated pixel bounding box.
    """
    x0, y0, x1, y1 = rect
    z = zoom

    if rotation == 90:
        # Rotated 90 deg clockwise: new width is H, new height is W
        px0 = int((page_h - y1) * z)
        py0 = int(x0 * z)
        px1 = int((page_h - y0) * z)
        py1 = int(x1 * z)
    elif rotation == 180:
        # Rotated 180 deg
        px0 = int((page_w - x1) * z)
        py0 = int((page_h - y1) * z)
        px1 = int((page_w - x0) * z)
        py1 = int((page_h - y0) * z)
    elif rotation == 270:
        # Rotated 270 deg clockwise
        px0 = int(y0 * z)
        py0 = int((page_w - x1) * z)
        px1 = int(y1 * z)
        py1 = int((page_w - x0) * z)
    else:
        # No rotation (0 deg)
        px0 = int(x0 * z)
        py0 = int(y0 * z)
        px1 = int(x1 * z)
        py1 = int(y1 * z)

    return (min(px0, px1), min(py0, py1), max(px0, px1), max(py0, py1))


def render_annotated_drawing_page(
    pdf_path: str,
    page_idx: int,
    focus_items: List[Dict[str, Any]],
    active_loc: Optional[str] = None,
    zoom: float = 1.2,
    rotation: int = 0
) -> Tuple[Optional[Image.Image], Dict[str, Tuple[int, int, int, int]]]:
    """
    Renders the PCB drawing page with multi-color border highlights for all focus items.
    Highlights:
      - 🟢 Green for ADDED
      - 🔴 Red for REMOVED (DNP)
      - 🟡 Amber for MODIFIED
      - 🎯 Prominent Spotlight reticle for active_loc
    Returns: (PIL.Image, pixel_coords_map)
    """
    if not os.path.exists(pdf_path):
        return None, {}

    doc = pymupdf.open(pdf_path)
    if page_idx < 0 or page_idx >= len(doc):
        doc.close()
        return None, {}

    page = doc[page_idx]
    orig_rot = page.rotation
    page.set_rotation(0)

    page_w = page.rect.width
    page_h = page.rect.height

    # Render base image at zoom
    mat = pymupdf.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Apply image rotation if requested
    if rotation == 90:
        img = img.transpose(Image.Transpose.ROTATE_270)  # Clockwise 90
    elif rotation == 180:
        img = img.transpose(Image.Transpose.ROTATE_180)
    elif rotation == 270:
        img = img.transpose(Image.Transpose.ROTATE_90)   # Clockwise 270

    # Locate focus items
    located = locate_focus_items_on_drawing(doc, page_idx, focus_items)
    page.set_rotation(orig_rot)
    doc.close()

    draw = ImageDraw.Draw(img)
    pixel_coords_map = {}

    # Color palette
    color_map = {
        "ADDED": "#00E676",      # Vivid green
        "REMOVED": "#FF5252",    # Vivid red
        "MODIFIED": "#FFAB00",   # Amber
    }

    # First pass: render all background highlight boxes
    for loc, data in located.items():
        item = data["item"]
        status = item.get("status", "ADDED")
        border_color = color_map.get(status, "#00E676")
        is_active = (active_loc and active_loc.upper() == loc.upper())

        for rect in data["rects"]:
            px0, py0, px1, py1 = _transform_rect(rect, page_w, page_h, zoom, rotation)
            # Add padding
            pad = 6
            b_x0, b_y0, b_x1, b_y1 = px0 - pad, py0 - pad, px1 + pad, py1 + pad

            # Save in pixel map
            pixel_coords_map[loc] = (b_x0, b_y0, b_x1, b_y1)

            if not is_active:
                # Draw outer glow + solid box
                draw.rectangle([b_x0 - 2, b_y0 - 2, b_x1 + 2, b_y1 + 2], outline=border_color, width=3)
                draw.rectangle([b_x0 + 1, b_y0 + 1, b_x1 - 1, b_y1 - 1], outline="#FFFFFF", width=1)

                # Draw small badge tag
                badge_w = max(50, len(loc) * 8 + 12)
                badge_h = 16
                draw.rectangle([b_x0, b_y0 - badge_h, b_x0 + badge_w, b_y0], fill="#0D1B2A")
                draw.rectangle([b_x0, b_y0 - badge_h, b_x0 + badge_w, b_y0], outline=border_color, width=1)
                draw.text((b_x0 + 4, b_y0 - badge_h + 2), f"{loc}", fill=border_color)

    # Second pass: render active_loc spotlight with high prominence
    if active_loc and active_loc in pixel_coords_map:
        b_x0, b_y0, b_x1, b_y1 = pixel_coords_map[active_loc]
        act_item = located[active_loc]["item"]
        status = act_item.get("status", "ADDED")
        spot_color = color_map.get(status, "#00E676")

        # Large glowing halo box
        for offset, alpha_col in [(6, "#00F0FF"), (4, spot_color), (2, "#FFFFFF")]:
            draw.rectangle([b_x0 - offset, b_y0 - offset, b_x1 + offset, b_y1 + offset],
                           outline=alpha_col, width=2)

        # Crosshair corner brackets
        corner_len = 16
        # Top-left
        draw.line([b_x0 - 10, b_y0 - 10, b_x0 - 10 + corner_len, b_y0 - 10], fill="#00F0FF", width=3)
        draw.line([b_x0 - 10, b_y0 - 10, b_x0 - 10, b_y0 - 10 + corner_len], fill="#00F0FF", width=3)
        # Top-right
        draw.line([b_x1 + 10, b_y0 - 10, b_x1 + 10 - corner_len, b_y0 - 10], fill="#00F0FF", width=3)
        draw.line([b_x1 + 10, b_y0 - 10, b_x1 + 10, b_y0 - 10 + corner_len], fill="#00F0FF", width=3)
        # Bottom-left
        draw.line([b_x0 - 10, b_y1 + 10, b_x0 - 10 + corner_len, b_y1 + 10], fill="#00F0FF", width=3)
        draw.line([b_x0 - 10, b_y1 + 10, b_x0 - 10, b_y1 + 10 - corner_len], fill="#00F0FF", width=3)
        # Bottom-right
        draw.line([b_x1 + 10, b_y1 + 10, b_x1 + 10 - corner_len, b_y1 + 10], fill="#00F0FF", width=3)
        draw.line([b_x1 + 10, b_y1 + 10, b_x1 + 10, b_y1 + 10 - corner_len], fill="#00F0FF", width=3)

        # Callout banner
        banner_txt = f"📍 {active_loc} | {act_item.get('status_vn', status)}"
        draw.rectangle([b_x0 - 4, b_y0 - 26, b_x0 + 190, b_y0 - 4], fill="#0F172A")
        draw.rectangle([b_x0 - 4, b_y0 - 26, b_x0 + 190, b_y0 - 4], outline="#00F0FF", width=2)
        draw.text((b_x0 + 6, b_y0 - 22), banner_txt, fill="#00F0FF")

    return img, pixel_coords_map
