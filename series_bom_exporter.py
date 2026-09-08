"""
Series BOM Exporter
Generates professional Excel workbooks for Series BOM Comparisons and QC FAI Checklists.
Includes:
- Sheet 1: "🎯 BIÊN BẢN KIỂM TRA FAI (QC)" (Focus Checklist only with actions & sign-off)
- Sheet 2: "📊 TOÀN BỘ ĐỐI CHIẾU (MATRIX)" (Full side-by-side component list)
"""

import os
from datetime import datetime
from typing import Dict, Any, List
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Design tokens
FONT_NAME = "Segoe UI"
COLOR_NAVY_TITLE = "0D1B2A"
COLOR_HEADER_BG = "1B4F72"
COLOR_HEADER_FG = "FFFFFF"
COLOR_BORDER = "C5CDD6"
COLOR_ZEBRA = "F7F9FB"

# Status colors
BG_ADDED = "E8F8F5"       # Light mint green
FG_ADDED = "0E6251"
BG_REMOVED = "FDEDEC"     # Light pink red
FG_REMOVED = "78281F"
BG_MODIFIED = "FEF9E7"    # Light amber yellow
FG_MODIFIED = "7D6608"
BG_MATCHED = "FFFFFF"
FG_MATCHED = "2C3E50"


def _border(top="thin", bottom="thin", left="thin", right="thin", color=COLOR_BORDER):
    return Border(
        top=Side(style=top, color=color) if top else None,
        bottom=Side(style=bottom, color=color) if bottom else None,
        left=Side(style=left, color=color) if left else None,
        right=Side(style=right, color=color) if right else None,
    )


def export_series_bom_report(comparison_result: Dict[str, Any], output_path: str) -> str:
    """
    Exports a comprehensive Excel report with QC FAI Checklist and Full Delta Matrix.
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    summary = comparison_result.get("summary", {})
    qc_focus = comparison_result.get("qc_focus_items", [])
    all_items = comparison_result.get("all_items", [])
    
    model_name = summary.get("model_b") or summary.get("model_a") or "N/A"
    series_a = summary.get("series_a") or "A"
    series_b = summary.get("series_b") or "B"
    
    # ── SHEET 1: BIÊN BẢN KIỂM TRA FAI (QC FOCUS CHECKLIST) ────────────────────
    ws_fai = wb.create_sheet(title="BIEN BAN FAI (QC)")
    ws_fai.views.sheetView[0].showGridLines = True
    
    # Row 1: Title banner
    ws_fai.merge_cells("A1:H1")
    cell_title = ws_fai["A1"]
    cell_title.value = "BIÊN BẢN KIỂM TRA ĐỔI MODEL / SERIES (FIRST ARTICLE INSPECTION - FAI)"
    cell_title.font = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
    cell_title.fill = PatternFill(start_color=COLOR_NAVY_TITLE, end_color=COLOR_NAVY_TITLE, fill_type="solid")
    cell_title.alignment = Alignment(horizontal="center", vertical="center")
    ws_fai.row_dimensions[1].height = 36
    
    # Row 2: Subtitle
    ws_fai.merge_cells("A2:H2")
    cell_sub = ws_fai["A2"]
    cell_sub.value = "HỆ THỐNG VIPQC AI — DANH SÁCH LINH KIỆN CẦN CHÚ Ý KIỂM SOÁT (CHỈ KIỂM LINH KIỆN NÀY)"
    cell_sub.font = Font(name=FONT_NAME, size=10, italic=True, color="FFFFFF")
    cell_sub.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    cell_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws_fai.row_dimensions[2].height = 20
    
    # Row 4-6: Metadata block
    ws_fai.row_dimensions[4].height = 22
    ws_fai.row_dimensions[5].height = 22
    ws_fai.row_dimensions[6].height = 22
    
    meta_info = [
        ("A4", "B4", "Mã Bo Mạch / Model:", f"{model_name}"),
        ("D4", "E4", "Chuyển Đổi Series:", f"Series {series_a} ➔ Series {series_b}"),
        ("G4", "H4", "Ngày Lập:", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("A5", "B5", "File BOM Gốc (A):", summary.get("file_a", "")),
        ("D5", "E5", "File BOM Mới (B):", summary.get("file_b", "")),
        ("G5", "H5", "Tỷ Lệ Giống Nhau:", f"{summary.get('match_percentage', 0)}% (Dùng chung)"),
        ("A6", "B6", "Tổng Vị Trí Bo Mạch:", f"{summary.get('total_locations', 0)} vị trí"),
        ("D6", "E6", "Tổng LK Cần Kiểm (QC Focus):", f"{summary.get('focus_count', 0)} vị trí (Cần kiểm 100%)"),
        ("G6", "H6", "Trạng Thái Thẩm Định:", summary.get("validation_status", "VALID"))
    ]
    
    for lbl_cell, val_cell, label, val in meta_info:
        ws_fai[lbl_cell].value = label
        ws_fai[lbl_cell].font = Font(name=FONT_NAME, size=9, bold=True, color="4A5568")
        ws_fai[lbl_cell].alignment = Alignment(horizontal="right", vertical="center")
        
        ws_fai[val_cell].value = val
        ws_fai[val_cell].font = Font(name=FONT_NAME, size=9, bold=True, color="1A202C")
        ws_fai[val_cell].alignment = Alignment(horizontal="left", vertical="center")

    # Row 8: KPI Stats summary cards
    ws_fai.row_dimensions[8].height = 28
    kpis = [
        ("A8", "B8", f"🟢 THÊM MỚI: {summary.get('added_count', 0)}", BG_ADDED, FG_ADDED),
        ("C8", "D8", f"🔴 BỎ TRỐNG (DNP): {summary.get('removed_count', 0)}", BG_REMOVED, FG_REMOVED),
        ("E8", "F8", f"🟡 ĐỔI MÃ: {summary.get('modified_count', 0)}", BG_MODIFIED, FG_MODIFIED),
        ("G8", "H8", f"⚪ DÙNG CHUNG: {summary.get('matched_count', 0)}", "F0F3F4", "566573"),
    ]
    for c1, c2, txt, bg, fg in kpis:
        ws_fai.merge_cells(f"{c1}:{c2}")
        cell = ws_fai[c1]
        cell.value = txt
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color=fg)
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _border()

    # Row 10: Table Header
    headers_fai = [
        ("STT", 6),
        ("Vị Trí (Ref Des)", 14),
        ("Phân Loại", 16),
        (f"Mã LK ({series_a})", 20),
        (f"Mã LK ({series_b})", 20),
        ("Chỉ Dẫn Hành Động Cụ Thể Cho QC (Action Guide)", 48),
        ("Kết Quả QC", 14),
        ("Ghi Chú", 18)
    ]
    
    ws_fai.row_dimensions[10].height = 28
    for col_idx, (h_name, h_width) in enumerate(headers_fai, 1):
        c = ws_fai.cell(row=10, column=col_idx, value=h_name)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=COLOR_HEADER_FG)
        c.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border(color="FFFFFF")
        ws_fai.column_dimensions[get_column_letter(col_idx)].width = h_width

    # Populate QC Focus items
    current_row = 11
    for idx, item in enumerate(qc_focus, 1):
        ws_fai.row_dimensions[current_row].height = 24
        
        status = item.get("status", "")
        if status == "ADDED":
            row_bg, row_fg = BG_ADDED, FG_ADDED
        elif status == "REMOVED":
            row_bg, row_fg = BG_REMOVED, FG_REMOVED
        elif status == "MODIFIED":
            row_bg, row_fg = BG_MODIFIED, FG_MODIFIED
        else:
            row_bg, row_fg = "FFFFFF", "212529"
            
        cells_data = [
            (idx, "center", False),
            (item.get("location", ""), "center", True),
            (item.get("status_vn", ""), "center", True),
            (item.get("part_a", ""), "left", False),
            (item.get("part_b", ""), "left", False),
            (item.get("action_guide", ""), "left", False),
            (item.get("qc_status", "PENDING"), "center", True),
            ("", "left", False)
        ]
        
        for col_idx, (val, align, bold) in enumerate(cells_data, 1):
            c = ws_fai.cell(row=current_row, column=col_idx, value=val)
            c.font = Font(name=FONT_NAME, size=9, bold=bold, color=row_fg if bold else "212529")
            c.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(col_idx == 6))
            c.border = _border()
            
        current_row += 1

    # Empty rows if no focus items
    if not qc_focus:
        ws_fai.merge_cells(f"A{current_row}:H{current_row}")
        c = ws_fai[f"A{current_row}"]
        c.value = "CHÚC MỪNG: Hai Series hoàn toàn trùng khớp 100%, không có linh kiện nào khác biệt!"
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="27AE60")
        c.alignment = Alignment(horizontal="center", vertical="center")
        current_row += 1

    # Sign-off box
    current_row += 2
    ws_fai.row_dimensions[current_row].height = 20
    ws_fai.merge_cells(f"A{current_row}:C{current_row}")
    ws_fai.merge_cells(f"D{current_row}:F{current_row}")
    ws_fai.merge_cells(f"G{current_row}:H{current_row}")
    
    ws_fai[f"A{current_row}"].value = "NGƯỜI KIỂM TRA QC (INSPECTOR)"
    ws_fai[f"D{current_row}"].value = "KỸ THUẬT SMT (SMT LEADER)"
    ws_fai[f"G{current_row}"].value = "TRƯỞNG PHÒNG QC (QC MANAGER)"
    
    for c_addr in [f"A{current_row}", f"D{current_row}", f"G{current_row}"]:
        ws_fai[c_addr].font = Font(name=FONT_NAME, size=9, bold=True, color="2C3E50")
        ws_fai[c_addr].alignment = Alignment(horizontal="center", vertical="center")
        
    current_row += 1
    ws_fai.row_dimensions[current_row].height = 45  # Space for signature
    ws_fai.merge_cells(f"A{current_row}:C{current_row}")
    ws_fai.merge_cells(f"D{current_row}:F{current_row}")
    ws_fai.merge_cells(f"G{current_row}:H{current_row}")
    
    for col_l in ["A", "D", "G"]:
        ws_fai[f"{col_l}{current_row}"].value = "(Ký và ghi rõ họ tên)"
        ws_fai[f"{col_l}{current_row}"].font = Font(name=FONT_NAME, size=8, italic=True, color="95A5A6")
        ws_fai[f"{col_l}{current_row}"].alignment = Alignment(horizontal="center", vertical="bottom")

    # ── SHEET 2: TOÀN BỘ ĐỐI CHIẾU (FULL MATRIX) ──────────────────────────────
    ws_matrix = wb.create_sheet(title="TOAN BO DOI CHIEU (MATRIX)")
    ws_matrix.views.sheetView[0].showGridLines = True
    
    headers_matrix = [
        ("STT", 6),
        ("Vị Trí (Ref Des)", 14),
        ("Trạng Thái", 16),
        (f"Mã LK ({series_a})", 20),
        (f"SL ({series_a})", 8),
        (f"Mã LK ({series_b})", 20),
        (f"SL ({series_b})", 8),
        (f"Thông Số Kỹ Thuật ({series_b})", 45),
        ("Ghi Chú", 20)
    ]
    
    ws_matrix.row_dimensions[1].height = 28
    for col_idx, (h_name, h_width) in enumerate(headers_matrix, 1):
        c = ws_matrix.cell(row=1, column=col_idx, value=h_name)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=COLOR_HEADER_FG)
        c.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _border(color="FFFFFF")
        ws_matrix.column_dimensions[get_column_letter(col_idx)].width = h_width

    for r_idx, item in enumerate(all_items, 2):
        ws_matrix.row_dimensions[r_idx].height = 20
        status = item.get("status", "")
        
        if status == "ADDED":
            row_bg, row_fg = BG_ADDED, FG_ADDED
        elif status == "REMOVED":
            row_bg, row_fg = BG_REMOVED, FG_REMOVED
        elif status == "MODIFIED":
            row_bg, row_fg = BG_MODIFIED, FG_MODIFIED
        else:
            row_bg = COLOR_ZEBRA if (r_idx % 2 == 0) else "FFFFFF"
            row_fg = "212529"
            
        m_cells = [
            (r_idx - 1, "center", False),
            (item.get("location", ""), "center", (status != "MATCHED")),
            (item.get("status_vn", ""), "center", (status != "MATCHED")),
            (item.get("part_a", ""), "left", False),
            (item.get("qty_a", ""), "center", False),
            (item.get("part_b", ""), "left", False),
            (item.get("qty_b", ""), "center", False),
            (item.get("spec_b", "") or item.get("spec_a", ""), "left", False),
            ("Dùng chung" if status == "MATCHED" else "Cần chú ý", "center", False)
        ]
        
        for col_idx, (val, align, bold) in enumerate(m_cells, 1):
            c = ws_matrix.cell(row=r_idx, column=col_idx, value=val)
            c.font = Font(name=FONT_NAME, size=9, bold=bold, color=row_fg if bold else "212529")
            c.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
            c.alignment = Alignment(horizontal=align, vertical="center")
            c.border = _border()

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    return output_path
