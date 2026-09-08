"""
VIPQC AI — Working Manual BOM Extraction & PCB Drawing Comparator Desktop App
Supports Light & Dark Modes, Multilingual UI (Tiếng Việt, 中文, English),
and automated high-accuracy BOM extraction into publication-ready Excel.
"""

import os
import sys
import colorsys
import math
import threading
import time
from datetime import datetime
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageTk
import webbrowser
import tempfile
import updater

# Ensure Windows C-runtime DLLs are available for PyMuPDF
if sys.platform == "win32":
    for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

try:
    import windnd
    HAS_WINDND = True
except Exception:
    HAS_WINDND = False

from bom_extractor import extract_full_bom
from excel_exporter import export_bom_to_excel, export_batch_bom_to_excel
from bom_comparator import (
    parse_excel_bom, get_wm_pages_info, extract_wm_page_items,
    compare_boms, export_comparison_excel, normalize_key
)
from model_comparator import (
    compare_model_manuals, render_component_spotlight,
    render_full_drawing_view, render_curtain_drawing_pair,
    get_drawing_pages_catalog, normalize_process_stage,
    export_model_comparison_excel, clear_curtain_cache,
    get_annotated_base_images, extract_model_info,
    validate_model_pair
)
import series_bom_comparator as sbc
import series_bom_exporter as sbe

def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ─── CustomTkinter Global Theme ───────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ─── Design Tokens (Light, Dark) ─────────────────────────────────────────────
BG_DEEP     = ("#F1F5F9", "#0B0F1A")   # Deepest background
BG_CARD     = ("#FFFFFF", "#131929")   # Card / panel background
BG_SURFACE  = ("#E2E8F0", "#1C2438")   # Surface elements
BG_HOVER    = ("#CBD5E1", "#232D45")   # Hover state
ACCENT_TEAL = ("#0D9488", "#00C9A7")   # Teal action/success accent
ACCENT_BLUE = ("#2563EB", "#3D8EFF")   # Blue interactive accent
ACCENT_AMBER= ("#D97706", "#FFB830")   # Warning / secondary accent
TEXT_PRIMARY= ("#0F172A", "#E8EDF5")   # Primary text
TEXT_MUTED  = ("#64748B", "#7A8BA6")   # Muted / secondary text
BORDER_CLR  = ("#CBD5E1", "#2A3650")   # Border outline

FONT_HERO   = ("Segoe UI", 22, "bold")
FONT_H1     = ("Segoe UI", 14, "bold")
FONT_H2     = ("Segoe UI", 12, "bold")
FONT_BODY   = ("Segoe UI", 11)
FONT_MONO   = ("Consolas", 10)
FONT_BADGE  = ("Segoe UI", 9, "bold")

# Process-code type → colour chip
PROC_COLORS = {
    "CHA": "#3D8EFF",   # blue – SMT Chip
    "SCP": "#00C9A7",   # teal – Single/DIP
    "RAD": "#FF8C42",   # orange – Radial
    "CHP": "#A855F7",   # purple – CHP
}


def proc_color(code: str) -> str:
    return PROC_COLORS.get(code[:3].upper() if code else "", "#7A8BA6")


# Max allowed file size (10 MB) to prevent lag and optimize system performance
MAX_FILE_SIZE_MB = 10
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024


def format_file_size(size_bytes: int) -> str:
    """Formats file size into human-readable B, KB, MB."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


# ─── Multilingual Localization Dictionary ─────────────────────────────────────
I18N = {
    "vi": {
        "lang_display": "Tiếng Việt",
        "app_title": "⚡ VIPQC AI",
        "app_subtitle": "Working Manual  →  Excel   |   Hệ Thống Trí Tuệ Nhân Tạo & So Sánh Bản Vẽ",
        "badge_files": "FILE ĐÃ CHỌN",
        "badge_items": "MỤC LINH KIỆN",
        "badge_qty": "TỔNG SỐ LƯỢNG",
        "sec_queue": "DANH SÁCH FILE",
        "btn_add_files": "+ Thêm File",
        "btn_add_folder": "📁 Thư Mục",
        "empty_files": "Chưa có file nào được chọn (Tối đa 10MB/file).\nKéo thả file vào đây hoặc nhấn '+ Thêm File'.",
        "btn_clear_all": "✕ Xóa tất cả",
        "sec_output": "CẤU HÌNH XUẤT FILE",
        "switch_individual": "Xuất Excel riêng mỗi PDF",
        "switch_batch": "Tạo file Master tổng hợp (Batch)",
        "btn_run": "⚡  TRÍCH XUẤT & XUẤT EXCEL",
        "btn_running": "⏳  ĐANG XỬ LÝ…",
        "btn_open_excel": "📊 Mở Excel",
        "btn_open_folder": "📂 Thư Mục",
        "sec_preview": "XEM TRƯỚC DỮ LIỆU",
        "search_placeholder": "🔍  Tìm kiếm mã linh kiện, vị trí, quy trình…",
        "filter_all": "Tất cả",
        "sec_log": "NHẬT KÝ XỬ LÝ",
        "status_ready": "Sẵn sàng. Kéo thả hoặc chọn file Working Manual PDF để bắt đầu.",
        "status_processing": "Đang xử lý [{i}/{n}]: {fname}",
        "status_done": "✅ Hoàn thành! {count} bản ghi BOM từ {n} file.",
        "log_ready": "Ứng dụng sẵn sàng. Hãy kéo thả hoặc thêm file PDF và nhấn ⚡ Trích Xuất.",
        "log_start": "🚀 BẮT ĐẦU XỬ LÝ {n} FILE…",
        "log_analyzing": "[{i}/{n}] Phân tích '{fname}'…",
        "log_parsed": "   ✓ PWB={pwb}  Quy trình={procs}  Items={items}  Qty={qty}",
        "log_exported": "   📊 Đã xuất: {name}",
        "log_error": "   ❌ Lỗi '{fname}': {err}",
        "log_master": "🏆 MASTER EXCEL: {name}",
        "log_finish": "🎉 HOÀN THÀNH  —  {rows} dòng BOM  |  Tổng Qty: {qty}",
        "log_added_files": "Đã thêm {count} file từ '{folder}'.",
        "msg_no_files_title": "Chưa có file",
        "msg_no_files_body": "Vui lòng thêm ít nhất 1 file PDF!",
        "msg_not_found_title": "Không tìm thấy",
        "msg_not_found_body": "File Excel chưa được tạo.",
        "dlg_select_pdf": "Chọn file Working Manual (PDF)",
        "dlg_select_folder": "Chọn thư mục chứa file PDF",
        "dlg_select_out": "Chọn thư mục lưu file Excel",
        "theme_light": "☀️ Sáng",
        "theme_dark": "🌙 Tối",
        "empty_preview_title": "Sẵn Sàng Trích Xuất Dữ Liệu BOM",
        "empty_preview_desc": "Thêm file PDF ở cột bên trái hoặc kéo thả trực tiếp, sau đó nhấn '⚡ TRÍCH XUẤT'.",
        "log_collapse": "▲ Thu gọn",
        "log_expand": "▼ Mở rộng",
        "toast_copied": "Đã sao chép: {val}",
        "ctx_copy_part": "📋 Sao chép Part No / S&O Code",
        "ctx_copy_remark": "📌 Sao chép Vị trí (Remarks)",
        "ctx_copy_rating": "🏷️ Sao chép Thông số (Rating)",
        "ctx_copy_sn": "🔢 Sao chép Mã nội bộ (S.N.)",
        "ctx_copy_row": "📝 Sao chép toàn bộ dòng (TSV)",
        "preview_count": "{filtered}/{total} linh kiện",
        "cols": {
            "seq": "STT",
            "item_no": "Item",
            "process": "Quy Trình",
            "variant": "Phiên Bản",
            "part_no": "Part No / S&O Code",
            "sn": "Mã Nội Bộ (S.N.)",
            "rating": "Thông Số / Rating",
            "qty": "SL",
            "remarks": "Vị Trí Linh Kiện",
            "alternates": "Mã Thay Thế",
            "page": "Trang",
        },
        "sidebar_menu": "CHỨC NĂNG",
        "tab_extract": "⚡ Trích Xuất BOM",
        "tab_compare": "🔍 So Sánh BOM",
        "copyright_title": "BẢN QUYỀN SỞ HỮU",
        "copyright_text": "Bản quyền thuộc về:\nLê Văn Vững - IPQC",
        "copyright_footer": "© Bản quyền thuộc về Lê Văn Vững - IPQC",
        "card_excel_title": "1. FILE EXCEL ĐỐI CHIẾU",
        "btn_choose_excel": "📊 Chọn File Excel (.xlsx)",
        "excel_no_file": "Chưa chọn file Excel (Tối đa 10MB). Nhấn chọn hoặc kéo thả file vào đây.",
        "excel_info": "📁 {name} ({size})  |  Quy trình: {proc}  |  {count} linh kiện",
        "card_pdf_title": "2. FILE WORKING MANUAL (PDF)",
        "btn_choose_pdf": "📄 Chọn File PDF Manual",
        "pdf_no_file": "Chưa chọn file PDF (Tối đa 10MB). Nhấn chọn hoặc kéo thả file vào đây.",
        "pdf_info": "📄 {name} ({size})  |  PWB: {pwb}  |  {pages} trang",
        "lbl_select_wm_page": "Chọn trang đối chiếu:",
        "btn_compare_now": "⚡  SO SÁNH NGAY",
        "btn_export_comparison": "📊 Xuất Báo Cáo Excel",
        "btn_reset_compare": "✕ Làm Lại",
        "title_file_too_large": "File Quá Lớn (> 10MB)",
        "err_file_too_large": "File '{name}' ({size}) vượt quá giới hạn dung lượng ({max})!\n\nVui lòng chọn file dưới 10MB để tránh giật lag và đảm bảo ứng dụng vận hành mượt mà.",
        "err_batch_too_large": "Các file sau vượt quá giới hạn dung lượng ({max}) và đã bị bỏ qua để tránh giật lag:\n\n{files}",
        "verdict_placeholder": "Vui lòng chọn file Excel và file Working Manual PDF rồi nhấn '⚡ SO SÁNH NGAY'.",
        "verdict_match_title": "✅ HOÀN TOÀN TRÙNG KHỚP (100% MATCHED)",
        "verdict_match_sub": "Tất cả {total} linh kiện trong file Excel đều có mặt và trùng khớp chính xác 100% với Working Manual ({proc})!",
        "verdict_mismatch_title": "⚠️ PHÁT HIỆN {total} ĐIỂM SAI LỆCH / THIẾU TRONG FILE EXCEL!",
        "verdict_mismatch_sub": "Chi tiết: {mismatch} mục sai lệch vị trí/SL  |  {missing} mục thiếu trong WM  |  {extra} mục File Excel đang thiếu!",
        "verdict_missing_excel_title": "⚠️ PHÁT HIỆN FILE EXCEL ĐANG THIẾU {extra} MÃ LINH KIỆN!",
        "verdict_missing_excel_sub": "Working Manual ({proc}) có {extra} mã linh kiện mà file Excel chưa có. Hãy kiểm tra các dòng được đánh dấu bên dưới.",
        "filter_comp_all": "Tất cả ({n})",
        "filter_comp_mismatch": "Chỉ xem sai lệch ({n})",
        "filter_comp_match": "Chỉ xem khớp ({n})",
        "badge_excel_items": "MỤC EXCEL",
        "badge_matched_items": "MỤC KHỚP",
        "badge_mismatch_items": "SAI LỆCH",
        "comp_cols": {
            "stt": "STT",
            "status": "Trạng Thái",
            "part_no_excel": "Mã LK (Excel)",
            "part_no_wm": "Mã LK (WM)",
            "qty_excel": "SL (Excel)",
            "qty_wm": "SL (WM)",
            "loc_excel": "Vị Trí Cắm (Excel)",
            "loc_wm": "Vị Trí Cắm (WM)",
            "rating": "Quy Cách (Rating / Spec)",
            "diff": "Chi Tiết Sai LỆch"
        },
        "tab_model_comp": "🔀 So Sánh 2 Model",
        "tab_series_bom": "📑 So Sánh 2 BOM",
        "card_model_a_title": "1. MODEL A (BẢN GỐC / SERIES 1)",
        "card_model_b_title": "2. MODEL B (BẢN MỚI / SERIES 2)",
        "btn_choose_model_a": "📄 Chọn PDF Model A",
        "btn_choose_model_b": "📄 Chọn PDF Model B",
        "model_a_no_file": "Chưa chọn file PDF Model A (Tối đa 10MB). Nhấn chọn hoặc kéo thả file vào đây.",
        "model_b_no_file": "Chưa chọn file PDF Model B (Tối đa 10MB). Nhấn chọn hoặc kéo thả file vào đây.",
        "btn_run_model_comp": "⚡  SO SÁNH 2 MODEL",
        "btn_export_model_ecn": "📊 Xuất Báo Cáo ECN",
        "btn_reset_model_comp": "✕ Làm Lại",
        "badge_model_total": "TỔNG VỊ TRÍ",
        "badge_model_match": "TRÙNG KHỚP",
        "badge_model_diff": "SAI LỆCH",
        "verdict_model_placeholder": "Vui lòng chọn 2 file Working Manual PDF của 2 Model rồi nhấn '⚡ SO SÁNH 2 MODEL'.",
        "verdict_model_match_title": "✅ HAI MODEL HOÀN TOÀN TRÙNG KHỚP (100% IDENTICAL)",
        "verdict_model_match_sub": "Tất cả {total} vị trí linh kiện giữa 2 Model đều trùng khớp 100% về Part Number và Trị số!",
        "verdict_model_diff_title": "⚠️ PHÁT HIỆN {diff} ĐIỂM KHÁC BIỆT GIỮA 2 MODEL!",
        "verdict_model_diff_sub": "Chi tiết: {added} mục lắp thêm  |  {removed} mục bỏ bớt (DNP)  |  {changed} mục đổi trị số.",
        "filter_model_diff_all": "Tất cả sai lệch ({n})",
        "filter_model_added": "🟢 Lắp thêm ({n})",
        "filter_model_removed": "🔴 Bỏ bớt ({n})",
        "filter_model_changed": "🟡 Đổi trị số ({n})",
        "filter_model_match": "⚪ Khớp 100% ({n})",
        "title_drawing_inspector": "SOI BẢN VẼ PCB: {ref}",
        "inspector_empty": "Click vào 1 dòng linh kiện ở bảng bên trái để phóng to và định vị trên bản vẽ PCB.",
        "model_cols": {
            "stt": "STT",
            "ref_des": "Vị Trí (Ref)",
            "stage": "Công Đoạn",
            "status": "Trạng Thái",
            "part_a": "Model A (Part No)",
            "rating_a": "Model A (Spec)",
            "part_b": "Model B (Part No)",
            "rating_b": "Model B (Spec)",
            "note": "Chi Tiết Thay Đổi"
        },
        "disp_mode_table": "📋 Bảng Dữ Liệu",
        "disp_mode_drawing": "📐 Bản Vẽ PCB",
        "disp_mode_split": "◫ Song Song",
        "btn_fullscreen_drawing": "⛶ Toàn Màn Hình",
    },
    "zh": {
        "lang_display": "中文",
        "app_title": "⚡ VIPQC AI",
        "app_subtitle": "Working Manual  →  Excel   |   自动BOM提取与图纸比对系统",
        "badge_files": "已选文件",
        "badge_items": "BOM物料项",
        "badge_qty": "物料总数",
        "sec_queue": "文件队列",
        "btn_add_files": "+ 添加文件",
        "btn_add_folder": "📁 文件夹",
        "empty_files": "尚未选择任何文件 (单文件最大10MB)。\n拖放文件至此或点击 '+ 添加文件'。",
        "btn_clear_all": "✕ 清空列表",
        "sec_output": "导出设置",
        "switch_individual": "为每个PDF单独导出Excel",
        "switch_batch": "生成汇总Master总表 (Batch)",
        "btn_run": "⚡  开始提取并导出EXCEL",
        "btn_running": "⏳  正在处理中…",
        "btn_open_excel": "📊 打开Excel",
        "btn_open_folder": "📂 打开目录",
        "sec_preview": "实时数据预览",
        "search_placeholder": "🔍  搜索物料编码、位号、工艺流程…",
        "filter_all": "全部",
        "sec_log": "处理日志",
        "status_ready": "就绪。拖放或选择 Working Manual PDF 文件以开始。",
        "status_processing": "正在处理 [{i}/{n}]: {fname}",
        "status_done": "✅ 完成！从 {n} 个文件中提取出 {count} 条BOM记录。",
        "log_ready": "应用程序已就绪。请拖放或添加 PDF 文件并点击 ⚡ 开始提取。",
        "log_start": "🚀 开始批量处理 {n} 个文件…",
        "log_analyzing": "[{i}/{n}] 正在解析 '{fname}'…",
        "log_parsed": "   ✓ PWB={pwb}  工艺={procs}  物料项={items}  数量={qty}",
        "log_exported": "   📊 已成功导出: {name}",
        "log_error": "   ❌ 解析错误 '{fname}': {err}",
        "log_master": "🏆 MASTER 总表: {name}",
        "log_finish": "🎉 处理完成  —  共 {rows} 行BOM数据  |  总物料数: {qty}",
        "log_added_files": "已从 '{folder}' 导入 {count} 个文件。",
        "msg_no_files_title": "未选择文件",
        "msg_no_files_body": "请至少添加 1 个 PDF 文件！",
        "msg_not_found_title": "文件不存在",
        "msg_not_found_body": "Excel文件尚未生成。",
        "dlg_select_pdf": "选择 Working Manual (PDF) 文件",
        "dlg_select_folder": "选择包含 PDF 文件的文件夹",
        "dlg_select_out": "选择保存 Excel 文件的文件夹",
        "theme_light": "☀️ 浅色",
        "theme_dark": "🌙 暗色",
        "empty_preview_title": "准备提取 BOM 清单数据",
        "empty_preview_desc": "请在左侧添加 PDF 文件或直接拖放，然后点击 '⚡ 开始提取' 查看物料明细。",
        "log_collapse": "▲ 折叠",
        "log_expand": "▼ 展开",
        "toast_copied": "已复制: {val}",
        "ctx_copy_part": "📋 复制 Part No / S&O Code",
        "ctx_copy_remark": "📌 复制物料位号 (Remarks)",
        "ctx_copy_rating": "🏷️ 复制规格参数 (Rating)",
        "ctx_copy_sn": "🔢 复制内部编号 (S.N.)",
        "ctx_copy_row": "📝 复制整行数据 (TSV)",
        "preview_count": "{filtered}/{total} 条物料",
        "cols": {
            "seq": "序号",
            "item_no": "项次",
            "process": "工艺代码",
            "variant": "机型变体",
            "part_no": "Part No / S&O Code",
            "sn": "内部编号 (S.N.)",
            "rating": "规格参数 / Rating",
            "qty": "数量",
            "remarks": "位号位置",
            "alternates": "替代料号",
            "page": "页码",
        },
        "sidebar_menu": "功能菜单",
        "tab_extract": "⚡ BOM 提取",
        "tab_compare": "🔍 BOM 对比",
        "copyright_title": "版权所有",
        "copyright_text": "版权归属:\nLê Văn Vững - IPQC",
        "copyright_footer": "© 版权所有: Lê Văn Vững - IPQC",
        "card_excel_title": "1. 对比 EXCEL 文件",
        "btn_choose_excel": "📊 选择 Excel 文件 (.xlsx)",
        "excel_no_file": "未选择 Excel 文件 (最大10MB)。点击选择或拖放文件到此处。",
        "excel_info": "📁 {name} ({size})  |  工序: {proc}  |  {count} 项物料",
        "card_pdf_title": "2. WORKING MANUAL (PDF) 文件",
        "btn_choose_pdf": "📄 选择 PDF Manual",
        "pdf_no_file": "未选择 PDF 文件 (最大10MB)。点击选择或拖放文件到此处。",
        "pdf_info": "📄 {name} ({size})  |  PWB: {pwb}  |  {pages} 页",
        "lbl_select_wm_page": "选择对比页面:",
        "btn_compare_now": "⚡  开始比对",
        "btn_export_comparison": "📊 导出比对报告",
        "btn_reset_compare": "✕ 重置",
        "title_file_too_large": "文件过大 (> 10MB)",
        "err_file_too_large": "文件 '{name}' ({size}) 超过最大限制 ({max})！\n\n请选择 10MB 以下的文件，以防止系统卡顿并确保流畅运行。",
        "err_batch_too_large": "以下文件超过大小限制 ({max})，已被自动跳过以避免卡顿:\n\n{files}",
        "verdict_placeholder": "请先选择 Excel 文件与 Working Manual PDF 文件，然后点击 '⚡ 开始比对'。",
        "verdict_match_title": "✅ 完全匹配 (100% MATCHED)",
        "verdict_match_sub": "Excel 文件中的所有 {total} 项物料均与 Working Manual ({proc}) 完全匹配！",
        "verdict_mismatch_title": "⚠️ 在 EXCEL 中发现 {total} 处差异/缺失！",
        "verdict_mismatch_sub": "详情: {mismatch} 项位置/数量差异  |  {missing} 项在 WM 中缺失  |  {extra} 项 Excel 文件缺失！",
        "verdict_missing_excel_title": "⚠️ 发现 EXCEL 文件缺少 {extra} 项物料！",
        "verdict_missing_excel_sub": "Working Manual ({proc}) 中有 {extra} 项物料在 Excel 文件中缺失。请查看下方高亮标记行。",
        "filter_comp_all": "全部 ({n})",
        "filter_comp_mismatch": "仅看差异 ({n})",
        "filter_comp_match": "仅看匹配 ({n})",
        "badge_excel_items": "EXCEL 物料数",
        "badge_matched_items": "匹配项",
        "badge_mismatch_items": "差异项",
        "comp_cols": {
            "stt": "序号",
            "status": "比对状态",
            "part_no_excel": "物料编码 (Excel)",
            "part_no_wm": "物料编码 (WM)",
            "qty_excel": "数量 (Excel)",
            "qty_wm": "数量 (WM)",
            "loc_excel": "元件位置 (Excel)",
            "loc_wm": "元件位置 (WM)",
            "rating": "规格 / Rating",
            "diff": "差异详情"
        },
        "tab_model_comp": "🔀 机型图纸比对",
        "tab_series_bom": "📑 BOM系列比对",
        "card_model_a_title": "1. 机型 A (基准 / 系列 1)",
        "card_model_b_title": "2. 机型 B (变更 / 系列 2)",
        "btn_choose_model_a": "📄 选择机型 A (PDF)",
        "btn_choose_model_b": "📄 选择机型 B (PDF)",
        "model_a_no_file": "尚未选择机型 A 文件 (最大10MB)。",
        "model_b_no_file": "尚未选择机型 B 文件 (最大10MB)。",
        "btn_run_model_comp": "⚡  开始比对机型",
        "btn_export_model_ecn": "📊 导出ECN报告",
        "btn_reset_model_comp": "✕ 重置",
        "badge_model_total": "总物料位",
        "badge_model_match": "完全一致",
        "badge_model_diff": "差异物料",
        "verdict_model_placeholder": "请选择两个机型的 Working Manual PDF 并点击 '⚡ 开始比对机型'。",
        "verdict_model_match_title": "✅ 两个机型完全一致 (100% MATCH)",
        "verdict_model_match_sub": "两个机型之间的全部 {total} 处物料编码与规格均100%一致！",
        "verdict_model_diff_title": "⚠️ 发现两个机型存在 {diff} 处差异！",
        "verdict_model_diff_sub": "详情: {added} 处新增  |  {removed} 处删除 (DNP)  |  {changed} 处参数变更。",
        "filter_model_diff_all": "全部差异 ({n})",
        "filter_model_added": "🟢 新增物料 ({n})",
        "filter_model_removed": "🔴 删除物料 ({n})",
        "filter_model_changed": "🟡 参数变更 ({n})",
        "filter_model_match": "⚪ 完全一致 ({n})",
        "title_drawing_inspector": "PCB图纸核验: {ref}",
        "inspector_empty": "点击左侧物料列表项以在PCB图纸上定位放大查看。",
        "model_cols": {
            "stt": "序号",
            "ref_des": "位置代码",
            "stage": "工序",
            "status": "状态",
            "part_a": "机型 A (物料编码)",
            "rating_a": "机型 A (规格)",
            "part_b": "机型 B (物料编码)",
            "rating_b": "机型 B (规格)",
            "note": "变更详情"
        },
        "disp_mode_table": "📋 数据表格",
        "disp_mode_drawing": "📐 PCB 图纸",
        "disp_mode_split": "◫ 双栏并列",
        "btn_fullscreen_drawing": "⛶ 全屏显示",
    },
    "en": {
        "lang_display": "English",
        "app_title": "⚡ VIPQC AI",
        "app_subtitle": "Working Manual  →  Excel   |   AI BOM Extraction & Drawing Comparator",
        "badge_files": "FILES QUEUED",
        "badge_items": "BOM ITEMS",
        "badge_qty": "TOTAL QTY",
        "sec_queue": "FILE QUEUE",
        "btn_add_files": "+ Add Files",
        "btn_add_folder": "📁 Folder",
        "empty_files": "No files selected yet.\nDrag & drop files here or click '+ Add Files'.",
        "btn_clear_all": "✕ Clear all",
        "sec_output": "OUTPUT SETTINGS",
        "switch_individual": "Export individual Excel per PDF",
        "switch_batch": "Create Master Consolidated Excel (Batch)",
        "btn_run": "⚡  EXTRACT & EXPORT EXCEL",
        "btn_running": "⏳  PROCESSING…",
        "btn_open_excel": "📊 Open Excel",
        "btn_open_folder": "📂 Folder",
        "sec_preview": "LIVE DATA PREVIEW",
        "search_placeholder": "🔍  Search part number, location, process…",
        "filter_all": "All",
        "sec_log": "PROCESSING LOG",
        "status_ready": "Ready. Drag & drop or select Working Manual PDF files to start.",
        "status_processing": "Processing [{i}/{n}]: {fname}",
        "status_done": "✅ Finished! {count} BOM records from {n} files.",
        "log_ready": "Application ready. Drag & drop or add PDF files and click ⚡ Extract.",
        "log_start": "🚀 STARTING PROCESSING {n} FILES…",
        "log_analyzing": "[{i}/{n}] Analyzing '{fname}'…",
        "log_parsed": "   ✓ PWB={pwb}  Process={procs}  Items={items}  Qty={qty}",
        "log_exported": "   📊 Exported: {name}",
        "log_error": "   ❌ Error '{fname}': {err}",
        "log_master": "🏆 MASTER EXCEL: {name}",
        "log_finish": "🎉 FINISHED  —  {rows} BOM rows  |  Total Qty: {qty}",
        "log_added_files": "Added {count} files from '{folder}'.",
        "msg_no_files_title": "No files",
        "msg_no_files_body": "Please add at least 1 PDF file!",
        "msg_not_found_title": "Not found",
        "msg_not_found_body": "Excel file has not been created yet.",
        "dlg_select_pdf": "Select Working Manual (PDF) Files",
        "dlg_select_folder": "Select Folder Containing PDF Files",
        "dlg_select_out": "Select Output Directory for Excel Files",
        "theme_light": "☀️ Light",
        "theme_dark": "🌙 Dark",
        "empty_preview_title": "Ready to Extract BOM Data",
        "empty_preview_desc": "Add PDF files in the left queue or drag & drop directly, then click '⚡ EXTRACT'.",
        "log_collapse": "▲ Collapse",
        "log_expand": "▼ Expand",
        "toast_copied": "Copied: {val}",
        "ctx_copy_part": "📋 Copy Part No / S&O Code",
        "ctx_copy_remark": "📌 Copy Remarks / Designators",
        "ctx_copy_rating": "🏷️ Copy Rating / Spec",
        "ctx_copy_sn": "🔢 Copy Internal S.N.",
        "ctx_copy_row": "📝 Copy Entire Row (TSV)",
        "preview_count": "{filtered}/{total} items",
        "cols": {
            "seq": "No.",
            "item_no": "Item",
            "process": "Process Code",
            "variant": "Model Variant",
            "part_no": "Part No / S&O Code",
            "sn": "Internal S.N.",
            "rating": "Rating / Spec",
            "qty": "Qty",
            "remarks": "Reference Designators",
            "alternates": "Alternate Parts",
            "page": "Page",
        },
        "sidebar_menu": "MAIN MENU",
        "tab_extract": "⚡ BOM Extractor",
        "tab_compare": "🔍 BOM Comparator",
        "copyright_title": "COPYRIGHT",
        "copyright_text": "Copyright belongs to:\nLê Văn Vững - IPQC",
        "copyright_footer": "© Copyright by Lê Văn Vững - IPQC",
        "card_excel_title": "1. EXCEL AUDIT FILE",
        "btn_choose_excel": "📊 Choose Excel File (.xlsx)",
        "excel_no_file": "No Excel file selected. Click to browse or drag & drop here.",
        "excel_info": "📁 {name}  |  Process: {proc}  |  {count} items",
        "card_pdf_title": "2. WORKING MANUAL (PDF)",
        "btn_choose_pdf": "📄 Choose PDF Manual",
        "pdf_no_file": "No PDF file selected. Click to browse or drag & drop here.",
        "pdf_info": "📄 {name}  |  PWB: {pwb}  |  {pages} pages",
        "lbl_select_wm_page": "Select page to compare:",
        "btn_compare_now": "⚡  RUN COMPARISON",
        "btn_export_comparison": "📊 Export Report",
        "btn_reset_compare": "✕ Reset",
        "verdict_placeholder": "Please select both an Excel file and a Working Manual PDF, then click '⚡ RUN COMPARISON'.",
        "verdict_match_title": "✅ 100% EXACT MATCH",
        "verdict_match_sub": "All {total} components in the Excel file are present and 100% matched with Working Manual ({proc})!",
        "verdict_mismatch_title": "⚠️ {total} DISCREPANCIES / MISSING ITEMS DETECTED!",
        "verdict_mismatch_sub": "Breakdown: {mismatch} position/qty mismatches  |  {missing} missing in WM  |  {extra} missing in Excel!",
        "verdict_missing_excel_title": "⚠️ DETECTED {extra} ITEMS MISSING IN EXCEL FILE!",
        "verdict_missing_excel_sub": "Working Manual ({proc}) has {extra} components that are missing in the Excel file. Please inspect highlighted rows below.",
        "filter_comp_all": "All ({n})",
        "filter_comp_mismatch": "Discrepancies ({n})",
        "filter_comp_match": "Matches ({n})",
        "badge_excel_items": "EXCEL ITEMS",
        "badge_matched_items": "MATCHED",
        "badge_mismatch_items": "MISMATCHES",
        "comp_cols": {
            "stt": "No.",
            "status": "Status",
            "part_no_excel": "Part No (Excel)",
            "part_no_wm": "Part No (WM)",
            "qty_excel": "Qty (Excel)",
            "qty_wm": "Qty (WM)",
            "loc_excel": "Locations (Excel)",
            "loc_wm": "Locations (WM)",
            "rating": "Rating / Specification",
            "diff": "Discrepancy Details"
        },
        "tab_model_comp": "🔀 Model Series Diff",
        "tab_series_bom": "📑 Series BOM Diff",
        "card_model_a_title": "1. MODEL A (BASE / SERIES 1)",
        "card_model_b_title": "2. MODEL B (NEW / SERIES 2)",
        "btn_choose_model_a": "📄 Choose Model A PDF",
        "btn_choose_model_b": "📄 Choose Model B PDF",
        "model_a_no_file": "No Model A PDF selected (Max 10MB).",
        "model_b_no_file": "No Model B PDF selected (Max 10MB).",
        "btn_run_model_comp": "⚡  COMPARE MODELS",
        "btn_export_model_ecn": "📊 Export ECN Report",
        "btn_reset_model_comp": "✕ Reset",
        "badge_model_total": "TOTAL POSITIONS",
        "badge_model_match": "MATCHED",
        "badge_model_diff": "DIFFERENCES",
        "verdict_model_placeholder": "Please select two Model Working Manual PDFs and click '⚡ COMPARE MODELS'.",
        "verdict_model_match_title": "✅ BOTH MODELS ARE 100% IDENTICAL",
        "verdict_model_match_sub": "All {total} component positions between both models match 100% in Part Number and Rating!",
        "verdict_model_diff_title": "⚠️ {diff} DIFFERENCES DETECTED BETWEEN MODELS!",
        "verdict_model_diff_sub": "Breakdown: {added} added  |  {removed} removed (DNP)  |  {changed} spec changed.",
        "filter_model_diff_all": "All Differences ({n})",
        "filter_model_added": "🟢 Added ({n})",
        "filter_model_removed": "🔴 Removed ({n})",
        "filter_model_changed": "🟡 Changed ({n})",
        "filter_model_match": "⚪ 100% Match ({n})",
        "title_drawing_inspector": "PCB INSPECTOR: {ref}",
        "inspector_empty": "Click a component row in the table to zoom and locate on PCB drawing.",
        "model_cols": {
            "stt": "No.",
            "ref_des": "Ref Des",
            "stage": "Stage",
            "status": "Status",
            "part_a": "Model A (Part No)",
            "rating_a": "Model A (Spec)",
            "part_b": "Model B (Part No)",
            "rating_b": "Model B (Spec)",
            "note": "Change Note"
        },
        "disp_mode_table": "📋 Data Table",
        "disp_mode_drawing": "📐 PCB Drawing",
        "disp_mode_split": "◫ Side-by-Side",
        "btn_fullscreen_drawing": "⛶ Full Screen",
    }
}


class AnimatedProgressBar(ctk.CTkProgressBar):
    """Thin glowing progress bar."""
    def __init__(self, master, **kw):
        super().__init__(master,
                         height=6,
                         corner_radius=3,
                         fg_color=BG_SURFACE,
                         progress_color=ACCENT_TEAL,
                         **kw)
        self.set(0)


# ─── High-Definition Anti-Aliased RGB Rotating Border Generator ───────────────
RGB_FRAME_CACHE = {}


def get_rgb_frames(w=126, h=54, r=10, thickness=2, num_frames=24, is_dark=True):
    """
    Pre-renders smooth, anti-aliased rotating conic rainbow borders using PIL supersampling.
    Caches 24 frames so runtime animation costs virtually 0% CPU.
    """
    card_hex = BG_CARD[1] if is_dark else BG_CARD[0]
    surface_hex = BG_SURFACE[1] if is_dark else BG_SURFACE[0]
    cache_key = (w, h, r, thickness, num_frames, card_hex, surface_hex)
    if cache_key in RGB_FRAME_CACHE:
        return RGB_FRAME_CACHE[cache_key]

    scale = 2
    W = w * scale
    H = h * scale
    R = r * scale
    T = thickness * scale
    cx = W / 2.0
    cy = H / 2.0

    def hex_to_rgb(h_str):
        h_str = h_str.lstrip("#")
        return tuple(int(h_str[i:i+2], 16) for i in (0, 2, 4))

    c_card = hex_to_rgb(card_hex)
    c_surface = hex_to_rgb(surface_hex)

    frames = []
    diag = int(math.hypot(W, H)) + 10
    bbox = [cx - diag, cy - diag, cx + diag, cy + diag]
    n_slices = 120
    step = 360.0 / n_slices

    for f in range(num_frames):
        shift = f / num_frames
        im = Image.new("RGB", (W, H), c_card)
        draw = ImageDraw.Draw(im)

        for s in range(n_slices):
            start_ang = s * step
            end_ang = start_ang + step + 0.5
            mid_ang = (start_ang + end_ang) / 2.0
            hue = (mid_ang / 360.0 - shift) % 1.0
            r_c, g_c, b_c = colorsys.hsv_to_rgb(hue, 0.85 if is_dark else 0.9, 1.0)
            col = (int(r_c * 255), int(g_c * 255), int(b_c * 255))
            draw.pieslice(bbox, start=start_ang, end=end_ang, fill=col)

        draw.rounded_rectangle([T, T, W - 1 - T, H - 1 - T], radius=max(1, R - T), fill=c_surface)

        alpha = Image.new("L", (W, H), 0)
        draw_a = ImageDraw.Draw(alpha)
        draw_a.rounded_rectangle([0, 0, W - 1, H - 1], radius=R, fill=255)

        bg_im = Image.new("RGB", (W, H), c_card)
        im = Image.composite(im, bg_im, alpha)

        im_down = im.resize((w, h), Image.Resampling.LANCZOS)
        frames.append(im_down)

    photo_frames = [ImageTk.PhotoImage(f) for f in frames]
    RGB_FRAME_CACHE[cache_key] = photo_frames
    return photo_frames


class StatBadge(tk.Canvas):
    """A sleek stat badge with smooth anti-aliased rotating RGB border."""
    def __init__(self, master, label: str, color=None, offset=0, w=130, h=56, **kw):
        is_dark = ctk.get_appearance_mode().lower() == "dark"
        card_hex = BG_CARD[1] if is_dark else BG_CARD[0]
        super().__init__(master, width=w, height=h, highlightthickness=0, bg=card_hex, bd=0)
        self.w = w
        self.h = h
        self.is_dark = is_dark
        self.offset = offset
        self.num_frames = 24
        self.photo_frames = get_rgb_frames(w, h, r=10, thickness=2, num_frames=self.num_frames, is_dark=is_dark)

        self.current_idx = offset % self.num_frames
        self.bg_img = self.create_image(0, 0, anchor="nw", image=self.photo_frames[self.current_idx])

        val_color = "#FFFFFF" if is_dark else "#0F172A"
        lbl_color = "#94A3B8" if is_dark else "#64748B"

        self.val_text = self.create_text(w // 2, 20, text="0", font=("Segoe UI", 18, "bold"), fill=val_color)
        self.lbl_text = self.create_text(w // 2, 40, text=label, font=("Segoe UI", 9, "bold"), fill=lbl_color)

    def step_animation(self, tick: int):
        idx = (tick + self.offset) % self.num_frames
        self.current_idx = idx
        self.itemconfigure(self.bg_img, image=self.photo_frames[idx])

    def set_value(self, val):
        self.itemconfigure(self.val_text, text=str(val))

    def set_label(self, label: str):
        self.itemconfigure(self.lbl_text, text=label)

    def set_color(self, color=None):
        pass

    def set_theme(self, is_dark: bool):
        self.is_dark = is_dark
        card_hex = BG_CARD[1] if is_dark else BG_CARD[0]
        self.configure(bg=card_hex)
        self.photo_frames = get_rgb_frames(self.w, self.h, r=10, thickness=2, num_frames=self.num_frames, is_dark=is_dark)
        self.itemconfigure(self.bg_img, image=self.photo_frames[self.current_idx % len(self.photo_frames)])
        val_color = "#FFFFFF" if is_dark else "#0F172A"
        lbl_color = "#94A3B8" if is_dark else "#64748B"
        self.itemconfigure(self.val_text, fill=val_color)
        self.itemconfigure(self.lbl_text, fill=lbl_color)


class FileRow(ctk.CTkFrame):
    """Single removable file row inside the queue list."""
    def __init__(self, master, path: str, on_remove, **kw):
        super().__init__(master, fg_color=BG_SURFACE, corner_radius=8,
                         border_width=0, **kw)
        self.path = path
        self._on_remove = on_remove

        # icon
        ctk.CTkLabel(self, text="📄", font=("Segoe UI", 13), width=28,
                     fg_color="transparent", text_color=ACCENT_BLUE).pack(side="left", padx=(8, 4), pady=6)
        # filename
        fname = os.path.basename(path)
        ctk.CTkLabel(self, text=fname, font=FONT_BODY, text_color=TEXT_PRIMARY,
                     fg_color="transparent", anchor="w").pack(side="left", fill="x", expand=True, padx=4)
        # formatted size label (e.g. 450 KB, 1.2 MB)
        size_bytes = os.path.getsize(path) if os.path.exists(path) else 0
        size_str = format_file_size(size_bytes)
        ctk.CTkLabel(self, text=size_str, font=("Segoe UI", 9, "bold"), text_color=TEXT_MUTED,
                     fg_color="transparent").pack(side="right", padx=(2, 6))
        # remove button
        ctk.CTkButton(self, text="✕", width=26, height=26,
                       fg_color="transparent", text_color=TEXT_MUTED,
                       hover_color=("#FCA5A5", "#3A2030"),
                       font=("Segoe UI", 11, "bold"),
                       command=self._remove).pack(side="right", padx=6)

    def _remove(self):
        self._on_remove(self.path)
        self.destroy()


# ─── GitHub Auto-Update Notification & Download Dialog ─────────────────────────
class UpdateDialog(ctk.CTkToplevel):
    """
    Holographic update notification dialog that displays new release details,
    the changelog, and allows one-click background download and in-place update.
    """
    def __init__(self, parent, update_info: dict):
        super().__init__(parent)
        self.parent = parent
        self.update_info = update_info
        self._is_downloading = False

        self.title("🚀 Cập Nhật VIPQC AI")
        self.geometry("540x480")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Center over parent window
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 540) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 480) // 2
        self.geometry(f"+{max(10, px)}+{max(10, py)}")

        self.configure(fg_color=BG_CARD)

        # Outer card with glowing accent border
        card = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12, border_width=1.5, border_color=ACCENT_TEAL)
        card.pack(fill="both", expand=True, padx=16, pady=16)

        # Header with rocket icon and version banner
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=(16, 6))

        ctk.CTkLabel(hdr, text="🚀", font=("Segoe UI", 26)).pack(side="left", padx=(0, 10))

        title_box = ctk.CTkFrame(hdr, fg_color="transparent")
        title_box.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            title_box, text="ĐÃ CÓ PHIÊN BẢN CẬP NHẬT MỚI!",
            font=("Segoe UI", 12, "bold"), text_color=ACCENT_TEAL, anchor="w"
        ).pack(fill="x")

        curr_v = update_info.get("current_version", updater.CURRENT_VERSION)
        lat_v = update_info.get("latest_version", "")
        ctk.CTkLabel(
            title_box, text=f"Hiện tại: v{curr_v}   ➜   Mới nhất: v{lat_v}",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x")

        # Subtitle / Release title
        rel_title = update_info.get("title") or f"VIPQC AI v{lat_v}"
        ctk.CTkLabel(
            card, text=rel_title,
            font=("Segoe UI", 11, "bold"), text_color=ACCENT_BLUE, anchor="w"
        ).pack(fill="x", padx=20, pady=(2, 6))

        # Changelog Header
        ctk.CTkLabel(
            card, text="📋 Nội dung bản cập nhật mới (Changelog):",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED, anchor="w"
        ).pack(fill="x", padx=20, pady=(2, 2))

        # Changelog Textbox
        tb = ctk.CTkTextbox(
            card, height=130, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY,
            border_width=1, border_color=BORDER_CLR, font=("Segoe UI", 10)
        )
        tb.pack(fill="x", padx=20, pady=(0, 10))
        changelog_text = update_info.get("changelog") or "Bản cập nhật cải tiến hiệu năng và tính năng mới."
        tb.insert("1.0", changelog_text)
        tb.configure(state="disabled")

        # Download Progress Area (Initially Hidden)
        self.prog_frame = ctk.CTkFrame(card, fg_color="transparent")

        self.lbl_progress = ctk.CTkLabel(
            self.prog_frame, text="Đang chuẩn bị tải xuống...",
            font=("Segoe UI", 9, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.lbl_progress.pack(fill="x", pady=(0, 4))

        self.prog_bar = ctk.CTkProgressBar(
            self.prog_frame, height=10, corner_radius=5,
            fg_color=BG_SURFACE, progress_color=ACCENT_TEAL
        )
        self.prog_bar.set(0.0)
        self.prog_bar.pack(fill="x")

        # Buttons Row
        self.btn_row = ctk.CTkFrame(card, fg_color="transparent")
        self.btn_row.pack(side="bottom", fill="x", padx=20, pady=(10, 14))

        self.btn_update = ctk.CTkButton(
            self.btn_row, text="⚡ Cập Nhật Tự Động",
            font=("Segoe UI", 11, "bold"), height=36,
            fg_color=ACCENT_TEAL, hover_color=("#0F766E", "#00A88C"),
            command=self._start_auto_update
        )
        self.btn_update.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.btn_github = ctk.CTkButton(
            self.btn_row, text="🌐 Mở GitHub",
            font=("Segoe UI", 10, "bold"), height=36, width=105,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            command=self._open_github
        )
        self.btn_github.pack(side="left", padx=(0, 6))

        self.btn_cancel = ctk.CTkButton(
            self.btn_row, text="❌ Để Sau",
            font=("Segoe UI", 10), height=36, width=80,
            fg_color="transparent", text_color=TEXT_MUTED, hover_color=BG_HOVER,
            command=self.destroy
        )
        self.btn_cancel.pack(side="right")

    def _start_auto_update(self):
        if self._is_downloading:
            return
        self._is_downloading = True

        self.btn_update.configure(state="disabled", text="⚡ Đang tải...")
        self.btn_github.configure(state="disabled")
        self.btn_cancel.configure(state="disabled")
        self.prog_frame.pack(fill="x", padx=20, pady=(0, 10))

        download_url = self.update_info.get("download_url") or updater.DEFAULT_EXE_DOWNLOAD_URL
        latest_v = self.update_info.get("latest_version", "new")
        target_file = os.path.join(tempfile.gettempdir(), f"VIPQC_AI_Update_{latest_v}.exe")

        def _bg():
            try:
                def on_progress(pct, downloaded, total):
                    def update_ui():
                        self.prog_bar.set(pct / 100.0)
                        dl_mb = downloaded / (1024 * 1024)
                        tot_mb = total / (1024 * 1024)
                        self.lbl_progress.configure(
                            text=f"Đang tải: {pct:.0f}% ({dl_mb:.1f} MB / {tot_mb:.1f} MB)..."
                        )
                    self.after(0, update_ui)

                ok = updater.download_update_file(download_url, target_file, progress_callback=on_progress)
                if ok:
                    self.after(0, lambda: self.lbl_progress.configure(text="✅ Tải thành công! Đang thay thế và khởi động lại..."))
                    time.sleep(1.2)
                    updater.apply_update_and_restart(target_file)
                else:
                    self.after(0, lambda: self._on_error("File tải về không hoàn chỉnh hoặc rỗng."))
            except Exception as e:
                self.after(0, lambda err=e: self._on_error(str(err)))

        threading.Thread(target=_bg, daemon=True).start()

    def _on_error(self, err_msg):
        self._is_downloading = False
        self.btn_update.configure(state="normal", text="Thử Lại")
        self.btn_github.configure(state="normal")
        self.btn_cancel.configure(state="normal")
        self.lbl_progress.configure(text=f"❌ Lỗi: {err_msg}", text_color="#EF4444")
        messagebox.showerror(
            "Lỗi Cập Nhật",
            f"Không thể tải bản cập nhật tự động:\n{err_msg}\n\nBạn có thể bấm 'Mở GitHub' để tải file trực tiếp."
        )

    def _open_github(self):
        url = self.update_info.get("release_url") or updater.GITHUB_REPO_URL
        webbrowser.open(url)


class BOMCompareView(ctk.CTkFrame):
    """
    Dedicated workspace for comparing Excel production BOM against Working Manual PDF pages.
    Supports file uploads, drag & drop, multi-page selection, live bidirectional diff,
    interactive treeview table, and color-coded Excel audit report export.
    """
    def __init__(self, master, app, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.app = app

        # Internal State
        self.excel_data = None
        self.pdf_path = None
        self.wm_pages_info = []
        self.selected_page_idx = 0
        self.selected_page_str = ctk.StringVar(value="")
        self.selected_sheet_str = ctk.StringVar(value="")
        self.comparison_result = None
        self.filter_mode = "all"
        self.search_query = ""
        self._is_comparing: bool = False
        self._comp_hud = None

        self.col_specs = [
            ("stt", 45, "center"),
            ("status", 125, "center"),
            ("part_no_excel", 150, "w"),
            ("part_no_wm", 150, "w"),
            ("qty_excel", 75, "e"),
            ("qty_wm", 75, "e"),
            ("loc_excel", 190, "w"),
            ("loc_wm", 190, "w"),
            ("rating", 180, "w"),
            ("diff", 260, "w"),
        ]

        self._build_ui()

    def t(self, key: str, **kwargs) -> str:
        return self.app.t(key, **kwargs)

    def _build_ui(self):
        # 1. TOP CARDS (Excel File Left + PDF Manual Right)
        top_cards = ctk.CTkFrame(self, fg_color="transparent")
        top_cards.pack(fill="x", pady=(0, 10))

        # Card Left: Excel Input
        self.card_excel = ctk.CTkFrame(top_cards, fg_color=BG_CARD, corner_radius=12,
                                       border_width=1, border_color=BORDER_CLR)
        self.card_excel.pack(side="left", fill="both", expand=True, padx=(0, 6))

        c_ex_top = ctk.CTkFrame(self.card_excel, fg_color="transparent")
        c_ex_top.pack(fill="x", padx=14, pady=(12, 4))

        self.lbl_card_ex_title = ctk.CTkLabel(
            c_ex_top, text=self.t("card_excel_title"),
            font=("Segoe UI", 12, "bold"), text_color=ACCENT_BLUE
        )
        self.lbl_card_ex_title.pack(side="left")

        self.lbl_excel_details = ctk.CTkLabel(
            c_ex_top, text="",
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_TEAL, anchor="e"
        )
        self.lbl_excel_details.pack(side="right")

        c_ex_mid = ctk.CTkFrame(self.card_excel, fg_color="transparent")
        c_ex_mid.pack(fill="x", padx=14, pady=(0, 6))

        self.btn_choose_excel = ctk.CTkButton(
            c_ex_mid, text=self.t("btn_choose_excel"),
            font=FONT_H2, fg_color=ACCENT_BLUE, hover_color="#1D4ED8",
            corner_radius=8, height=36, command=self._select_excel
        )
        self.btn_choose_excel.pack(side="left", padx=(0, 10))

        self.lbl_excel_name = ctk.CTkLabel(
            c_ex_mid, text=self.t("excel_no_file"),
            font=FONT_BODY, text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_excel_name.pack(side="left", fill="x", expand=True)

        # Excel bottom row (Sheet selector with prev/next buttons)
        self.c_ex_bot = ctk.CTkFrame(self.card_excel, fg_color=BG_SURFACE, corner_radius=8, height=40)
        self.c_ex_bot.pack(fill="x", padx=14, pady=(0, 10))
        self.c_ex_bot.pack_propagate(False)

        self.lbl_sheet_picker = ctk.CTkLabel(
            self.c_ex_bot, text="📑  Sheet:",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED
        )
        self.lbl_sheet_picker.pack(side="left", padx=(10, 4))

        self.btn_prev_sheet = ctk.CTkButton(
            self.c_ex_bot, text="◀", width=30, height=28,
            font=("Segoe UI", 10, "bold"), fg_color=BG_CARD,
            hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            corner_radius=6, command=self._prev_sheet, state="disabled"
        )
        self.btn_prev_sheet.pack(side="left", padx=(0, 3), pady=4)

        self.opt_sheets = ctk.CTkOptionMenu(
            self.c_ex_bot, variable=self.selected_sheet_str, values=["Sheet1"],
            height=30, corner_radius=6,
            font=("Segoe UI", 10, "bold"), dropdown_font=("Segoe UI", 10),
            fg_color=BG_CARD, button_color=BG_HOVER, text_color=TEXT_PRIMARY,
            command=self._on_sheet_change
        )
        self.opt_sheets.pack(side="left", fill="x", expand=True, padx=(0, 3), pady=4)

        self.btn_next_sheet = ctk.CTkButton(
            self.c_ex_bot, text="▶", width=30, height=28,
            font=("Segoe UI", 10, "bold"), fg_color=BG_CARD,
            hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            corner_radius=6, command=self._next_sheet, state="disabled"
        )
        self.btn_next_sheet.pack(side="left", padx=(0, 8), pady=4)

        # Card Right: PDF Manual Input
        self.card_pdf = ctk.CTkFrame(top_cards, fg_color=BG_CARD, corner_radius=12,
                                     border_width=1, border_color=BORDER_CLR)
        self.card_pdf.pack(side="right", fill="both", expand=True, padx=(6, 0))

        c_pdf_top = ctk.CTkFrame(self.card_pdf, fg_color="transparent")
        c_pdf_top.pack(fill="x", padx=14, pady=(12, 4))

        self.lbl_card_pdf_title = ctk.CTkLabel(
            c_pdf_top, text=self.t("card_pdf_title"),
            font=("Segoe UI", 12, "bold"), text_color=ACCENT_TEAL
        )
        self.lbl_card_pdf_title.pack(side="left")

        self.lbl_pdf_details = ctk.CTkLabel(
            c_pdf_top, text="",
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_BLUE, anchor="e"
        )
        self.lbl_pdf_details.pack(side="right")

        c_pdf_mid = ctk.CTkFrame(self.card_pdf, fg_color="transparent")
        c_pdf_mid.pack(fill="x", padx=14, pady=(0, 6))

        self.btn_choose_pdf = ctk.CTkButton(
            c_pdf_mid, text=self.t("btn_choose_pdf"),
            font=FONT_H2, fg_color=ACCENT_TEAL, hover_color="#0F766E",
            corner_radius=8, height=36, command=self._select_pdf
        )
        self.btn_choose_pdf.pack(side="left", padx=(0, 10))

        self.lbl_pdf_name = ctk.CTkLabel(
            c_pdf_mid, text=self.t("pdf_no_file"),
            font=FONT_BODY, text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_pdf_name.pack(side="left", fill="x", expand=True)

        # PROMINENT PAGE PICKER ROW (Nổi bật, người dùng click chọn trang bất kỳ hoặc bấm ◀ / ▶)
        self.c_pdf_bot = ctk.CTkFrame(
            self.card_pdf, fg_color=BG_SURFACE, corner_radius=8,
            border_width=1, border_color=ACCENT_TEAL, height=40
        )
        self.c_pdf_bot.pack(fill="x", padx=14, pady=(0, 10))
        self.c_pdf_bot.pack_propagate(False)

        self.lbl_page_picker = ctk.CTkLabel(
            self.c_pdf_bot, text="📑  " + self.t("lbl_select_wm_page"),
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_TEAL
        )
        self.lbl_page_picker.pack(side="left", padx=(10, 4))

        self.btn_prev_page = ctk.CTkButton(
            self.c_pdf_bot, text="◀", width=30, height=28,
            font=("Segoe UI", 10, "bold"), fg_color=BG_CARD,
            hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            corner_radius=6, command=self._prev_page, state="disabled"
        )
        self.btn_prev_page.pack(side="left", padx=(0, 3), pady=4)

        self.opt_pages = ctk.CTkOptionMenu(
            self.c_pdf_bot, variable=self.selected_page_str, values=["(Chưa nạp PDF - Hãy chọn file PDF trước)"],
            height=30, corner_radius=6,
            font=("Segoe UI", 10, "bold"), dropdown_font=("Segoe UI", 10),
            fg_color=BG_CARD, button_color=ACCENT_TEAL, text_color=TEXT_PRIMARY,
            command=self._on_page_change
        )
        self.opt_pages.pack(side="left", fill="x", expand=True, padx=(0, 3), pady=4)

        self.btn_next_page = ctk.CTkButton(
            self.c_pdf_bot, text="▶", width=30, height=28,
            font=("Segoe UI", 10, "bold"), fg_color=BG_CARD,
            hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            corner_radius=6, command=self._next_page, state="disabled"
        )
        self.btn_next_page.pack(side="left", padx=(0, 8), pady=4)

        # 2. ACTION BAR & VERDICT BANNER
        mid_bar = ctk.CTkFrame(self, fg_color="transparent")
        mid_bar.pack(fill="x", pady=(0, 10))

        self.btn_run_comp = ctk.CTkButton(
            mid_bar, text=self.t("btn_compare_now"),
            font=("Segoe UI", 12, "bold"), fg_color=ACCENT_TEAL,
            hover_color="#0F766E", corner_radius=8, height=42, width=175,
            command=self._run_comparison
        )
        self.btn_run_comp.pack(side="left", padx=(0, 8))

        # Target page indicator chip
        self.chip_target_page = ctk.CTkFrame(
            mid_bar, fg_color=BG_CARD, corner_radius=8,
            border_width=1, border_color=BORDER_CLR, height=42
        )
        self.chip_target_page.pack(side="left", padx=(0, 8))

        self.lbl_target_chip = ctk.CTkLabel(
            self.chip_target_page,
            text="📑 Trang đối chiếu: (Chưa chọn)",
            font=("Segoe UI", 10, "bold"),
            text_color=ACCENT_TEAL
        )
        self.lbl_target_chip.pack(padx=12, pady=10)

        self.btn_export_comp = ctk.CTkButton(
            mid_bar, text=self.t("btn_export_comparison"),
            font=FONT_H2, fg_color=BG_CARD,
            hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            border_width=1, border_color=BORDER_CLR,
            corner_radius=8, height=42,
            command=self._export_report, state="disabled"
        )
        self.btn_export_comp.pack(side="left", padx=(0, 8))

        self.btn_reset_comp = ctk.CTkButton(
            mid_bar, text=self.t("btn_reset_compare"),
            font=FONT_H2, fg_color="transparent",
            hover_color=("#FEE2E2", "#3A2030"), text_color=TEXT_MUTED,
            corner_radius=8, height=42, width=75,
            command=self._reset
        )
        self.btn_reset_comp.pack(side="left", padx=(0, 10))

        # Verdict Banner
        self.verdict_banner = ctk.CTkFrame(
            mid_bar, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CLR, height=48
        )
        self.verdict_banner.pack(side="right", fill="x", expand=True)

        v_inner = ctk.CTkFrame(self.verdict_banner, fg_color="transparent")
        v_inner.pack(fill="both", expand=True, padx=14, pady=4)

        self.lbl_verdict_title = ctk.CTkLabel(
            v_inner, text=self.t("verdict_placeholder"),
            font=("Segoe UI", 12, "bold"), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_verdict_title.pack(fill="x")

        self.lbl_verdict_sub = ctk.CTkLabel(
            v_inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_verdict_sub.pack(fill="x")

        # 3. COMPARISON TABLE CARD
        self.table_card = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                                       border_width=1, border_color=BORDER_CLR)
        self.table_card.pack(fill="both", expand=True)

        tbl_top = ctk.CTkFrame(self.table_card, fg_color="transparent")
        tbl_top.pack(fill="x", padx=16, pady=(12, 8))

        self.filter_tabs = ctk.CTkSegmentedButton(
            tbl_top,
            values=[self.t("filter_comp_all", n=0),
                    self.t("filter_comp_mismatch", n=0),
                    self.t("filter_comp_match", n=0)],
            corner_radius=8, height=34,
            font=("Segoe UI", 10, "bold"),
            fg_color=BG_SURFACE,
            selected_color=ACCENT_TEAL,
            selected_hover_color=ACCENT_TEAL,
            unselected_color=BG_SURFACE,
            unselected_hover_color=BG_HOVER,
            text_color=TEXT_PRIMARY,
            command=self._on_filter_change
        )
        self.filter_tabs.set(self.t("filter_comp_all", n=0))
        self.filter_tabs.pack(side="left", padx=(0, 12))

        self.search_comp = ctk.CTkEntry(
            tbl_top,
            placeholder_text="🔍  Tìm kiếm mã LK, vị trí, quy cách…",
            font=FONT_BODY, height=34, corner_radius=8,
            fg_color=BG_SURFACE, border_width=1, border_color=BORDER_CLR,
            text_color=TEXT_PRIMARY
        )
        self.search_comp.pack(side="left", fill="x", expand=True, padx=(0, 12))
        self.search_comp.bind("<KeyRelease>", self._on_search_key)

        self.lbl_comp_count = ctk.CTkLabel(
            tbl_top, text="0 dòng đối chiếu",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED
        )
        self.lbl_comp_count.pack(side="right")

        # Treeview Container
        self.tree_container = tk.Frame(self.table_card, bg="#131929")
        self.tree_container.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        col_ids = [c[0] for c in self.col_specs]
        self.tree = ttk.Treeview(
            self.tree_container,
            columns=col_ids,
            show="headings",
            style="Comp.Treeview",
            selectmode="browse"
        )
        self.vsb = tk.Scrollbar(self.tree_container, orient="vertical", command=self.tree.yview)
        self.hsb = tk.Scrollbar(self.tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.vsb.pack(side="right", fill="y")
        self.hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        for cid, width, anchor in self.col_specs:
            hdr_text = self.t("comp_cols").get(cid, cid)
            self.tree.heading(cid, text=hdr_text, anchor=anchor)
            self.tree.column(cid, width=width, minwidth=40, anchor=anchor)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        # Context Menu
        self.ctx_menu = tk.Menu(self, tearoff=0)

        # Empty state hero overlay inside table
        self.empty_overlay = ctk.CTkFrame(self.table_card, fg_color=BG_CARD, corner_radius=14)
        self.empty_overlay.place(relx=0.5, rely=0.55, anchor="center")

        ctk.CTkLabel(self.empty_overlay, text="⚖️", font=("Segoe UI", 48), fg_color="transparent").pack(pady=(20, 8))
        self.lbl_empty_comp_title = ctk.CTkLabel(
            self.empty_overlay, text="Sẵn Sàng Đối Chiếu BOM",
            font=("Segoe UI", 15, "bold"), text_color=TEXT_PRIMARY, fg_color="transparent"
        )
        self.lbl_empty_comp_title.pack(padx=28, pady=(0, 6))

        self.lbl_empty_comp_desc = ctk.CTkLabel(
            self.empty_overlay,
            text="Nạp file Excel đối chiếu và file Working Manual PDF ở trên rồi nhấn '⚡ SO SÁNH NGAY'.",
            font=("Segoe UI", 10), text_color=TEXT_MUTED, fg_color="transparent"
        )
        self.lbl_empty_comp_desc.pack(padx=28, pady=(0, 20))

    # ── Excel Handling ────────────────────────────────────────────────────────
    def _select_excel(self):
        f = filedialog.askopenfilename(
            title="Chọn file Excel đối chiếu",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if f:
            self.load_excel(f)

    def load_excel(self, filepath: str):
        filepath = os.path.normpath(filepath)
        if not os.path.exists(filepath):
            return

        size = os.path.getsize(filepath)
        if size > MAX_FILE_SIZE_BYTES:
            size_str = format_file_size(size)
            msg = self.t("err_file_too_large", name=os.path.basename(filepath), size=size_str, max=f"{MAX_FILE_SIZE_MB} MB")
            messagebox.showwarning(self.t("title_file_too_large"), msg)
            self.app.show_toast(f"⚠️ {os.path.basename(filepath)} ({size_str}) vượt quá {MAX_FILE_SIZE_MB}MB!")
            return

        try:
            self.excel_data = parse_excel_bom(filepath)
            fname = self.excel_data["filename"]
            proc = self.excel_data["process_code"] or "N/A"
            cnt = self.excel_data["total_items"]
            size_str = format_file_size(size)

            self.lbl_excel_name.configure(text=f"📄 {fname} ({size_str})", text_color=TEXT_PRIMARY)
            self.lbl_excel_details.configure(
                text=self.t("excel_info", name=fname, size=size_str, proc=proc, count=cnt)
            )

            # Update sheet selector
            sheets = self.excel_data.get("all_sheets", [])
            if sheets:
                self.opt_sheets.configure(values=sheets)
                self.selected_sheet_str.set(self.excel_data.get("sheet_name", sheets[0]))
                st = "normal" if len(sheets) > 1 else "disabled"
                if hasattr(self, "btn_prev_sheet"):
                    self.btn_prev_sheet.configure(state=st)
                if hasattr(self, "btn_next_sheet"):
                    self.btn_next_sheet.configure(state=st)

            # Attempt smart auto-match with loaded PDF
            self._try_auto_match_page()
            self.app.set_status(f"Đã nạp file Excel '{fname}' ({size_str}, {cnt} linh kiện).")

            # Tự động chạy so sánh nếu đã có file Working Manual
            if self.pdf_path and self.wm_pages_info:
                self._run_comparison()
        except Exception as e:
            messagebox.showerror("Lỗi đọc Excel", f"Không thể đọc file Excel:\n{e}")

    def _prev_sheet(self):
        if not self.excel_data or not self.excel_data.get("all_sheets"):
            return
        sheets = self.excel_data["all_sheets"]
        curr = self.selected_sheet_str.get()
        idx = sheets.index(curr) if curr in sheets else 0
        new_sheet = sheets[(idx - 1) % len(sheets)]
        self.selected_sheet_str.set(new_sheet)
        self._on_sheet_change(new_sheet)

    def _next_sheet(self):
        if not self.excel_data or not self.excel_data.get("all_sheets"):
            return
        sheets = self.excel_data["all_sheets"]
        curr = self.selected_sheet_str.get()
        idx = sheets.index(curr) if curr in sheets else 0
        new_sheet = sheets[(idx + 1) % len(sheets)]
        self.selected_sheet_str.set(new_sheet)
        self._on_sheet_change(new_sheet)

    def _on_sheet_change(self, choice: str):
        if self.excel_data and self.excel_data.get("filepath"):
            try:
                self.excel_data = parse_excel_bom(self.excel_data["filepath"], sheet_name=choice)
                cnt = self.excel_data["total_items"]
                proc = self.excel_data["process_code"] or "N/A"
                fname = self.excel_data["filename"]
                size_str = format_file_size(os.path.getsize(self.excel_data["filepath"]))
                self.lbl_excel_details.configure(
                    text=self.t("excel_info", name=fname, size=size_str, proc=proc, count=cnt)
                )
                self._try_auto_match_page()
                if self.pdf_path and self.wm_pages_info:
                    self._run_comparison()
            except Exception as e:
                messagebox.showerror("Lỗi đọc sheet", f"Lỗi: {e}")

    # ── PDF Manual Handling ───────────────────────────────────────────────────
    def _select_pdf(self):
        f = filedialog.askopenfilename(
            title="Chọn file Working Manual (PDF)",
            filetypes=[("PDF Manual", "*.pdf")]
        )
        if f:
            self.load_pdf(f)

    def load_pdf(self, filepath: str):
        filepath = os.path.normpath(filepath)
        if not os.path.exists(filepath):
            return

        size = os.path.getsize(filepath)
        if size > MAX_FILE_SIZE_BYTES:
            size_str = format_file_size(size)
            msg = self.t("err_file_too_large", name=os.path.basename(filepath), size=size_str, max=f"{MAX_FILE_SIZE_MB} MB")
            messagebox.showwarning(self.t("title_file_too_large"), msg)
            self.app.show_toast(f"⚠️ {os.path.basename(filepath)} ({size_str}) vượt quá {MAX_FILE_SIZE_MB}MB!")
            return

        try:
            self.pdf_path = filepath
            self.wm_pages_info = get_wm_pages_info(filepath)
            fname = os.path.basename(filepath)
            total_pages = len(self.wm_pages_info)
            pwb = next((p["pwb_code"] for p in self.wm_pages_info if p["pwb_code"]), "N/A")
            size_str = format_file_size(size)

            self.lbl_pdf_name.configure(text=f"📄 {fname} ({size_str})", text_color=TEXT_PRIMARY)
            self.lbl_pdf_details.configure(
                text=self.t("pdf_info", name=fname, size=size_str, pwb=pwb, pages=total_pages)
            )

            # Populate page picker dropdown
            page_options = [p["display_text"] for p in self.wm_pages_info]
            if page_options:
                self.opt_pages.configure(values=page_options)
                self.selected_page_str.set(page_options[0])
                self.selected_page_idx = 0
                p0 = self.wm_pages_info[0]
                proc0 = p0.get("process_code") or "Trang 1"
                cnt0 = p0.get("items_count", 0)
                if hasattr(self, "lbl_target_chip"):
                    self.lbl_target_chip.configure(
                        text=f"📑 Đối chiếu: Trang 1/{total_pages} ({proc0} - {cnt0} LK)"
                    )
                st = "normal" if total_pages > 1 else "disabled"
                if hasattr(self, "btn_prev_page"):
                    self.btn_prev_page.configure(state=st)
                if hasattr(self, "btn_next_page"):
                    self.btn_next_page.configure(state=st)

            # Attempt smart auto-match
            self._try_auto_match_page()
            self.app.set_status(f"Đã nạp Working Manual '{fname}' ({total_pages} trang).")

            # Tự động chạy so sánh nếu đã có file Excel
            if self.excel_data:
                self._run_comparison()
        except Exception as e:
            messagebox.showerror("Lỗi đọc PDF", f"Không thể đọc file PDF Working Manual:\n{e}")

    def _prev_page(self):
        if not self.wm_pages_info:
            return
        new_idx = (self.selected_page_idx - 1) % len(self.wm_pages_info)
        self._set_page_by_index(new_idx)

    def _next_page(self):
        if not self.wm_pages_info:
            return
        new_idx = (self.selected_page_idx + 1) % len(self.wm_pages_info)
        self._set_page_by_index(new_idx)

    def _set_page_by_index(self, idx: int):
        if 0 <= idx < len(self.wm_pages_info):
            self.selected_page_idx = idx
            target_text = self.wm_pages_info[idx]["display_text"]
            self.selected_page_str.set(target_text)
            self._on_page_change(target_text)

    def _on_page_change(self, choice: str):
        for idx, p in enumerate(self.wm_pages_info):
            if (p["display_text"] == choice or
                choice.startswith(f"Trang {p['page_num']}/") or
                choice.startswith(f"Trang {p['page_num']}:")):
                self.selected_page_idx = idx
                break

        if self.wm_pages_info and self.selected_page_idx < len(self.wm_pages_info):
            p_info = self.wm_pages_info[self.selected_page_idx]
            proc = p_info.get("process_code") or f"Trang {self.selected_page_idx + 1}"
            cnt = p_info.get("items_count", 0)
            total = len(self.wm_pages_info)
            if hasattr(self, "lbl_target_chip"):
                self.lbl_target_chip.configure(
                    text=f"📑 Đối chiếu: Trang {self.selected_page_idx + 1}/{total} ({proc} - {cnt} LK)"
                )
            self.app.show_toast(f"Đã chọn: Trang {self.selected_page_idx + 1}/{total} ({proc})")
            self.app.set_status(f"Đã chọn trang {self.selected_page_idx + 1}/{total} ({proc} - {cnt} linh kiện) để đối chiếu.")

        # TỰ ĐỘNG CHẠY LẠI SO SÁNH KHI NGƯỜI DÙNG ĐỔI TRANG!
        if self.excel_data and self.pdf_path:
            self._run_comparison()

    def _try_auto_match_page(self):
        """Automatically pre-selects the WM page if its process code matches the Excel file."""
        if not self.excel_data or not self.wm_pages_info:
            return

        target_proc = normalize_key(self.excel_data.get("process_code", ""))
        if not target_proc:
            return

        total_pages = len(self.wm_pages_info)
        for idx, p in enumerate(self.wm_pages_info):
            p_proc = normalize_key(p.get("process_code", ""))
            p_vars = [normalize_key(v) for v in p.get("variants", [])]
            if target_proc and (target_proc in p_proc or p_proc in target_proc or any(target_proc in v for v in p_vars)):
                self.selected_page_idx = idx
                self.selected_page_str.set(p["display_text"])
                if hasattr(self, "lbl_target_chip"):
                    proc_name = p.get("process_code") or f"Trang {idx + 1}"
                    cnt_m = p.get("items_count", 0)
                    self.lbl_target_chip.configure(
                        text=f"📑 Đối chiếu: Trang {idx + 1}/{total_pages} ({proc_name} - {cnt_m} LK)"
                    )
                break

    # ── Comparison Execution ──────────────────────────────────────────────────
    def _run_comparison(self):
        if getattr(self, "_is_comparing", False):
            return

        if not self.excel_data:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file Excel đối chiếu trước!")
            return
        if not self.pdf_path or not self.wm_pages_info:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file Working Manual (PDF) trước!")
            return

        self._is_comparing = True
        self.btn_run_comp.configure(
            state="disabled", fg_color=("#94A3B8", "#1A3A35"),
            text="⚡ Đang đối chiếu AI... (0%)"
        )

        # Clear tree & hide empty overlay
        self.tree.delete(*self.tree.get_children())
        if hasattr(self, "empty_overlay"):
            self.empty_overlay.place_forget()

        # Update Verdict Banner to analyzing state
        self.verdict_banner.configure(fg_color=BG_CARD, border_color="#00F0FF")
        self.lbl_verdict_title.configure(
            text="⚡ VIPQC AI BOM AUDIT ENGINE ĐANG ĐỐI CHIẾU...",
            text_color="#00F0FF"
        )
        self.lbl_verdict_sub.configure(
            text="Đang phân tích cấu trúc Working Manual & đối chiếu ma trận linh kiện với Excel...",
            text_color=TEXT_MUTED
        )

        # Remove previous HUD if any
        if getattr(self, "_comp_hud", None) is not None:
            try:
                self._comp_hud.destroy()
            except Exception:
                pass
            self._comp_hud = None

        # Holographic HUD Card over self.table_card
        is_dark = (self.app.current_theme == "dark")
        self._comp_hud = ctk.CTkFrame(
            self.table_card, fg_color="#0B132B" if is_dark else "#0F172A",
            corner_radius=12, border_width=1.5, border_color="#00F0FF"
        )
        self._comp_hud.place(relx=0.5, rely=0.45, anchor="center")

        hud_inner = ctk.CTkFrame(self._comp_hud, fg_color="transparent")
        hud_inner.pack(padx=24, pady=16)

        # Row 1: Header + Telemetry
        h_row1 = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row1.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(h_row1, text="🔍", font=("Segoe UI", 16)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            h_row1, text="VIPQC AI BOM AUDIT & COMPARISON ENGINE",
            font=("Segoe UI", 11, "bold"), text_color="#00F0FF"
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(
            h_row1, text="60 FPS • SMART MATRIX DIFF",
            font=("Consolas", 9, "bold"), text_color="#38BDF8"
        ).pack(side="right")

        # Row 2: Progress bar + Percentage
        h_row2 = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row2.pack(fill="x", pady=(0, 8))

        comp_prog = ctk.CTkProgressBar(
            h_row2, width=380, height=10, corner_radius=5,
            fg_color="#1E293B", progress_color="#00F0FF"
        )
        comp_prog.set(0.0)
        comp_prog.pack(side="left", padx=(0, 10))

        lbl_comp_pct = ctk.CTkLabel(
            h_row2, text="0%", font=("Consolas", 11, "bold"), text_color="#00F0FF", width=44
        )
        lbl_comp_pct.pack(side="left")

        # Row 3: Live Terminal Log Phase
        lbl_comp_log = ctk.CTkLabel(
            hud_inner, text="[01/05] 📡 Khởi động VIPQC AI Audit Engine & nạp dữ liệu...",
            font=("Consolas", 10), text_color="#E2E8F0", anchor="w"
        )
        lbl_comp_log.pack(fill="x")

        # Background calculation worker
        # Snapshot input data before launching background worker
        excel_data_snap = self.excel_data
        pdf_path_snap = self.pdf_path
        page_idx_snap = self.selected_page_idx

        comp_data = {"res": None, "err": None, "done": False}
        def _bg_worker():
            try:
                target_proc = excel_data_snap.get("process_code", "") if excel_data_snap else ""
                wm_page_data = extract_wm_page_items(pdf_path_snap, page_idx_snap, variant_name=target_proc)
                comp_data["res"] = compare_boms(excel_data_snap, wm_page_data)
            except Exception as ex:
                comp_data["err"] = ex
            finally:
                comp_data["done"] = True

        threading.Thread(target=_bg_worker, daemon=True).start()

        # Animation parameters (3.5 seconds)
        start_time = time.time()
        anim_duration = 3.5

        def _update_comp_anim():
            if not getattr(self, "_is_comparing", False):
                return

            now = time.time()
            elapsed = now - start_time
            t = min(1.0, elapsed / anim_duration)

            comp_prog.set(t)
            pct = int(t * 100)
            lbl_comp_pct.configure(text=f"{pct}%")
            self.btn_run_comp.configure(text=f"⚡ Đang đối chiếu AI... ({pct}%)")

            lang = getattr(self.app, "current_lang", "vi")
            if lang == "zh":
                if t < 0.20:
                    lbl_comp_log.configure(text="[01/05] 📡 启动 VIPQC AI 智能核验引擎并载入数据...")
                elif t < 0.40:
                    lbl_comp_log.configure(text="[02/05] 📑 分析 Working Manual 结构与工序代码坐标...")
                elif t < 0.65:
                    lbl_comp_log.configure(text="[03/05] 🔬 交叉比对物料编码与安装位置 (Excel vs WM)...")
                elif t < 0.85:
                    lbl_comp_log.configure(text="[04/05] ⚡ 核查数量、规格差异及缺失/多余元件...")
                else:
                    lbl_comp_log.configure(text="[05/05] 🎯 汇总智能比对结果并完成最终判定！")
            elif lang == "en":
                if t < 0.20:
                    lbl_comp_log.configure(text="[01/05] 📡 Starting VIPQC AI Audit Engine & ingesting data...")
                elif t < 0.40:
                    lbl_comp_log.configure(text="[02/05] 📑 Analyzing Working Manual structure & Process Code...")
                elif t < 0.65:
                    lbl_comp_log.configure(text="[03/05] 🔬 Cross-referencing Part No & Locations (Excel vs WM)...")
                elif t < 0.85:
                    lbl_comp_log.configure(text="[04/05] ⚡ Auditing Qty, Rating specs & missing/extra items...")
                else:
                    lbl_comp_log.configure(text="[05/05] 🎯 Aggregating comparison verdict & finalizing audit!")
            else:
                if t < 0.20:
                    lbl_comp_log.configure(text="[01/05] 📡 Khởi động VIPQC AI Audit Engine & nạp dữ liệu...")
                elif t < 0.40:
                    lbl_comp_log.configure(text="[02/05] 📑 Phân tích cấu trúc Working Manual & toạ độ Process Code...")
                elif t < 0.65:
                    lbl_comp_log.configure(text="[03/05] 🔬 Đối chiếu ma trận Part No & Vị trí (Excel vs WM)...")
                elif t < 0.85:
                    lbl_comp_log.configure(text="[04/05] ⚡ Rà soát sai lệch Qty, Rating và linh kiện thừa/thiếu...")
                else:
                    lbl_comp_log.configure(text="[05/05] 🎯 Tổng hợp kết quả đối chiếu & hoàn tất đánh giá!")

            if t < 1.0 or not comp_data["done"]:
                self.after(16, _update_comp_anim)
            else:
                self._finish_comparison(comp_data)

        self.after(20, _update_comp_anim)

    def _finish_comparison(self, comp_data: dict):
        self._is_comparing = False

        # Pulse border on table_card
        self.table_card.configure(border_color="#00F0FF", border_width=2)
        self.after(180, lambda: self.table_card.configure(border_color=BORDER_CLR, border_width=1))

        # Cleanup HUD
        if getattr(self, "_comp_hud", None) is not None:
            try:
                self._comp_hud.destroy()
            except Exception:
                pass
            self._comp_hud = None

        self.btn_run_comp.configure(state="normal", fg_color=ACCENT_TEAL, text=self.t("btn_compare_now"))

        if comp_data.get("err"):
            messagebox.showerror("Lỗi đối chiếu", f"Đã xảy ra lỗi khi so sánh:\n{comp_data['err']}")
            return

        res = comp_data.get("res")
        if not res:
            return
        self.comparison_result = res

        # Update Verdict Banner
        if res["is_all_matched"]:
            self.verdict_banner.configure(
                fg_color=("#DCFCE7", "#064E3B"),
                border_color=("#16A34A", "#059669")
            )
            self.lbl_verdict_title.configure(
                text=self.t("verdict_match_title"),
                text_color=("#15803D", "#34D399")
            )
            self.lbl_verdict_sub.configure(
                text=self.t("verdict_match_sub", total=res["total_excel"], proc=res["wm_process"]),
                text_color=("#166534", "#A7F3D0")
            )
        else:
            self.verdict_banner.configure(
                fg_color=("#FEE2E2", "#3D1414"),
                border_color=("#DC2626", "#EF4444")
            )
            if res["count_missing_excel"] > 0 and res["count_mismatched"] == 0 and res["count_missing_wm"] == 0:
                self.lbl_verdict_title.configure(
                    text=self.t("verdict_missing_excel_title", extra=res["count_missing_excel"]),
                    text_color=("#B91C1C", "#F87171")
                )
                self.lbl_verdict_sub.configure(
                    text=self.t("verdict_missing_excel_sub", extra=res["count_missing_excel"], proc=res["wm_process"]),
                    text_color=("#991B1B", "#FECACA")
                )
            else:
                self.lbl_verdict_title.configure(
                    text=self.t("verdict_mismatch_title", total=res["total_discrepancies"]),
                    text_color=("#B91C1C", "#F87171")
                )
                self.lbl_verdict_sub.configure(
                    text=self.t("verdict_mismatch_sub",
                                mismatch=res["count_mismatched"],
                                missing=res["count_missing_wm"],
                                extra=res["count_missing_excel"]),
                    text_color=("#991B1B", "#FECACA")
                )

        # Update Filter Tab counts
        self.filter_tabs.configure(values=[
            self.t("filter_comp_all", n=res["total_rows"]),
            self.t("filter_comp_mismatch", n=res["total_discrepancies"]),
            self.t("filter_comp_match", n=res["count_matched"])
        ])

        # Populate Table
        self._populate_table()

        # Enable Export
        self.btn_export_comp.configure(state="normal")
        if hasattr(self, "empty_overlay"):
            self.empty_overlay.place_forget()

        # Update top badges
        self.app._update_badges_for_compare()
        self.app.set_status(f"Đã đối chiếu xong: {res['count_matched']} khớp, {res['total_discrepancies']} sai lệch.")
        self.app.show_toast(f"✅ Đối chiếu hoàn tất: {res['count_matched']} khớp, {res['total_discrepancies']} sai lệch")

    def _populate_table(self):
        self.tree.delete(*self.tree.get_children())
        if not self.comparison_result:
            return

        rows = self.comparison_result["comparison_rows"]
        q = self.search_query.lower()

        filtered = []
        for r in rows:
            # Filter mode
            if self.filter_mode == "mismatch" and r["status"] == "MATCH":
                continue
            if self.filter_mode == "match" and r["status"] != "MATCH":
                continue

            # Search query
            if q:
                searchable = f"{r['part_no_excel']} {r['part_no_wm']} {r['locations_excel']} {r['locations_wm']} {r['rating_excel']} {r['rating_wm']} {r['diff_summary']}".lower()
                if q not in searchable:
                    continue

            filtered.append(r)

        for idx, r in enumerate(filtered, 1):
            st = r["status"]
            tag = "tag_match" if st == "MATCH" else ("tag_mismatch" if st == "MISMATCH" else ("tag_missing_wm" if st == "MISSING_IN_WM" else "tag_missing_excel"))

            vals = (
                idx,
                r["status_label"],
                r["part_no_excel"],
                r["part_no_wm"],
                r["qty_excel"],
                r["qty_wm"],
                r["locations_excel"],
                r["locations_wm"],
                r["rating_excel"],
                r["diff_summary"]
            )
            self.tree.insert("", "end", values=vals, tags=(tag,))

        self.lbl_comp_count.configure(
            text=f"{len(filtered)}/{len(rows)} dòng đối chiếu"
        )

    def _on_filter_change(self, choice: str):
        if "sai lệch" in choice.lower() or "差异" in choice or "discrepancies" in choice.lower():
            self.filter_mode = "mismatch"
        elif "khớp" in choice.lower() or "匹配" in choice or "matches" in choice.lower():
            self.filter_mode = "match"
        else:
            self.filter_mode = "all"
        self._populate_table()

    def _on_search_key(self, event):
        self.search_query = self.search_comp.get().strip()
        self._populate_table()

    # ── Export & Reset ────────────────────────────────────────────────────────
    def _export_report(self):
        if not self.comparison_result:
            return

        default_name = f"BOM_Comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path = filedialog.asksaveasfilename(
            title="Lưu Báo Cáo Đối Chiếu BOM",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if out_path:
            try:
                export_comparison_excel(self.comparison_result, out_path)
                self.app.show_toast(f"Đã xuất báo cáo: {os.path.basename(out_path)}")
                ans = messagebox.askyesno("Xuất file thành công", f"Đã lưu báo cáo so sánh tại:\n{out_path}\n\nBạn có muốn mở file ngay không?")
                if ans:
                    os.startfile(out_path)
            except Exception as e:
                messagebox.showerror("Lỗi xuất file", f"Không thể lưu file báo cáo:\n{e}")

    def _reset(self):
        if getattr(self, "_is_comparing", False):
            self._is_comparing = False
        if getattr(self, "_comp_hud", None) is not None:
            try:
                self._comp_hud.destroy()
            except Exception:
                pass
            self._comp_hud = None
        self.btn_run_comp.configure(state="normal", fg_color=ACCENT_TEAL, text=self.t("btn_compare_now"))

        self.excel_data = None
        self.pdf_path = None
        self.wm_pages_info.clear()
        self.comparison_result = None
        self.filter_mode = "all"
        self.search_query = ""

        self.lbl_excel_name.configure(text=self.t("excel_no_file"), text_color=TEXT_MUTED)
        self.lbl_excel_details.configure(text="")
        self.opt_sheets.configure(values=["Sheet1"])
        self.selected_sheet_str.set("Sheet1")
        if hasattr(self, "btn_prev_sheet"):
            self.btn_prev_sheet.configure(state="disabled")
        if hasattr(self, "btn_next_sheet"):
            self.btn_next_sheet.configure(state="disabled")

        self.lbl_pdf_name.configure(text=self.t("pdf_no_file"), text_color=TEXT_MUTED)
        self.lbl_pdf_details.configure(text="")
        self.opt_pages.configure(values=["(Chưa nạp PDF - Hãy chọn file PDF trước)"])
        self.selected_page_str.set("(Chưa nạp PDF - Hãy chọn file PDF trước)")
        if hasattr(self, "btn_prev_page"):
            self.btn_prev_page.configure(state="disabled")
        if hasattr(self, "btn_next_page"):
            self.btn_next_page.configure(state="disabled")

        if hasattr(self, "lbl_target_chip"):
            self.lbl_target_chip.configure(text="📑 Trang đối chiếu: (Chưa chọn)")

        self.verdict_banner.configure(fg_color=BG_CARD, border_color=BORDER_CLR)
        self.lbl_verdict_title.configure(text=self.t("verdict_placeholder"), text_color=TEXT_MUTED)
        self.lbl_verdict_sub.configure(text="")

        self.btn_export_comp.configure(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.empty_overlay.place(relx=0.5, rely=0.55, anchor="center")

        self.app._update_badges_for_compare()
        self.app.set_status("Đã làm mới khung đối chiếu.")

    def handle_drop_files(self, files):
        for item in files:
            if isinstance(item, bytes):
                item = item.decode("utf-8", errors="ignore")
            item = os.path.normpath(item)
            ext = os.path.splitext(item)[1].lower()
            if ext in [".xlsx", ".xls"]:
                self.load_excel(item)
            elif ext == ".pdf":
                self.load_pdf(item)

    # ── Interaction & Theme ───────────────────────────────────────────────────
    def _on_double_click(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        vals = self.tree.item(item_id, "values")
        if vals and len(vals) > 2:
            pn = vals[2] if vals[2] != "-" else vals[3]
            self.clipboard_clear()
            self.clipboard_append(pn)
            self.app.show_toast(f"Đã chép: {pn}")

    def _on_right_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        self.tree.selection_set(item_id)
        vals = self.tree.item(item_id, "values")

        self.ctx_menu.delete(0, "end")
        self.ctx_menu.add_command(label=f"📋 Chép Mã LK Excel: {vals[2]}", command=lambda: self._copy_text(vals[2]))
        self.ctx_menu.add_command(label=f"📋 Chép Mã LK WM: {vals[3]}", command=lambda: self._copy_text(vals[3]))
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label=f"📌 Chép Vị Trí Excel: {vals[6]}", command=lambda: self._copy_text(vals[6]))
        self.ctx_menu.add_command(label=f"📌 Chép Vị Trí WM: {vals[7]}", command=lambda: self._copy_text(vals[7]))
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="📝 Chép toàn bộ dòng", command=lambda: self._copy_text("\t".join(str(v) for v in vals)))
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _copy_text(self, txt):
        if txt and txt != "-":
            self.clipboard_clear()
            self.clipboard_append(str(txt))
            self.app.show_toast(f"Đã sao chép: {txt[:25]}…")

    def apply_theme(self, is_dark: bool):
        bg_card = BG_CARD[1] if is_dark else BG_CARD[0]
        bg_surf = BG_SURFACE[1] if is_dark else BG_SURFACE[0]
        b_clr = BORDER_CLR[1] if is_dark else BORDER_CLR[0]
        txt_prim = TEXT_PRIMARY[1] if is_dark else TEXT_PRIMARY[0]
        txt_mut = TEXT_MUTED[1] if is_dark else TEXT_MUTED[0]

        self.card_excel.configure(fg_color=bg_card, border_color=b_clr)
        self.card_pdf.configure(fg_color=bg_card, border_color=b_clr)
        self.table_card.configure(fg_color=bg_card, border_color=b_clr)
        self.empty_overlay.configure(fg_color=bg_card)
        self.tree_container.configure(bg=bg_card)

        if hasattr(self, "c_ex_bot"):
            self.c_ex_bot.configure(fg_color=bg_surf)
        if hasattr(self, "c_pdf_bot"):
            self.c_pdf_bot.configure(fg_color=bg_surf)
        if hasattr(self, "chip_target_page"):
            self.chip_target_page.configure(fg_color=bg_card, border_color=b_clr)

        for btn in [getattr(self, "btn_prev_page", None),
                    getattr(self, "btn_next_page", None),
                    getattr(self, "btn_prev_sheet", None),
                    getattr(self, "btn_next_sheet", None)]:
            if btn:
                btn.configure(fg_color=bg_card, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_prim)

        # Style Treeview
        style = ttk.Style()
        tree_bg = "#0B0F1A" if is_dark else "#FFFFFF"
        tree_fg = "#E8EDF5" if is_dark else "#0F172A"
        hdr_bg = "#1C2438" if is_dark else "#E2E8F0"
        hdr_fg = "#00C9A7" if is_dark else "#0D9488"
        sel_bg = "#2D4A6E" if is_dark else "#BFDBFE"
        sel_fg = "#FFFFFF" if is_dark else "#1E3A8A"

        style.configure("Comp.Treeview",
                        background=tree_bg,
                        fieldbackground=tree_bg,
                        foreground=tree_fg,
                        rowheight=28,
                        font=("Segoe UI", 10),
                        relief="flat")
        style.configure("Comp.Treeview.Heading",
                        background=hdr_bg,
                        foreground=hdr_fg,
                        font=("Segoe UI", 10, "bold"),
                        relief="flat",
                        padding=5)
        style.map("Comp.Treeview",
                  background=[("selected", sel_bg)],
                  foreground=[("selected", sel_fg)])

        # Tags colors
        self.tree.tag_configure("tag_match", background="#062E1E" if is_dark else "#DCFCE7")
        self.tree.tag_configure("tag_mismatch", background="#3B2A05" if is_dark else "#FEF3C7")
        self.tree.tag_configure("tag_missing_wm", background="#381010" if is_dark else "#FEE2E2")
        self.tree.tag_configure("tag_missing_excel", background="#451820" if is_dark else "#FFE4E6")

    def update_language(self):
        self.lbl_card_ex_title.configure(text=self.t("card_excel_title"))
        self.btn_choose_excel.configure(text=self.t("btn_choose_excel"))
        if not self.excel_data:
            self.lbl_excel_name.configure(text=self.t("excel_no_file"))

        self.lbl_card_pdf_title.configure(text=self.t("card_pdf_title"))
        self.lbl_page_picker.configure(text=self.t("lbl_select_wm_page"))
        self.btn_choose_pdf.configure(text=self.t("btn_choose_pdf"))
        if not self.pdf_path:
            self.lbl_pdf_name.configure(text=self.t("pdf_no_file"))

        self.btn_run_comp.configure(text=self.t("btn_compare_now"))
        self.btn_export_comp.configure(text=self.t("btn_export_comparison"))
        self.btn_reset_comp.configure(text=self.t("btn_reset_compare"))

        # Column headings
        for cid, width, anchor in self.col_specs:
            hdr_text = self.t("comp_cols").get(cid, cid)
            self.tree.heading(cid, text=hdr_text)


# ──────────────────────────────────────────────────────────────────────────────
# MODAL / FULLSCREEN PCB DRAWING INSPECTION WINDOW
# ──────────────────────────────────────────────────────────────────────────────
class DrawingFullScreenWindow(ctk.CTkToplevel):
    """
    Dedicated maximized / full-screen window for high-resolution PCB drawing inspection,
    pan/zoom navigation, curtain swipe comparison, and component-to-component stepping.
    """
    def __init__(self, parent, app, pdf_a_path: str, pdf_b_path: str,
                 comp_result: dict, drawing_catalog: list, active_item: dict | None = None,
                 current_stage_id: str = "ALL", active_page_a: int = 1, active_page_b: int = 1,
                 split_ratio: float = 0.5, view_mode: str = "curtain", zoom_level: float = 1.0,
                 rotation_angle: int = 0,
                 on_close_cb=None, on_select_item_cb=None):
        super().__init__(parent)
        self.app = app
        self.parent = parent
        self.pdf_a_path = pdf_a_path
        self.pdf_b_path = pdf_b_path
        self.comp_result = comp_result or {}
        self.drawing_catalog = drawing_catalog or []
        self.active_item = active_item
        self.current_stage_id = current_stage_id
        self.active_page_a = active_page_a
        self.active_page_b = active_page_b
        self.split_ratio = split_ratio
        self.view_mode = view_mode
        self.zoom_level = zoom_level
        self.rotation_angle = rotation_angle
        self.on_close_cb = on_close_cb
        self.on_select_item_cb = on_select_item_cb

        self.current_meta: dict = {}
        self.tk_canvas_img = None
        self.canvas_img_id = None
        self._curtain_render_pending: bool = False
        self.dragging_curtain: bool = False
        self.draw_offset_x: int = 0
        self.draw_offset_y: int = 0
        self.pan_offset_x: int = 0
        self.pan_offset_y: int = 0
        self._pan_start_x: int = 0
        self._pan_start_y: int = 0
        self._pan_init_offset_x: int = 0
        self._pan_init_offset_y: int = 0

        self.title("📐 Soi Chi Tiết Bản Vẽ PCB Toàn Màn Hình - VIPQC AI")
        self.geometry("1440x900")
        self.minsize(1024, 650)
        try:
            self.state("zoomed")
        except Exception:
            pass

        # Match theme
        is_dark = (self.app.current_theme == "dark")
        self.configure(fg_color=BG_DEEP[1] if is_dark else BG_DEEP[0])

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.bind("<Escape>", lambda e: self._close())
        self.bind("<Left>", lambda e: self._prev_diff())
        self.bind("<Right>", lambda e: self._next_diff())
        self.bind("<r>", lambda e: self._rotate_drawing())
        self.bind("<R>", lambda e: self._rotate_drawing())

        self._build_ui()
        self.focus_force()

        # Render initial view
        self.after(50, lambda: self._render_drawing(center_ref=True))

    def t(self, key: str, **kwargs) -> str:
        return self.app.t(key, **kwargs)

    def _build_ui(self):
        is_dark = (self.app.current_theme == "dark")
        bar_bg = BG_CARD[1] if is_dark else BG_CARD[0]
        surf_bg = BG_SURFACE[1] if is_dark else BG_SURFACE[0]
        bdr_clr = BORDER_CLR[1] if is_dark else BORDER_CLR[0]
        txt_prim = TEXT_PRIMARY[1] if is_dark else TEXT_PRIMARY[0]
        txt_mut = TEXT_MUTED[1] if is_dark else TEXT_MUTED[0]

        # ── Top Bar ──────────────────────────────────────────────────────────
        self.top_bar = ctk.CTkFrame(self, fg_color=bar_bg, corner_radius=0, height=52,
                                    border_width=1, border_color=bdr_clr)
        self.top_bar.pack(side="top", fill="x")
        self.top_bar.pack_propagate(False)

        top_inner = ctk.CTkFrame(self.top_bar, fg_color="transparent")
        top_inner.pack(fill="both", expand=True, padx=12, pady=8)

        # Title & Active Ref Chip
        self.lbl_title = ctk.CTkLabel(
            top_inner, text="📐 BẢN VẼ PCB TOÀN MÀN HÌNH",
            font=("Segoe UI", 12, "bold"), text_color=ACCENT_TEAL
        )
        self.lbl_title.pack(side="left", padx=(0, 10))

        self.chip_active_ref = ctk.CTkLabel(
            top_inner, text="📍 Đang soi: ---",
            font=("Segoe UI", 10, "bold"), text_color=txt_prim,
            fg_color=surf_bg, corner_radius=6, padx=10, pady=3
        )
        self.chip_active_ref.pack(side="left", padx=(0, 14))

        # Stage Picker
        ctk.CTkLabel(top_inner, text="Bản vẽ:", font=("Segoe UI", 10, "bold"), text_color=txt_prim).pack(side="left", padx=(0, 4))
        cat_labels = [c["label"] for c in self.drawing_catalog] if self.drawing_catalog else ["🌐 Tất Cả Công Đoạn"]
        self.opt_stage_picker = ctk.CTkOptionMenu(
            top_inner, values=cat_labels,
            font=("Segoe UI", 9, "bold"), height=30, width=170,
            fg_color=surf_bg, button_color=ACCENT_TEAL, button_hover_color="#0F766E",
            text_color=txt_prim, command=self._on_stage_change
        )
        curr_label = next((c["label"] for c in self.drawing_catalog if c["id"] == self.current_stage_id), cat_labels[0])
        self.opt_stage_picker.set(curr_label)
        self.opt_stage_picker.pack(side="left", padx=(0, 10))

        # View Mode Toggle
        self.seg_view_mode = ctk.CTkSegmentedButton(
            top_inner, values=["↔️ Kéo Màn", "📄 Model A", "📄 Model B"],
            corner_radius=8, height=30, font=("Segoe UI", 9, "bold"),
            fg_color=surf_bg, selected_color=ACCENT_TEAL, selected_hover_color=ACCENT_TEAL,
            unselected_color=surf_bg, unselected_hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0],
            text_color=txt_prim, command=self._on_view_mode_toggle
        )
        init_mode = "↔️ Kéo Màn" if self.view_mode == "curtain" else ("📄 Model A" if self.view_mode == "a" else "📄 Model B")
        self.seg_view_mode.set(init_mode)
        self.seg_view_mode.pack(side="left", padx=(0, 10))

        # Zoom Controls
        self.btn_zoom_out = ctk.CTkButton(
            top_inner, text="➖", width=30, height=28, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_prim,
            font=("Segoe UI", 10, "bold"), command=self._zoom_out
        )
        self.btn_zoom_out.pack(side="left", padx=(0, 2))

        self.lbl_zoom = ctk.CTkLabel(
            top_inner, text=f"{int(self.zoom_level * 100)}%", width=46,
            font=("Segoe UI", 10, "bold"), text_color=txt_prim
        )
        self.lbl_zoom.pack(side="left", padx=(0, 2))

        self.btn_zoom_in = ctk.CTkButton(
            top_inner, text="➕", width=30, height=28, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_prim,
            font=("Segoe UI", 10, "bold"), command=self._zoom_in
        )
        self.btn_zoom_in.pack(side="left", padx=(0, 4))

        self.btn_zoom_100 = ctk.CTkButton(
            top_inner, text="1:1", width=34, height=28, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_mut,
            font=("Segoe UI", 9), command=self._zoom_100
        )
        self.btn_zoom_100.pack(side="left", padx=(0, 4))

        self.btn_zoom_fit = ctk.CTkButton(
            top_inner, text="📐 Vừa Màn", width=68, height=28, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_mut,
            font=("Segoe UI", 9), command=self._zoom_fit
        )
        self.btn_zoom_fit.pack(side="left", padx=(0, 4))

        rot_lbl = f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else "🔄 Xoay"
        self.btn_rotate = ctk.CTkButton(
            top_inner, text=rot_lbl, width=64, height=28, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_prim,
            font=("Segoe UI", 9, "bold"), command=self._rotate_drawing
        )
        self.btn_rotate.pack(side="left", padx=(0, 10))

        # Curtain Swipe Slider (Manual Drag & Slider Only - No Auto Wipe)
        self.curtain_ctrl_frame = ctk.CTkFrame(top_inner, fg_color="transparent")
        self.curtain_ctrl_frame.pack(side="left", fill="x", expand=True, padx=(0, 12))

        self.lbl_curtain_val = ctk.CTkLabel(
            self.curtain_ctrl_frame, text=f"Màn: {int(self.split_ratio * 100)}%",
            font=("Segoe UI", 9, "bold"), text_color=txt_mut, width=58
        )
        self.lbl_curtain_val.pack(side="left", padx=(0, 4))

        self.slider_curtain = ctk.CTkSlider(
            self.curtain_ctrl_frame, from_=0.0, to=1.0, number_of_steps=100, height=14,
            fg_color=surf_bg, progress_color=ACCENT_TEAL, button_color=ACCENT_TEAL,
            button_hover_color="#00A88C", command=self._on_curtain_slider
        )
        self.slider_curtain.set(self.split_ratio)
        self.slider_curtain.pack(side="left", fill="x", expand=True)

        if self.view_mode != "curtain":
            self.curtain_ctrl_frame.pack_forget()

        # Close Button
        self.btn_close = ctk.CTkButton(
            top_inner, text="❌ Đóng (Esc)", width=96, height=30, corner_radius=8,
            fg_color=("#EF4444", "#DC2626"), hover_color=("#DC2626", "#B91C1C"),
            font=("Segoe UI", 10, "bold"), text_color="#FFFFFF", command=self._close
        )
        self.btn_close.pack(side="right")

        # ── Bottom Bar: Info & Step Navigation ────────────────────────────────
        self.bottom_bar = ctk.CTkFrame(self, fg_color=bar_bg, corner_radius=0, height=52,
                                       border_width=1, border_color=bdr_clr)
        self.bottom_bar.pack(side="bottom", fill="x")
        self.bottom_bar.pack_propagate(False)

        bot_inner = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")
        bot_inner.pack(fill="both", expand=True, padx=12, pady=6)

        # Diff stepper controls
        self.btn_prev_diff = ctk.CTkButton(
            bot_inner, text="◀ Sai Lệch Trước", width=120, height=32, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_prim,
            font=("Segoe UI", 9, "bold"), command=self._prev_diff
        )
        self.btn_prev_diff.pack(side="left", padx=(0, 6))

        self.lbl_diff_counter = ctk.CTkLabel(
            bot_inner, text="Mục: -- / --", font=("Segoe UI", 10, "bold"),
            text_color=ACCENT_TEAL, width=80
        )
        self.lbl_diff_counter.pack(side="left", padx=(0, 6))

        self.btn_next_diff = ctk.CTkButton(
            bot_inner, text="Sai Lệch Tiếp ▶", width=120, height=32, corner_radius=6,
            fg_color=surf_bg, hover_color=BG_HOVER[1] if is_dark else BG_HOVER[0], text_color=txt_prim,
            font=("Segoe UI", 9, "bold"), command=self._next_diff
        )
        self.btn_next_diff.pack(side="left", padx=(0, 16))

        # Detail text info
        self.lbl_det_info = ctk.CTkLabel(
            bot_inner, text="Chọn linh kiện để xem thông số chi tiết",
            font=("Segoe UI", 10), text_color=txt_prim, anchor="w"
        )
        self.lbl_det_info.pack(side="left", fill="x", expand=True)

        self._update_item_display()

        # ── Middle Canvas ─────────────────────────────────────────────────────
        self.canvas_frame = ctk.CTkFrame(self, fg_color="#0B0F1A")
        self.canvas_frame.pack(side="top", fill="both", expand=True)

        self.canvas_vsb = tk.Scrollbar(self.canvas_frame, orient="vertical")
        self.canvas_hsb = tk.Scrollbar(self.canvas_frame, orient="horizontal")

        self.canvas = tk.Canvas(
            self.canvas_frame, bg="#0B0F1A", highlightthickness=0,
            xscrollcommand=self.canvas_hsb.set, yscrollcommand=self.canvas_vsb.set
        )
        self.canvas_vsb.config(command=self.canvas.yview)
        self.canvas_hsb.config(command=self.canvas.xview)

        self.canvas_vsb.pack(side="right", fill="y")
        self.canvas_hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Bind canvas events
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_b1_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_b1_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_b1_release)
        self.canvas.bind("<ButtonPress-2>", self._on_canvas_b2_press)
        self.canvas.bind("<B2-Motion>", self._on_canvas_b2_motion)
        self.canvas.bind("<MouseWheel>", self._on_canvas_mousewheel)
        self.canvas.bind("<Motion>", self._on_canvas_mouse_motion)

    # ── State & Selection ────────────────────────────────────────────────────
    def _get_current_diff_idx(self) -> int:
        diff_items = self.comp_result.get("diff_items", []) if self.comp_result else []
        if not diff_items or not self.active_item:
            return 0
        ref = self.active_item.get("ref_des")
        stage = self.active_item.get("stage")
        for i, it in enumerate(diff_items):
            if it.get("ref_des") == ref and it.get("stage") == stage:
                return i
        return 0

    def _prev_diff(self):
        diff_items = self.comp_result.get("diff_items", []) if self.comp_result else []
        if not diff_items:
            return
        curr = self._get_current_diff_idx()
        new_idx = (curr - 1) % len(diff_items)
        self._select_diff_by_idx(new_idx)

    def _next_diff(self):
        diff_items = self.comp_result.get("diff_items", []) if self.comp_result else []
        if not diff_items:
            return
        curr = self._get_current_diff_idx()
        new_idx = (curr + 1) % len(diff_items)
        self._select_diff_by_idx(new_idx)

    def _select_diff_by_idx(self, idx: int):
        diff_items = self.comp_result.get("diff_items", []) if self.comp_result else []
        if 0 <= idx < len(diff_items):
            self.active_item = diff_items[idx]
            stage = self.active_item.get("stage", "")
            norm_stage = normalize_process_stage(stage)
            cat_match = next((c for c in self.drawing_catalog if c["id"] != "ALL" and normalize_process_stage(c["stage"]) == norm_stage), None)
            if cat_match:
                self.active_page_a = cat_match["page_idx_a"]
                self.active_page_b = cat_match["page_idx_b"]
            self._update_item_display()
            self._render_drawing(center_ref=True)
            if self.on_select_item_cb:
                self.on_select_item_cb(self.active_item)

    def set_active_item(self, item: dict):
        self.active_item = item
        stage = item.get("stage", "")
        norm_stage = normalize_process_stage(stage)
        cat_match = next((c for c in self.drawing_catalog if c["id"] != "ALL" and normalize_process_stage(c["stage"]) == norm_stage), None)
        if cat_match:
            self.active_page_a = cat_match["page_idx_a"]
            self.active_page_b = cat_match["page_idx_b"]
        self._update_item_display()
        self._render_drawing(center_ref=True)

    def _update_item_display(self):
        diff_items = self.comp_result.get("diff_items", []) if self.comp_result else []
        total_diffs = len(diff_items)
        curr_idx = self._get_current_diff_idx() + 1 if total_diffs > 0 else 0
        self.lbl_diff_counter.configure(text=f"Mục: {curr_idx} / {total_diffs}")

        if not self.active_item:
            self.chip_active_ref.configure(text="📍 Đang soi: ---")
            self.lbl_det_info.configure(text="Chưa chọn linh kiện để soi bản vẽ.")
            return

        it = self.active_item
        ref = it.get("ref_des", "")
        stage = it.get("stage", "SMT")
        status = it.get("status", "CHANGED")
        stat_icon = "🟢 LẮP THÊM" if status == "ADDED" else ("🔴 BỎ BỚT" if status == "REMOVED" else ("🟡 ĐỔI TRỊ SỐ" if status == "CHANGED" else "⚪ TRÙNG KHỚP"))

        self.chip_active_ref.configure(text=f"📍 {ref}  |  {stage}  |  {stat_icon}")
        info_txt = (
            f"Vị trí: {ref}  •  Công đoạn: {stage}  •  {stat_icon}  |  "
            f"Model A: {it.get('part_a')} ({it.get('rating_a')})  |  "
            f"Model B: {it.get('part_b')} ({it.get('rating_b')})  |  "
            f"Ghi chú: {it.get('note')}"
        )
        self.lbl_det_info.configure(text=info_txt)

    # ── Controls & Zoom ──────────────────────────────────────────────────────
    def _on_stage_change(self, choice: str):
        matched = next((c for c in self.drawing_catalog if c["label"] == choice), None)
        if matched:
            self.canvas_img_id = None
            self.current_stage_id = matched["id"]
            self.active_page_a = matched["page_idx_a"]
            self.active_page_b = matched["page_idx_b"]
            self._render_drawing(center_ref=False)

    def _on_view_mode_toggle(self, choice: str):
        self.canvas_img_id = None
        c_low = choice.lower()
        if "kéo" in c_low or "swipe" in c_low:
            self.view_mode = "curtain"
            self.curtain_ctrl_frame.pack(side="left", fill="x", expand=True, padx=(0, 12))
        elif "model a" in c_low:
            self.view_mode = "a"
            self.curtain_ctrl_frame.pack_forget()
        else:
            self.view_mode = "b"
            self.curtain_ctrl_frame.pack_forget()
        self._render_drawing(center_ref=False)

    def _zoom_in(self):
        self.canvas_img_id = None
        self.zoom_level = min(5.0, round(self.zoom_level * 1.25, 2))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing(center_ref=False)

    def _zoom_out(self):
        self.canvas_img_id = None
        self.zoom_level = max(0.35, round(self.zoom_level / 1.25, 2))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing(center_ref=False)

    def _zoom_100(self):
        self.canvas_img_id = None
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.zoom_level = 1.0
        self.lbl_zoom.configure(text="100%")
        self._render_drawing(center_ref=False)

    def _rotate_drawing(self):
        self.canvas_img_id = None
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.rotation_angle = (self.rotation_angle + 90) % 360
        self.btn_rotate.configure(text=f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else "🔄 Xoay")
        self._zoom_fit()

    def _zoom_fit(self):
        self.canvas_img_id = None
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.update_idletasks()
        cw = max(200, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        base_w, base_h = (792.0, 612.0) if (self.rotation_angle % 180 == 90) else (612.0, 792.0)
        fit_zoom = min(cw / base_w, ch / base_h)
        self.zoom_level = max(0.35, round(fit_zoom, 2))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing(center_ref=False)

    def _request_fast_curtain_render(self):
        if not getattr(self, "_curtain_render_pending", False):
            self._curtain_render_pending = True
            self.after_idle(self._execute_fast_curtain_render)

    def _execute_fast_curtain_render(self):
        self._curtain_render_pending = False
        self._render_drawing(center_ref=False)

    def _on_curtain_slider(self, val):
        self.split_ratio = float(val)
        self.lbl_curtain_val.configure(text=f"Màn: {int(self.split_ratio * 100)}%")
        self._request_fast_curtain_render()

    def _update_scrollbars(self):
        cw = max(200, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        draw_x0 = getattr(self, "draw_offset_x", 0)
        draw_y0 = getattr(self, "draw_offset_y", 0)
        mw = self.current_meta.get("width", cw)
        mh = self.current_meta.get("height", ch)
        min_x = min(0, draw_x0 + self.pan_offset_x)
        min_y = min(0, draw_y0 + self.pan_offset_y)
        max_x = max(cw, draw_x0 + self.pan_offset_x + mw)
        max_y = max(ch, draw_y0 + self.pan_offset_y + mh)
        self.canvas.config(scrollregion=(min_x, min_y, max_x, max_y))

    # ── Canvas Interactive Handlers ──────────────────────────────────────────
    def _on_canvas_mouse_motion(self, event):
        split_x = self.current_meta.get("split_x")
        draw_x0 = getattr(self, "draw_offset_x", 0)
        curtain_x = draw_x0 + self.pan_offset_x + (split_x if split_x is not None else -9999)
        if self.view_mode == "curtain" and split_x is not None and abs(event.x - curtain_x) < 30:
            self.canvas.config(cursor="sb_h_double_arrow")
        else:
            self.canvas.config(cursor="")

    def _on_canvas_b1_press(self, event):
        split_x = self.current_meta.get("split_x")
        draw_x0 = getattr(self, "draw_offset_x", 0)
        curtain_x = draw_x0 + self.pan_offset_x + (split_x if split_x is not None else -9999)
        if self.view_mode == "curtain" and split_x is not None and abs(event.x - curtain_x) < 30:
            self.dragging_curtain = True
            self.canvas.config(cursor="sb_h_double_arrow")
        else:
            self.dragging_curtain = False
            self.canvas.config(cursor="fleur")
            self._pan_start_x = event.x
            self._pan_start_y = event.y
            self._pan_init_offset_x = self.pan_offset_x
            self._pan_init_offset_y = self.pan_offset_y

    def _on_canvas_b1_motion(self, event):
        if self.dragging_curtain:
            draw_x0 = getattr(self, "draw_offset_x", 0)
            w = max(1, self.current_meta.get("width", 1))
            ratio = max(0.0, min(1.0, (event.x - (draw_x0 + self.pan_offset_x)) / float(w)))
            self.split_ratio = ratio
            self.slider_curtain.set(ratio)
            self.lbl_curtain_val.configure(text=f"Màn: {int(ratio * 100)}%")
            self._request_fast_curtain_render()
        else:
            dx = event.x - self._pan_start_x
            dy = event.y - self._pan_start_y
            self.pan_offset_x = self._pan_init_offset_x + dx
            self.pan_offset_y = self._pan_init_offset_y + dy
            draw_x0 = getattr(self, "draw_offset_x", 0)
            draw_y0 = getattr(self, "draw_offset_y", 0)
            if getattr(self, "canvas_img_id", None) is not None:
                self.canvas.coords(self.canvas_img_id, draw_x0 + self.pan_offset_x, draw_y0 + self.pan_offset_y)
            self._update_scrollbars()

    def _on_canvas_b1_release(self, event):
        if self.dragging_curtain:
            self._render_drawing(center_ref=False)
        self.dragging_curtain = False
        self.canvas.config(cursor="")

    def _on_canvas_b2_press(self, event):
        self.canvas.config(cursor="fleur")
        self._pan_start_x = event.x
        self._pan_start_y = event.y
        self._pan_init_offset_x = self.pan_offset_x
        self._pan_init_offset_y = self.pan_offset_y

    def _on_canvas_b2_motion(self, event):
        dx = event.x - self._pan_start_x
        dy = event.y - self._pan_start_y
        self.pan_offset_x = self._pan_init_offset_x + dx
        self.pan_offset_y = self._pan_init_offset_y + dy
        draw_x0 = getattr(self, "draw_offset_x", 0)
        draw_y0 = getattr(self, "draw_offset_y", 0)
        if getattr(self, "canvas_img_id", None) is not None:
            self.canvas.coords(self.canvas_img_id, draw_x0 + self.pan_offset_x, draw_y0 + self.pan_offset_y)
        self._update_scrollbars()

    def _on_canvas_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    def _render_drawing(self, center_ref: bool = False):
        if not self.pdf_a_path and not self.pdf_b_path:
            return

        try:
            act_ref = self.active_item.get("ref_des") if self.active_item else None
            all_diffs = self.comp_result.get("diff_items", []) if self.comp_result else []

            res_img, meta = render_curtain_drawing_pair(
                pdf_path_a=self.pdf_a_path,
                pdf_path_b=self.pdf_b_path,
                page_idx_a=self.active_page_a,
                page_idx_b=self.active_page_b,
                zoom=self.zoom_level,
                split_ratio=self.split_ratio,
                mode=self.view_mode,
                diff_items=all_diffs,
                active_ref=act_ref,
                rotation=self.rotation_angle
            )
            self.current_meta = meta

            cw = max(200, self.canvas.winfo_width())
            ch = max(200, self.canvas.winfo_height())
            draw_x0 = max(0, (cw - meta["width"]) // 2)
            draw_y0 = max(0, (ch - meta["height"]) // 2)
            self.draw_offset_x = draw_x0
            self.draw_offset_y = draw_y0

            if center_ref and meta.get("active_center"):
                cx, cy = meta["active_center"]
                self.pan_offset_x = (cw // 2) - (draw_x0 + cx)
                self.pan_offset_y = (ch // 2) - (draw_y0 + cy)

            img_x = draw_x0 + self.pan_offset_x
            img_y = draw_y0 + self.pan_offset_y

            self.tk_canvas_img = ImageTk.PhotoImage(res_img)
            if getattr(self, "canvas_img_id", None) is not None:
                self.canvas.coords(self.canvas_img_id, img_x, img_y)
                self.canvas.itemconfig(self.canvas_img_id, image=self.tk_canvas_img)
            else:
                self.canvas.delete("all")
                self.canvas_img_id = self.canvas.create_image(img_x, img_y, image=self.tk_canvas_img, anchor="nw", tags="drawing")
            self._update_scrollbars()
        except Exception as e:
            print(f"Error in fullscreen rendering: {e}")

    def _close(self):
        if hasattr(self.parent, "rotation_angle"):
            self.parent.rotation_angle = self.rotation_angle
            if hasattr(self.parent, "btn_rotate"):
                self.parent.btn_rotate.configure(text=f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else "🔄 Xoay")
        if self.on_close_cb:
            self.on_close_cb()
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3: MODEL SERIES COMPARISON & PCB DRAWING INSPECTOR
# ──────────────────────────────────────────────────────────────────────────────
class ModelCompareView(ctk.CTkFrame):
    """
    Compares two Working Manual PDFs (different revisions or series of a model),
    identifies Added, Removed, and Value-Changed components, and displays
    an interactive PCB drawing inspector with spotlight zoom on the selected component.
    """
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.pdf_a_path: str | None = None
        self.pdf_b_path: str | None = None
        self.comp_result: dict | None = None
        self.active_item: dict | None = None
        self.drawing_catalog: list[dict] = []
        self.current_stage_id: str = "ALL"
        self.active_page_a: int = 1
        self.active_page_b: int = 1
        self.view_mode: str = "curtain"  # "curtain", "a", "b"
        self.display_mode: str = "split" # "table", "drawing", "split"
        self.split_ratio: float = 0.5
        self.zoom_level: float = 1.0
        self.rotation_angle: int = 0
        self.current_meta: dict = {}
        self.tk_canvas_img = None
        self.canvas_img_id = None
        self._curtain_render_pending: bool = False
        self.dragging_curtain: bool = False
        self.draw_offset_x: int = 0
        self.draw_offset_y: int = 0
        self.pan_offset_x: int = 0
        self.pan_offset_y: int = 0
        self._pan_start_x: int = 0
        self._pan_start_y: int = 0
        self._pan_init_offset_x: int = 0
        self._pan_init_offset_y: int = 0
        self.current_filter: str = "diff_all"
        self.search_query: str = ""
        self.active_ctk_image: ctk.CTkImage | None = None
        self.fullscreen_window: DrawingFullScreenWindow | None = None
        self._is_scanning: bool = False
        self._scan_hud = None

        self.col_specs = [
            ("stt",      45,  "center"),
            ("ref_des",  85,  "center"),
            ("stage",    80,  "center"),
            ("status",   110, "center"),
            ("part_a",   150, "w"),
            ("rating_a", 130, "w"),
            ("part_b",   150, "w"),
            ("rating_b", 130, "w"),
            ("note",     190, "w"),
        ]

        self._build_ui()

    def t(self, key: str, **kwargs) -> str:
        return self.app.t(key, **kwargs)

    def _build_ui(self):
        # ── 1. BOTTOM DOCKED: Model A Card, Model B Card, and Compare Box Underneath ──
        self.bottom_dock = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_dock.pack(side="bottom", fill="x", pady=(4, 6))

        # --- Row 1: Model A Card (left 50%) and Model B Card (right 50%) ---
        self.row_models = ctk.CTkFrame(self.bottom_dock, fg_color="transparent")
        self.row_models.pack(fill="x", pady=(0, 6))
        self.cards_row = self.row_models  # Backwards-compatible alias

        # Card Model A (Left 50%)
        self.card_model_a = ctk.CTkFrame(self.row_models, fg_color=BG_CARD, corner_radius=12,
                                         border_width=1, border_color=BORDER_CLR)
        self.card_model_a.pack(side="left", fill="both", expand=True, padx=(0, 4))

        inner_a = ctk.CTkFrame(self.card_model_a, fg_color="transparent")
        inner_a.pack(fill="both", expand=True, padx=14, pady=8)

        self.lbl_card_a_title = ctk.CTkLabel(
            inner_a, text=self.t("card_model_a_title"),
            font=("Segoe UI", 11, "bold"), text_color=ACCENT_BLUE, anchor="w"
        )
        self.lbl_card_a_title.pack(fill="x")

        row_a = ctk.CTkFrame(inner_a, fg_color="transparent")
        row_a.pack(fill="x", pady=(4, 0))

        self.btn_choose_model_a = ctk.CTkButton(
            row_a, text=self.t("btn_choose_model_a"),
            font=("Segoe UI", 10, "bold"), height=30, corner_radius=8,
            fg_color=ACCENT_BLUE, hover_color=("#1D4ED8", "#2563EB"),
            command=self._select_pdf_a
        )
        self.btn_choose_model_a.pack(side="left", padx=(0, 10))

        box_a = ctk.CTkFrame(row_a, fg_color="transparent")
        box_a.pack(side="left", fill="x", expand=True)

        self.lbl_model_a_name = ctk.CTkLabel(
            box_a, text=self.t("model_a_no_file"),
            font=("Segoe UI", 10, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.lbl_model_a_name.pack(fill="x")

        self.lbl_model_a_sub = ctk.CTkLabel(
            box_a, text="",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_model_a_sub.pack(fill="x")

        # Card Model B (Right 50%)
        self.card_model_b = ctk.CTkFrame(self.row_models, fg_color=BG_CARD, corner_radius=12,
                                         border_width=1, border_color=BORDER_CLR)
        self.card_model_b.pack(side="left", fill="both", expand=True, padx=(4, 0))

        inner_b = ctk.CTkFrame(self.card_model_b, fg_color="transparent")
        inner_b.pack(fill="both", expand=True, padx=14, pady=8)

        self.lbl_card_b_title = ctk.CTkLabel(
            inner_b, text=self.t("card_model_b_title"),
            font=("Segoe UI", 11, "bold"), text_color=ACCENT_TEAL, anchor="w"
        )
        self.lbl_card_b_title.pack(fill="x")

        row_b = ctk.CTkFrame(inner_b, fg_color="transparent")
        row_b.pack(fill="x", pady=(4, 0))

        self.btn_choose_model_b = ctk.CTkButton(
            row_b, text=self.t("btn_choose_model_b"),
            font=("Segoe UI", 10, "bold"), height=30, corner_radius=8,
            fg_color=ACCENT_TEAL, hover_color=("#0F766E", "#00A88C"),
            command=self._select_pdf_b
        )
        self.btn_choose_model_b.pack(side="left", padx=(0, 10))

        box_b = ctk.CTkFrame(row_b, fg_color="transparent")
        box_b.pack(side="left", fill="x", expand=True)

        self.lbl_model_b_name = ctk.CTkLabel(
            box_b, text=self.t("model_b_no_file"),
            font=("Segoe UI", 10, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.lbl_model_b_name.pack(fill="x")

        self.lbl_model_b_sub = ctk.CTkLabel(
            box_b, text="",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_model_b_sub.pack(fill="x")

        # --- Row 2: Box "So sánh 2 model" (Directly underneath Model A & B, Full Width) ---
        self.card_compare_box = ctk.CTkFrame(self.bottom_dock, fg_color=BG_CARD, corner_radius=12,
                                             border_width=1.5, border_color=BORDER_CLR, height=52)
        self.card_compare_box.pack(fill="x")
        self.card_verdict = self.card_compare_box  # Alias for verdict configure
        self.card_acts = self.card_compare_box     # Alias for acts configure

        c_inner = ctk.CTkFrame(self.card_compare_box, fg_color="transparent")
        c_inner.pack(fill="both", expand=True, padx=14, pady=6)

        # Left: Verdict status info
        v_box = ctk.CTkFrame(c_inner, fg_color="transparent")
        v_box.pack(side="left", fill="both", expand=True, padx=(0, 12))

        self.lbl_verdict_title = ctk.CTkLabel(
            v_box, text=self.t("verdict_model_placeholder"),
            font=("Segoe UI", 11, "bold"), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_verdict_title.pack(fill="x")

        self.lbl_verdict_sub = ctk.CTkLabel(
            v_box, text="Chọn 2 file PDF của Model A và Model B ở trên rồi nhấn '⚡ SO SÁNH 2 MODEL'.",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_verdict_sub.pack(fill="x")

        # Right: Comparison Action Buttons
        act_box = ctk.CTkFrame(c_inner, fg_color="transparent")
        act_box.pack(side="right")

        self.btn_run_model_comp = ctk.CTkButton(
            act_box, text=self.t("btn_run_model_comp"),
            font=("Segoe UI", 11, "bold"), height=34, width=175, corner_radius=8,
            fg_color=ACCENT_TEAL, hover_color=("#0F766E", "#00A88C"),
            command=self._run_comparison
        )
        self.btn_run_model_comp.pack(side="left", padx=(0, 8))

        self.btn_export_model_ecn = ctk.CTkButton(
            act_box, text=self.t("btn_export_model_ecn"),
            font=("Segoe UI", 10, "bold"), height=34, width=115, corner_radius=8,
            fg_color=ACCENT_BLUE, hover_color=("#1D4ED8", "#2563EB"),
            state="disabled",
            command=self._export_ecn
        )
        self.btn_export_model_ecn.pack(side="left", padx=(0, 8))

        self.btn_reset_model_comp = ctk.CTkButton(
            act_box, text=self.t("btn_reset_model_comp"),
            font=("Segoe UI", 10, "bold"), height=34, width=80, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_MUTED, hover_color=BG_HOVER,
            command=self._reset
        )
        self.btn_reset_model_comp.pack(side="left")

        # ── 3. MAIN WORKSPACE CARD (Top, Expandable) ──────────────────────────
        self.workspace_card = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                                           border_width=1, border_color=BORDER_CLR)
        self.workspace_card.pack(side="top", fill="both", expand=True, pady=(0, 4))

        # --- Top Workspace Toolbar ---
        ws_top = ctk.CTkFrame(self.workspace_card, fg_color="transparent")
        ws_top.pack(fill="x", padx=12, pady=(10, 6))

        # Filter Tabs
        self.filter_tabs = ctk.CTkSegmentedButton(
            ws_top,
            values=[self.t("filter_model_diff_all", n=0),
                    self.t("filter_model_added", n=0),
                    self.t("filter_model_removed", n=0),
                    self.t("filter_model_changed", n=0),
                    self.t("filter_model_match", n=0)],
            corner_radius=8, height=32,
            font=("Segoe UI", 9, "bold"),
            fg_color=BG_SURFACE,
            selected_color=ACCENT_TEAL,
            selected_hover_color=ACCENT_TEAL,
            unselected_color=BG_SURFACE,
            unselected_hover_color=BG_HOVER,
            text_color=TEXT_PRIMARY,
            command=self._on_filter_change
        )
        self.filter_tabs.set(self.t("filter_model_diff_all", n=0))
        self.filter_tabs.pack(side="left", padx=(0, 8))

        # Search Entry
        self.search_entry = ctk.CTkEntry(
            ws_top,
            placeholder_text="🔍  Tìm vị trí, mã LK...",
            font=FONT_BODY, height=32, corner_radius=8,
            fg_color=BG_SURFACE, border_width=1, border_color=BORDER_CLR,
            text_color=TEXT_PRIMARY
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.search_entry.bind("<KeyRelease>", self._on_search_key)

        # Display Mode Switcher (Integrated into table container)
        modes = [self.t("disp_mode_table"), self.t("disp_mode_drawing"), self.t("disp_mode_split")]
        self.seg_display_mode = ctk.CTkSegmentedButton(
            ws_top, values=modes,
            corner_radius=8, height=32, font=("Segoe UI", 9, "bold"),
            fg_color=BG_SURFACE, selected_color=ACCENT_TEAL, selected_hover_color=ACCENT_TEAL,
            unselected_color=BG_SURFACE, unselected_hover_color=BG_HOVER,
            text_color=TEXT_PRIMARY, command=self._on_display_mode_change
        )
        self.seg_display_mode.set(self.t("disp_mode_split"))
        self.seg_display_mode.pack(side="left", padx=(0, 8))

        # Full-Screen Button
        self.btn_fullscreen = ctk.CTkButton(
            ws_top, text=self.t("btn_fullscreen_drawing"),
            width=124, height=32, corner_radius=8,
            fg_color=ACCENT_BLUE, hover_color=("#1D4ED8", "#2563EB"),
            font=("Segoe UI", 10, "bold"), text_color="#FFFFFF",
            command=self._open_fullscreen_drawing
        )
        self.btn_fullscreen.pack(side="left", padx=(0, 10))

        self.lbl_table_count = ctk.CTkLabel(
            ws_top, text="0 dòng",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED
        )
        self.lbl_table_count.pack(side="right")

        # --- Content Container: Houses Table & Drawing Inspector ---
        self.content_frame = ctk.CTkFrame(self.workspace_card, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        # ── LEFT SUB-FRAME: Table Card ────────────────────────────────────────
        self.table_card = ctk.CTkFrame(self.content_frame, fg_color="transparent")

        # Treeview Container
        self.tree_container = tk.Frame(self.table_card, bg="#131929")
        self.tree_container.pack(fill="both", expand=True)

        col_ids = [c[0] for c in self.col_specs]
        self.tree = ttk.Treeview(
            self.tree_container,
            columns=col_ids,
            show="headings",
            style="Comp.Treeview",
            selectmode="browse"
        )
        self.vsb = tk.Scrollbar(self.tree_container, orient="vertical", command=self.tree.yview)
        self.hsb = tk.Scrollbar(self.tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.vsb.pack(side="right", fill="y")
        self.hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        for cid, width, anchor in self.col_specs:
            hdr_text = self.t("model_cols").get(cid, cid)
            self.tree.heading(cid, text=hdr_text, anchor=anchor)
            self.tree.column(cid, width=width, minwidth=35, anchor=anchor)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_select)

        # Empty overlay for table
        self.empty_overlay = ctk.CTkFrame(self.table_card, fg_color=BG_CARD, corner_radius=14)
        self.empty_overlay.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(self.empty_overlay, text="🔀", font=("Segoe UI", 40), fg_color="transparent").pack(pady=(14, 4))
        self.lbl_empty_title = ctk.CTkLabel(
            self.empty_overlay, text="Sẵn Sàng So Sánh 2 Model",
            font=("Segoe UI", 13, "bold"), text_color=TEXT_PRIMARY, fg_color="transparent"
        )
        self.lbl_empty_title.pack(padx=20, pady=(0, 4))
        self.lbl_empty_sub = ctk.CTkLabel(
            self.empty_overlay,
            text="Chọn 2 file Working Manual PDF ở phía dưới rồi nhấn '⚡ SO SÁNH 2 MODEL'.",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, fg_color="transparent"
        )
        self.lbl_empty_sub.pack(padx=20, pady=(0, 14))

        # ── RIGHT SUB-FRAME: Drawing Inspector Card ───────────────────────────
        self.inspector_card = ctk.CTkFrame(self.content_frame, fg_color=BG_SURFACE, corner_radius=10,
                                           border_width=1, border_color=BORDER_CLR)

        insp_inner = ctk.CTkFrame(self.inspector_card, fg_color="transparent")
        insp_inner.pack(fill="both", expand=True, padx=10, pady=8)

        self.lbl_insp_title = ctk.CTkLabel(
            insp_inner, text="🔍  " + self.t("title_drawing_inspector", ref="---"),
            font=("Segoe UI", 11, "bold"), text_color=ACCENT_TEAL, anchor="w"
        )
        self.lbl_insp_title.pack(fill="x", pady=(0, 4))

        # --- Toolbar Row 1: Stage Picker & View Mode ---
        tb_row1 = ctk.CTkFrame(insp_inner, fg_color="transparent")
        tb_row1.pack(fill="x", pady=(0, 4))

        self.lbl_stage = ctk.CTkLabel(
            tb_row1, text="Bản vẽ:", font=("Segoe UI", 10, "bold"), text_color=TEXT_PRIMARY
        )
        self.lbl_stage.pack(side="left", padx=(0, 4))

        self.opt_stage_picker = ctk.CTkOptionMenu(
            tb_row1,
            values=["🌐 Tất Cả Công Đoạn (Tự Động)"],
            font=("Segoe UI", 9, "bold"), height=28,
            fg_color=BG_CARD, button_color=ACCENT_TEAL, button_hover_color="#0F766E",
            text_color=TEXT_PRIMARY,
            command=self._on_stage_change
        )
        self.opt_stage_picker.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.seg_view_mode = ctk.CTkSegmentedButton(
            tb_row1,
            values=["↔️ Kéo Màn", "📄 Model A", "📄 Model B"],
            corner_radius=8, height=28,
            font=("Segoe UI", 9, "bold"),
            fg_color=BG_CARD,
            selected_color=ACCENT_TEAL,
            selected_hover_color=ACCENT_TEAL,
            unselected_color=BG_CARD,
            unselected_hover_color=BG_HOVER,
            text_color=TEXT_PRIMARY,
            command=self._on_view_mode_toggle
        )
        self.seg_view_mode.set("↔️ Kéo Màn")
        self.seg_view_mode.pack(side="right")

        # --- Toolbar Row 2: Zoom Controls & Curtain Swipe Slider (NO AUTO WIPE) ---
        tb_row2 = ctk.CTkFrame(insp_inner, fg_color="transparent")
        tb_row2.pack(fill="x", pady=(0, 6))

        self.btn_zoom_out = ctk.CTkButton(
            tb_row2, text="➖", width=28, height=26, corner_radius=6,
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            font=("Segoe UI", 9, "bold"), command=self._zoom_out
        )
        self.btn_zoom_out.pack(side="left", padx=(0, 2))

        self.lbl_zoom = ctk.CTkLabel(
            tb_row2, text="100%", width=46, font=("Segoe UI", 9, "bold"), text_color=TEXT_PRIMARY
        )
        self.lbl_zoom.pack(side="left", padx=(0, 2))

        self.btn_zoom_in = ctk.CTkButton(
            tb_row2, text="➕", width=28, height=26, corner_radius=6,
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            font=("Segoe UI", 9, "bold"), command=self._zoom_in
        )
        self.btn_zoom_in.pack(side="left", padx=(0, 4))

        self.btn_zoom_100 = ctk.CTkButton(
            tb_row2, text="1:1", width=32, height=26, corner_radius=6,
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_MUTED,
            font=("Segoe UI", 9), command=self._zoom_100
        )
        self.btn_zoom_100.pack(side="left", padx=(0, 4))

        self.btn_zoom_fit = ctk.CTkButton(
            tb_row2, text="📐 Vừa", width=44, height=26, corner_radius=6,
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_MUTED,
            font=("Segoe UI", 9), command=self._zoom_fit
        )
        self.btn_zoom_fit.pack(side="left", padx=(0, 4))

        rot_lbl = f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else "🔄 Xoay"
        self.btn_rotate = ctk.CTkButton(
            tb_row2, text=rot_lbl, width=64, height=26, corner_radius=6,
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            font=("Segoe UI", 9, "bold"), command=self._rotate_drawing
        )
        self.btn_rotate.pack(side="left", padx=(0, 8))

        # Curtain Swipe Controls Frame (Manual Only)
        self.curtain_ctrl_frame = ctk.CTkFrame(tb_row2, fg_color="transparent")
        self.curtain_ctrl_frame.pack(side="right", fill="x", expand=True)

        self.lbl_curtain_val = ctk.CTkLabel(
            self.curtain_ctrl_frame, text="Màn: 50%", font=("Segoe UI", 9),
            text_color=TEXT_MUTED, width=54
        )
        self.lbl_curtain_val.pack(side="left", padx=(0, 4))

        self.slider_curtain = ctk.CTkSlider(
            self.curtain_ctrl_frame, from_=0.0, to=1.0, number_of_steps=100, height=14,
            fg_color=BG_CARD, progress_color=ACCENT_TEAL,
            button_color=ACCENT_TEAL, button_hover_color="#00A88C",
            command=self._on_curtain_slider
        )
        self.slider_curtain.set(0.5)
        self.slider_curtain.pack(side="left", fill="x", expand=True)

        # --- Interactive Drawing Canvas Container ---
        self.canvas_container = ctk.CTkFrame(insp_inner, fg_color=BG_CARD, corner_radius=8,
                                             border_width=1, border_color=BORDER_CLR)
        self.canvas_container.pack(fill="both", expand=True, pady=(0, 6))

        self.canvas_vsb = tk.Scrollbar(self.canvas_container, orient="vertical")
        self.canvas_hsb = tk.Scrollbar(self.canvas_container, orient="horizontal")

        self.canvas = tk.Canvas(
            self.canvas_container, bg="#0B0F1A", highlightthickness=0,
            xscrollcommand=self.canvas_hsb.set, yscrollcommand=self.canvas_vsb.set
        )
        self.canvas_vsb.config(command=self.canvas.yview)
        self.canvas_hsb.config(command=self.canvas.xview)

        self.canvas_vsb.pack(side="right", fill="y")
        self.canvas_hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.bind("<ButtonPress-1>", self._on_canvas_b1_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_b1_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_b1_release)
        self.canvas.bind("<ButtonPress-2>", self._on_canvas_b2_press)
        self.canvas.bind("<B2-Motion>", self._on_canvas_b2_motion)
        self.canvas.bind("<MouseWheel>", self._on_canvas_mousewheel)
        self.canvas.bind("<Motion>", self._on_canvas_mouse_motion)
        self.canvas.bind("<r>", lambda e: self._rotate_drawing())
        self.canvas.bind("<R>", lambda e: self._rotate_drawing())

        # Empty Canvas Overlay
        self.empty_canvas_overlay = ctk.CTkFrame(self.canvas, fg_color="transparent")
        self.empty_canvas_overlay.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(self.empty_canvas_overlay, text="📐", font=("Segoe UI", 36), fg_color="transparent").pack(pady=(0, 4))
        self.lbl_empty_canvas_title = ctk.CTkLabel(
            self.empty_canvas_overlay, text="Chưa Nạp Bản Vẽ",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY, fg_color="transparent"
        )
        self.lbl_empty_canvas_title.pack()
        self.lbl_empty_canvas_sub = ctk.CTkLabel(
            self.empty_canvas_overlay, text="Nhấn '⚡ SO SÁNH 2 MODEL' để xem bản vẽ & hiệu ứng kéo màn.",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, fg_color="transparent"
        )
        self.lbl_empty_canvas_sub.pack(pady=(2, 0))

        # --- Details Panel (Compact 2-Row Layout to maximize canvas height) ---
        self.det_box = ctk.CTkFrame(insp_inner, fg_color=BG_CARD, corner_radius=8,
                                    border_width=1, border_color=BORDER_CLR)
        self.det_box.pack(fill="x")

        det_inner = ctk.CTkFrame(self.det_box, fg_color="transparent")
        det_inner.pack(fill="both", expand=True, padx=10, pady=4)

        d_row1 = ctk.CTkFrame(det_inner, fg_color="transparent")
        d_row1.pack(fill="x")

        self.lbl_det_ref = ctk.CTkLabel(
            d_row1, text="📍 Vị trí: ---  |  Công đoạn: ---",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.lbl_det_ref.pack(side="left", fill="x", expand=True)

        self.lbl_det_status = ctk.CTkLabel(
            d_row1, text="Trạng thái: ---",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED, anchor="e"
        )
        self.lbl_det_status.pack(side="right")

        d_row2 = ctk.CTkFrame(det_inner, fg_color="transparent")
        d_row2.pack(fill="x", pady=(2, 0))

        self.lbl_det_part_a = ctk.CTkLabel(
            d_row2, text="Model A: ---",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_det_part_a.pack(side="left", padx=(0, 10))

        self.lbl_det_part_b = ctk.CTkLabel(
            d_row2, text="Model B: ---",
            font=("Segoe UI", 9), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_det_part_b.pack(side="left", padx=(0, 10))

        self.lbl_det_note = ctk.CTkLabel(
            d_row2, text="Đánh giá: ---",
            font=("Segoe UI", 9, "bold"), text_color=ACCENT_AMBER, anchor="w"
        )
        self.lbl_det_note.pack(side="left", fill="x", expand=True)

        # Default layout: Split View (Side-by-side) with equal 50:50 distribution
        self.content_frame.columnconfigure(0, weight=1, uniform="split_col")
        self.content_frame.columnconfigure(1, weight=1, uniform="split_col")
        self.content_frame.rowconfigure(0, weight=1)

        self.table_card.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.inspector_card.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

    # ── Display Mode Switcher & Fullscreen ────────────────────────────────────
    def _on_display_mode_change(self, choice: str):
        c_low = choice.lower()
        self.table_card.grid_forget()
        self.inspector_card.grid_forget()

        if "bảng" in c_low or "table" in c_low or "表格" in c_low:
            self.display_mode = "table"
            if not self.row_models.winfo_ismapped():
                self.row_models.pack(fill="x", pady=(0, 6), before=self.card_compare_box)
            self.content_frame.columnconfigure(0, weight=1, uniform="")
            self.content_frame.columnconfigure(1, weight=0, uniform="")
            self.table_card.grid(row=0, column=0, columnspan=2, sticky="nsew")
        elif "bản vẽ" in c_low or "drawing" in c_low or "图纸" in c_low:
            self.display_mode = "drawing"
            # Hide model selection cards row to give maximum height to PCB drawing inspector!
            if self.row_models.winfo_ismapped():
                self.row_models.pack_forget()
            self.content_frame.columnconfigure(0, weight=0, uniform="")
            self.content_frame.columnconfigure(1, weight=1, uniform="")
            self.inspector_card.grid(row=0, column=0, columnspan=2, sticky="nsew")
            self.canvas_img_id = None
            self.after(60, self._zoom_fit)
        else: # split
            self.display_mode = "split"
            if not self.row_models.winfo_ismapped():
                self.row_models.pack(fill="x", pady=(0, 6), before=self.card_compare_box)
            self.content_frame.columnconfigure(0, weight=1, uniform="split_col")
            self.content_frame.columnconfigure(1, weight=1, uniform="split_col")
            self.table_card.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
            self.inspector_card.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
            self.canvas_img_id = None
            self._render_drawing(center_ref=False)

    def _open_fullscreen_drawing(self):
        if not self.pdf_a_path and not self.pdf_b_path:
            messagebox.showwarning("Chưa Có Bản Vẽ", "Vui lòng chọn file PDF và nhấn '⚡ SO SÁNH 2 MODEL' trước khi mở toàn màn hình!")
            return

        if self.fullscreen_window is not None:
            try:
                self.fullscreen_window.lift()
                self.fullscreen_window.focus_force()
                return
            except Exception:
                self.fullscreen_window = None

        self.fullscreen_window = DrawingFullScreenWindow(
            parent=self,
            app=self.app,
            pdf_a_path=self.pdf_a_path,
            pdf_b_path=self.pdf_b_path,
            comp_result=self.comp_result or {},
            drawing_catalog=self.drawing_catalog,
            active_item=self.active_item,
            current_stage_id=self.current_stage_id,
            active_page_a=self.active_page_a,
            active_page_b=self.active_page_b,
            split_ratio=self.split_ratio,
            view_mode=self.view_mode,
            zoom_level=self.zoom_level,
            rotation_angle=self.rotation_angle,
            on_close_cb=self._on_fullscreen_closed,
            on_select_item_cb=self._on_item_selected_from_fullscreen
        )

    def _on_fullscreen_closed(self):
        self.fullscreen_window = None
        if hasattr(self, "btn_rotate"):
            self.btn_rotate.configure(text=f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else "🔄 Xoay")
        self._zoom_fit()

    def _on_item_selected_from_fullscreen(self, item: dict):
        self.active_item = item
        ref_des = item.get("ref_des")
        stage = item.get("stage")
        for child in self.tree.get_children():
            vals = self.tree.item(child, "values")
            if vals and len(vals) > 2 and vals[1] == ref_des and vals[2] == stage:
                self.tree.selection_set(child)
                self.tree.see(child)
                break
        self._update_inspector(skip_fullscreen=True)

    # ── File Selectors & Drop Handling ────────────────────────────────────────
    def _select_pdf_a(self):
        fn = filedialog.askopenfilename(
            title=self.t("btn_choose_model_a"),
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if fn:
            sz = os.path.getsize(fn) if os.path.exists(fn) else 0
            if sz > MAX_FILE_SIZE_BYTES:
                messagebox.showwarning(
                    self.t("title_file_too_large"),
                    self.t("err_file_too_large", name=os.path.basename(fn),
                           size=format_file_size(sz), max=f"{MAX_FILE_SIZE_MB}MB")
                )
                self.app.show_toast(f"⚠️ File vượt quá {MAX_FILE_SIZE_MB}MB!")
                return

            candidate_a = os.path.normpath(fn)
            if self.pdf_b_path:
                is_valid, err_msg, info_a, info_b = validate_model_pair(candidate_a, self.pdf_b_path)
                if not is_valid:
                    messagebox.showerror("Bản Vẽ Không Thỏa Mãn", err_msg)
                    self.app.show_toast("⚠️ Bản vẽ không cùng Model hoặc trùng Series!")
                    return
                self.pdf_a_path = candidate_a
                self.lbl_model_a_name.configure(text=f"📄 {os.path.basename(candidate_a)} ({format_file_size(sz)})")
                self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or 'Gốc'}")
                self.lbl_model_b_sub.configure(text=f"Model: {info_b['display_model']}  |  Series: {info_b['display_series'] or 'Mới'}")
                self.app.show_toast(f"✅ Hợp lệ: Model {info_a['display_model']} (Series {info_a['display_series'] or 'A'} vs {info_b['display_series'] or 'B'})")
            else:
                self.pdf_a_path = candidate_a
                info_a = extract_model_info(candidate_a)
                self.lbl_model_a_name.configure(text=f"📄 {os.path.basename(candidate_a)} ({format_file_size(sz)})")
                self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or '---'}")
                self.app.show_toast(f"Đã nạp Model A ({info_a['display_model']}). Vui lòng chọn Model B cùng Model nhưng khác Series.")

    def _select_pdf_b(self):
        fn = filedialog.askopenfilename(
            title=self.t("btn_choose_model_b"),
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if fn:
            sz = os.path.getsize(fn) if os.path.exists(fn) else 0
            if sz > MAX_FILE_SIZE_BYTES:
                messagebox.showwarning(
                    self.t("title_file_too_large"),
                    self.t("err_file_too_large", name=os.path.basename(fn),
                           size=format_file_size(sz), max=f"{MAX_FILE_SIZE_MB}MB")
                )
                self.app.show_toast(f"⚠️ File vượt quá {MAX_FILE_SIZE_MB}MB!")
                return

            candidate_b = os.path.normpath(fn)
            if self.pdf_a_path:
                is_valid, err_msg, info_a, info_b = validate_model_pair(self.pdf_a_path, candidate_b)
                if not is_valid:
                    messagebox.showerror("Bản Vẽ Không Thỏa Mãn", err_msg)
                    self.app.show_toast("⚠️ Bản vẽ không cùng Model hoặc trùng Series!")
                    return
                self.pdf_b_path = candidate_b
                self.lbl_model_b_name.configure(text=f"📄 {os.path.basename(candidate_b)} ({format_file_size(sz)})")
                self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or 'Gốc'}")
                self.lbl_model_b_sub.configure(text=f"Model: {info_b['display_model']}  |  Series: {info_b['display_series'] or 'Mới'}")
                self.app.show_toast(f"✅ Hợp lệ: Model {info_b['display_model']} (Series {info_a['display_series'] or 'A'} vs {info_b['display_series'] or 'B'})")
            else:
                self.pdf_b_path = candidate_b
                info_b = extract_model_info(candidate_b)
                self.lbl_model_b_name.configure(text=f"📄 {os.path.basename(candidate_b)} ({format_file_size(sz)})")
                self.lbl_model_b_sub.configure(text=f"Model: {info_b['display_model']}  |  Series: {info_b['display_series'] or '---'}")
                self.app.show_toast(f"Đã nạp Model B ({info_b['display_model']}). Vui lòng chọn Model A cùng Model nhưng khác Series.")

    def handle_drop_files(self, files):
        pdf_files = []
        for item in files:
            if isinstance(item, bytes):
                item = item.decode("utf-8", errors="ignore")
            item = os.path.normpath(item)
            if os.path.isfile(item) and item.lower().endswith(".pdf"):
                sz = os.path.getsize(item) if os.path.exists(item) else 0
                if sz > MAX_FILE_SIZE_BYTES:
                    messagebox.showwarning(
                        self.t("title_file_too_large"),
                        self.t("err_file_too_large", name=os.path.basename(item),
                               size=format_file_size(sz), max=f"{MAX_FILE_SIZE_MB}MB")
                    )
                    continue
                pdf_files.append(item)

        if len(pdf_files) >= 2:
            cand_a = pdf_files[0]
            cand_b = pdf_files[1]
            is_valid, err_msg, info_a, info_b = validate_model_pair(cand_a, cand_b)
            if not is_valid:
                messagebox.showerror("Bản Vẽ Không Thỏa Mãn", err_msg)
                self.app.show_toast("⚠️ 2 file kéo thả không cùng Model hoặc trùng Series!")
                return
            self.pdf_a_path = cand_a
            self.pdf_b_path = cand_b
            sz_a = os.path.getsize(self.pdf_a_path)
            sz_b = os.path.getsize(self.pdf_b_path)
            self.lbl_model_a_name.configure(text=f"📄 {os.path.basename(self.pdf_a_path)} ({format_file_size(sz_a)})")
            self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or 'Gốc'}")
            self.lbl_model_b_name.configure(text=f"📄 {os.path.basename(self.pdf_b_path)} ({format_file_size(sz_b)})")
            self.lbl_model_b_sub.configure(text=f"Model: {info_b['display_model']}  |  Series: {info_b['display_series'] or 'Mới'}")
            self.app.show_toast(f"✅ Đã nạp cặp bản vẽ: Model {info_a['display_model']} (Series {info_a['display_series'] or 'A'} vs {info_b['display_series'] or 'B'})")
        elif len(pdf_files) == 1:
            cand = pdf_files[0]
            sz = os.path.getsize(cand)
            if not self.pdf_a_path:
                if self.pdf_b_path:
                    is_valid, err_msg, info_a, info_b = validate_model_pair(cand, self.pdf_b_path)
                    if not is_valid:
                        messagebox.showerror("Bản Vẽ Không Thỏa Mãn", err_msg)
                        self.app.show_toast("⚠️ File kéo thả không khớp Model với Model B!")
                        return
                    self.pdf_a_path = cand
                    self.lbl_model_a_name.configure(text=f"📄 {os.path.basename(cand)} ({format_file_size(sz)})")
                    self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or 'Gốc'}")
                    self.lbl_model_b_sub.configure(text=f"Model: {info_b['display_model']}  |  Series: {info_b['display_series'] or 'Mới'}")
                    self.app.show_toast("✅ Đã nạp đủ 2 file Model A & Model B!")
                else:
                    self.pdf_a_path = cand
                    info_a = extract_model_info(cand)
                    self.lbl_model_a_name.configure(text=f"📄 {os.path.basename(cand)} ({format_file_size(sz)})")
                    self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or '---'}")
                    self.app.show_toast("Đã nạp Model A! Kéo thả thêm file cho Model B.")
            else:
                is_valid, err_msg, info_a, info_b = validate_model_pair(self.pdf_a_path, cand)
                if not is_valid:
                    messagebox.showerror("Bản Vẽ Không Thỏa Mãn", err_msg)
                    self.app.show_toast("⚠️ File kéo thả không khớp Model với Model A!")
                    return
                self.pdf_b_path = cand
                self.lbl_model_b_name.configure(text=f"📄 {os.path.basename(cand)} ({format_file_size(sz)})")
                self.lbl_model_a_sub.configure(text=f"Model: {info_a['display_model']}  |  Series: {info_a['display_series'] or 'Gốc'}")
                self.lbl_model_b_sub.configure(text=f"Model: {info_b['display_model']}  |  Series: {info_b['display_series'] or 'Mới'}")
                self.app.show_toast("✅ Đã nạp đủ 2 file Model A & Model B!")

    # ── Comparison Execution with AI Laser Scan Animation ─────────────────────
    def _run_comparison(self):
        if not self.pdf_a_path or not self.pdf_b_path:
            messagebox.showwarning("Thiếu File", "Vui lòng chọn cả 2 file PDF của Model A và Model B!")
            return

        is_valid, err_msg, info_a, info_b = validate_model_pair(self.pdf_a_path, self.pdf_b_path)
        if not is_valid:
            messagebox.showerror("Không Thể So Sánh", err_msg)
            return

        if getattr(self, "_is_scanning", False):
            return

        self._is_scanning = True

        # Ensure drawing is visible during scan
        if self.display_mode == "table":
            self._on_display_mode_change("◫ Song Song")
            self.seg_display_mode.set("◫ Song Song")

        self.btn_run_model_comp.configure(state="disabled", text="⚡ Đang quét Laser AI... (0%)")
        self.update_idletasks()

        # Hide empty canvas overlay & preload base drawing preview
        self.empty_canvas_overlay.place_forget()
        bw, bh = 600, 400
        try:
            init_pg_a = getattr(self, "active_page_a", 1)
            init_pg_b = getattr(self, "active_page_b", 1)
            base_a, base_b, _, bw, bh = get_annotated_base_images(
                self.pdf_a_path, self.pdf_b_path, init_pg_a, init_pg_b,
                zoom=self.zoom_level, rotation=self.rotation_angle
            )
            cw = max(200, self.canvas.winfo_width())
            ch = max(200, self.canvas.winfo_height())
            self.pan_offset_x = 0
            self.pan_offset_y = 0
            self.draw_offset_x = draw_x0
            self.draw_offset_y = draw_y0
            self.current_meta = {"width": bw, "height": bh}

            self.tk_canvas_img = ImageTk.PhotoImage(base_b)
            self.canvas.delete("all")
            self.canvas_img_id = self.canvas.create_image(draw_x0, draw_y0, image=self.tk_canvas_img, anchor="nw", tags="drawing")
            self._update_scrollbars()
        except Exception:
            cw = max(200, self.canvas.winfo_width())
            ch = max(200, self.canvas.winfo_height())
            draw_x0 = 0
            draw_y0 = 0
            bw = cw
            bh = ch
            self.draw_offset_x = 0
            self.draw_offset_y = 0

        # Remove any lingering HUD
        if getattr(self, "_scan_hud", None) is not None:
            try:
                self._scan_hud.destroy()
            except Exception:
                pass
            self._scan_hud = None

        # Build Cyberpunk Holographic HUD Card over the canvas, precisely centered on the drawing
        is_dark = (self.app.current_theme == "dark")
        self._scan_hud = ctk.CTkFrame(
            self.canvas_container, fg_color="#0B132B" if is_dark else "#0F172A",
            corner_radius=12, border_width=1.5, border_color="#00F0FF"
        )
        hud_center_x = draw_x0 + bw // 2
        hud_top_y = max(10, draw_y0 + 15)
        self._scan_hud.place(x=hud_center_x, y=hud_top_y, anchor="n")

        hud_inner = ctk.CTkFrame(self._scan_hud, fg_color="transparent")
        hud_inner.pack(padx=20, pady=12)

        # Top row: Title & Telemetry
        h_row1 = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row1.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(h_row1, text="📡", font=("Segoe UI", 16)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            h_row1, text="VIPQC AI NEURAL SCANNER — QUÉT VECTOR BẢN VẼ PCB",
            font=("Segoe UI", 11, "bold"), text_color="#00F0FF"
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(
            h_row1, text="60 FPS • 4K VECTOR ENGINE",
            font=("Consolas", 9, "bold"), text_color="#38BDF8"
        ).pack(side="right")

        # Middle row: Animated Progress Bar + Percentage
        h_row2 = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row2.pack(fill="x", pady=(0, 6))

        scan_prog = ctk.CTkProgressBar(
            h_row2, width=380, height=10, corner_radius=5,
            fg_color="#1E293B", progress_color="#00F0FF"
        )
        scan_prog.set(0.0)
        scan_prog.pack(side="left", padx=(0, 10))

        lbl_scan_pct = ctk.CTkLabel(
            h_row2, text="0%", font=("Consolas", 11, "bold"), text_color="#00F0FF", width=44
        )
        lbl_scan_pct.pack(side="left")

        # Bottom row: Live Terminal Log Phase
        lbl_scan_log = ctk.CTkLabel(
            hud_inner, text="[01/05] 📡 Khởi động VIPQC Neural Vision Engine...",
            font=("Consolas", 10), text_color="#E2E8F0", anchor="w"
        )
        lbl_scan_log.pack(fill="x")

        # Background comparison worker
        comp_data = {"res": None, "err": None, "done": False}
        def _bg_worker():
            try:
                clear_curtain_cache()
                res = compare_model_manuals(self.pdf_a_path, self.pdf_b_path)
                comp_data["res"] = res
            except Exception as ex:
                comp_data["err"] = ex
            finally:
                comp_data["done"] = True

        threading.Thread(target=_bg_worker, daemon=True).start()

        # Scan parameters (3.5 seconds)
        start_time = time.time()
        scan_duration = 3.5

        reticles = [
            (0.28, 0.22, "R652", "#F59E0B"),
            (0.48, 0.38, "IC101", "#10B981"),
            (0.68, 0.52, "C882", "#38BDF8"),
            (0.35, 0.68, "D301", "#10B981"),
            (0.72, 0.32, "Q502", "#F59E0B"),
            (0.56, 0.78, "L601", "#38BDF8"),
        ]

        def _update_scan():
            if not getattr(self, "_is_scanning", False):
                return

            now = time.time()
            elapsed = now - start_time
            t = min(1.0, elapsed / scan_duration)

            cw = max(200, self.canvas.winfo_width())
            ch = max(200, self.canvas.winfo_height())

            draw_x0 = getattr(self, "draw_offset_x", max(0, (cw - bw) // 2))
            draw_y0 = getattr(self, "draw_offset_y", max(0, (ch - bh) // 2))
            draw_x1 = draw_x0 + bw
            draw_y1 = draw_y0 + bh
            draw_w = max(1, bw)
            draw_h = max(1, bh)

            # 2 passes: down, then up across the drawing height
            cycle = t * 2.0
            if cycle <= 1.0:
                beam_y = draw_y0 + cycle * draw_h
            else:
                beam_y = draw_y0 + (2.0 - cycle) * draw_h

            # Progress UI
            scan_prog.set(t)
            pct = int(t * 100)
            lbl_scan_pct.configure(text=f"{pct}%")
            self.btn_run_model_comp.configure(text=f"⚡ Đang quét Laser AI... ({pct}%)")

            # Dynamic terminal phases
            if t < 0.20:
                lbl_scan_log.configure(text="[01/05] 📡 Khởi động VIPQC Neural Vision Engine...")
            elif t < 0.40:
                lbl_scan_log.configure(text="[02/05] 📐 Căn chỉnh toạ độ bản vẽ & PWB Fiducials...")
            elif t < 0.65:
                lbl_scan_log.configure(text="[03/05] 🔬 Quét laser vector linh kiện (Chip, IC, Resistor)...")
            elif t < 0.85:
                lbl_scan_log.configure(text="[04/05] ⚡ Đối chiếu ma trận sai lệch Model A vs Model B...")
            else:
                lbl_scan_log.configure(text="[05/05] 🎯 Tổng hợp bản đồ sai lệch & Hoàn tất phân tích!")

            # Redraw canvas laser elements strictly confined within [draw_x0, draw_x1]
            self.canvas.delete("ai_scan_anim")

            # 1. Trailing aura gradient - strictly bounded within drawing width and height
            trail_h = min(36, draw_h * 0.12)
            if cycle <= 1.0:
                self.canvas.create_polygon(
                    draw_x0, max(draw_y0, beam_y - trail_h),
                    draw_x1, max(draw_y0, beam_y - trail_h),
                    draw_x1, beam_y,
                    draw_x0, beam_y,
                    fill="#0891B2", stipple="gray25", tags="ai_scan_anim"
                )
            else:
                self.canvas.create_polygon(
                    draw_x0, beam_y,
                    draw_x1, beam_y,
                    draw_x1, min(draw_y1, beam_y + trail_h),
                    draw_x0, min(draw_y1, beam_y + trail_h),
                    fill="#0891B2", stipple="gray25", tags="ai_scan_anim"
                )

            # 2. Outer glow line strictly between [draw_x0, draw_x1]
            self.canvas.create_line(draw_x0, beam_y, draw_x1, beam_y, fill="#0284C7", width=6, stipple="gray50", tags="ai_scan_anim")
            # 3. Cyan glow line
            self.canvas.create_line(draw_x0, beam_y, draw_x1, beam_y, fill="#22D3EE", width=3, tags="ai_scan_anim")
            # 4. Core laser beam
            self.canvas.create_line(draw_x0, beam_y, draw_x1, beam_y, fill="#00F0FF", width=1.5, tags="ai_scan_anim")
            # 5. Pure white center strand
            inset = min(30, draw_w * 0.08)
            self.canvas.create_line(draw_x0 + inset, beam_y, draw_x1 - inset, beam_y, fill="#FFFFFF", width=1, tags="ai_scan_anim")

            # Side laser beacons - positioned inside drawing boundaries
            beacon_inset = min(48, draw_w * 0.12)
            self.canvas.create_text(draw_x0 + beacon_inset, beam_y - 12, text="▶ LASER SCAN", fill="#00F0FF", font=("Consolas", 8, "bold"), tags="ai_scan_anim")
            self.canvas.create_text(draw_x1 - beacon_inset, beam_y - 12, text="LASER SCAN ◀", fill="#00F0FF", font=("Consolas", 8, "bold"), tags="ai_scan_anim")

            # Center targeting reticle - sweeps strictly within central 60% of the drawing
            sweep_x = int(draw_x0 + (draw_w * 0.22) + ((draw_w * 0.56) * (((beam_y - draw_y0) / float(max(1, draw_h))) % 1.0)))
            self.canvas.create_text(sweep_x, beam_y - 14, text="[ ⌖ AI MATRIX ]", fill="#38BDF8", font=("Consolas", 9, "bold"), tags="ai_scan_anim")

            # Dynamic component lock-on reticles - strictly inside drawing bounds
            for rx_pct, ry_pct, rname, rclr in reticles:
                rx = int(draw_x0 + rx_pct * draw_w)
                ry = int(draw_y0 + ry_pct * draw_h)
                if abs(beam_y - ry) < 65:
                    self.canvas.create_rectangle(rx - 15, ry - 15, rx + 15, ry + 15, outline=rclr, width=1.5, tags="ai_scan_anim")
                    self.canvas.create_text(rx, ry - 22, text=f"⌖ {rname}", fill=rclr, font=("Consolas", 8, "bold"), tags="ai_scan_anim")

            if t < 1.0 or not comp_data["done"]:
                self.after(16, _update_scan)
            else:
                self._finish_scan(comp_data, draw_x0, draw_y0, draw_x1, draw_y1)

        self.after(20, _update_scan)

    def _finish_scan(self, comp_data: dict, draw_x0: int = 0, draw_y0: int = 0, draw_x1: int = 0, draw_y1: int = 0):
        self._is_scanning = False
        cw = max(200, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        if draw_x1 <= 0:
            draw_x1 = cw
        if draw_y1 <= 0:
            draw_y1 = ch

        # Flash effect strictly over the drawing!
        self.canvas.delete("ai_scan_anim")
        self.canvas.create_rectangle(draw_x0, draw_y0, draw_x1, draw_y1, fill="#00F0FF", stipple="gray25", tags="ai_scan_flash")

        def _cleanup_and_render():
            self.canvas.delete("ai_scan_flash")
            if getattr(self, "_scan_hud", None) is not None:
                try:
                    self._scan_hud.destroy()
                except Exception:
                    pass
                self._scan_hud = None

            self.btn_run_model_comp.configure(state="normal", text=self.t("btn_run_model_comp"))

            if comp_data.get("err"):
                messagebox.showerror("Lỗi So Sánh", f"Đã xảy ra lỗi khi so sánh: {comp_data['err']}")
                return

            res = comp_data.get("res") or {}
            self.comp_result = res
            s = res.get("summary", {})

            # Catalog of stages & pages
            self.drawing_catalog = res.get("drawing_catalog", [])
            if not self.drawing_catalog:
                self.drawing_catalog = get_drawing_pages_catalog(self.pdf_a_path, self.pdf_b_path)

            cat_labels = [c["label"] for c in self.drawing_catalog]
            if cat_labels:
                self.opt_stage_picker.configure(values=cat_labels)
                self.opt_stage_picker.set(cat_labels[0])
                self.current_stage_id = self.drawing_catalog[0]["id"]
                self.active_page_a = self.drawing_catalog[0]["page_idx_a"]
                self.active_page_b = self.drawing_catalog[0]["page_idx_b"]

            # Update Model A / B metadata subtitles
            mod_a = res.get("model_a", {})
            mod_b = res.get("model_b", {})
            self.lbl_model_a_sub.configure(text=f"PWB: {mod_a.get('pwb_code')}  |  Quy trình: {', '.join(mod_a.get('processes', []))}")
            self.lbl_model_b_sub.configure(text=f"PWB: {mod_b.get('pwb_code')}  |  Quy trình: {', '.join(mod_b.get('processes', []))}")

            # Update Verdict Banner
            if s.get("is_identical"):
                self.card_verdict.configure(border_color="#10B981")
                self.lbl_verdict_title.configure(text=self.t("verdict_model_match_title"), text_color="#10B981")
                self.lbl_verdict_sub.configure(text=self.t("verdict_model_match_sub", total=s.get("total_positions", 0)))
            else:
                self.card_verdict.configure(border_color="#F59E0B")
                self.lbl_verdict_title.configure(
                    text=self.t("verdict_model_diff_title", diff=s.get("diff_count", 0)),
                    text_color="#F59E0B"
                )
                self.lbl_verdict_sub.configure(
                    text=self.t("verdict_model_diff_sub",
                                added=s.get("added_count", 0),
                                removed=s.get("removed_count", 0),
                                changed=s.get("changed_count", 0))
                )

            # Update filter tabs
            self.filter_tabs.configure(
                values=[
                    self.t("filter_model_diff_all", n=s.get("diff_count", 0)),
                    self.t("filter_model_added", n=s.get("added_count", 0)),
                    self.t("filter_model_removed", n=s.get("removed_count", 0)),
                    self.t("filter_model_changed", n=s.get("changed_count", 0)),
                    self.t("filter_model_match", n=s.get("match_count", 0))
                ]
            )
            default_tab = self.t("filter_model_diff_all", n=s.get("diff_count", 0)) if s.get("diff_count", 0) > 0 else self.t("filter_model_match", n=s.get("match_count", 0))
            self.filter_tabs.set(default_tab)
            self.current_filter = "diff_all" if s.get("diff_count", 0) > 0 else "match"

            self.canvas_img_id = None
            self._populate_tree()
            self._render_drawing(center_ref=False)

            self.btn_export_model_ecn.configure(state="normal")
            self.app._refresh_badges()
            self.app.show_toast(f"⚡ Đã quét xong bản vẽ! Tìm thấy {s.get('diff_count', 0)} điểm sai lệch.")

        self.after(120, _cleanup_and_render)

    def _on_stage_change(self, choice: str):
        matched = next((c for c in self.drawing_catalog if c["label"] == choice), None)
        if matched:
            self.canvas_img_id = None
            self.current_stage_id = matched["id"]
            self.active_page_a = matched["page_idx_a"]
            self.active_page_b = matched["page_idx_b"]
            self._populate_tree()
            self._render_drawing(center_ref=False)

    def _on_view_mode_toggle(self, choice: str):
        self.canvas_img_id = None
        c_low = choice.lower()
        if "kéo" in c_low or "swipe" in c_low:
            self.view_mode = "curtain"
            self.curtain_ctrl_frame.pack(side="right", fill="x", expand=True)
        elif "model a" in c_low or "gốc" in c_low:
            self.view_mode = "a"
            self.curtain_ctrl_frame.pack_forget()
        else:
            self.view_mode = "b"
            self.curtain_ctrl_frame.pack_forget()
        self._render_drawing(center_ref=False)

    def _zoom_in(self):
        self.canvas_img_id = None
        self.zoom_level = min(5.0, round(self.zoom_level * 1.25, 2))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing(center_ref=False)

    def _zoom_out(self):
        self.canvas_img_id = None
        self.zoom_level = max(0.35, round(self.zoom_level / 1.25, 2))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing(center_ref=False)

    def _zoom_100(self):
        self.canvas_img_id = None
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.zoom_level = 1.0
        self.lbl_zoom.configure(text="100%")
        self._render_drawing(center_ref=False)

    def _rotate_drawing(self):
        self.canvas_img_id = None
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.rotation_angle = (self.rotation_angle + 90) % 360
        self.btn_rotate.configure(text=f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else "🔄 Xoay")
        self._zoom_fit()

    def _zoom_fit(self):
        self.canvas_img_id = None
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.update_idletasks()
        cw = max(200, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        base_w, base_h = (792.0, 612.0) if (self.rotation_angle % 180 == 90) else (612.0, 792.0)
        fit_zoom = min(cw / base_w, ch / base_h)
        self.zoom_level = max(0.35, round(fit_zoom, 2))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing(center_ref=False)

    def _request_fast_curtain_render(self):
        if not getattr(self, "_curtain_render_pending", False):
            self._curtain_render_pending = True
            self.after_idle(self._execute_fast_curtain_render)

    def _execute_fast_curtain_render(self):
        self._curtain_render_pending = False
        self._render_drawing(center_ref=False)

    def _on_curtain_slider(self, val):
        self.split_ratio = float(val)
        self.lbl_curtain_val.configure(text=f"Màn: {int(self.split_ratio * 100)}%")
        self._request_fast_curtain_render()

    def _update_scrollbars(self):
        cw = max(200, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        draw_x0 = getattr(self, "draw_offset_x", 0)
        draw_y0 = getattr(self, "draw_offset_y", 0)
        mw = self.current_meta.get("width", cw)
        mh = self.current_meta.get("height", ch)
        min_x = min(0, draw_x0 + self.pan_offset_x)
        min_y = min(0, draw_y0 + self.pan_offset_y)
        max_x = max(cw, draw_x0 + self.pan_offset_x + mw)
        max_y = max(ch, draw_y0 + self.pan_offset_y + mh)
        self.canvas.config(scrollregion=(min_x, min_y, max_x, max_y))

    # ── Interactive Canvas Handlers ──────────────────────────────────────────
    def _on_canvas_mouse_motion(self, event):
        split_x = self.current_meta.get("split_x")
        draw_x0 = getattr(self, "draw_offset_x", 0)
        curtain_x = draw_x0 + self.pan_offset_x + (split_x if split_x is not None else -9999)
        if self.view_mode == "curtain" and split_x is not None and abs(event.x - curtain_x) < 30:
            self.canvas.config(cursor="sb_h_double_arrow")
        else:
            self.canvas.config(cursor="")

    def _on_canvas_b1_press(self, event):
        split_x = self.current_meta.get("split_x")
        draw_x0 = getattr(self, "draw_offset_x", 0)
        curtain_x = draw_x0 + self.pan_offset_x + (split_x if split_x is not None else -9999)
        if self.view_mode == "curtain" and split_x is not None and abs(event.x - curtain_x) < 30:
            self.dragging_curtain = True
            self.canvas.config(cursor="sb_h_double_arrow")
        else:
            self.dragging_curtain = False
            self.canvas.config(cursor="fleur")
            self._pan_start_x = event.x
            self._pan_start_y = event.y
            self._pan_init_offset_x = self.pan_offset_x
            self._pan_init_offset_y = self.pan_offset_y

    def _on_canvas_b1_motion(self, event):
        if self.dragging_curtain:
            draw_x0 = getattr(self, "draw_offset_x", 0)
            w = max(1, self.current_meta.get("width", 1))
            ratio = max(0.0, min(1.0, (event.x - (draw_x0 + self.pan_offset_x)) / float(w)))
            self.split_ratio = ratio
            self.slider_curtain.set(ratio)
            self.lbl_curtain_val.configure(text=f"Màn: {int(ratio * 100)}%")
            self._request_fast_curtain_render()
        else:
            dx = event.x - self._pan_start_x
            dy = event.y - self._pan_start_y
            self.pan_offset_x = self._pan_init_offset_x + dx
            self.pan_offset_y = self._pan_init_offset_y + dy
            draw_x0 = getattr(self, "draw_offset_x", 0)
            draw_y0 = getattr(self, "draw_offset_y", 0)
            if getattr(self, "canvas_img_id", None) is not None:
                self.canvas.coords(self.canvas_img_id, draw_x0 + self.pan_offset_x, draw_y0 + self.pan_offset_y)
            self._update_scrollbars()

    def _on_canvas_b1_release(self, event):
        if self.dragging_curtain:
            self._render_drawing(center_ref=False)
        self.dragging_curtain = False
        self.canvas.config(cursor="")

    def _on_canvas_b2_press(self, event):
        self.canvas.config(cursor="fleur")
        self._pan_start_x = event.x
        self._pan_start_y = event.y
        self._pan_init_offset_x = self.pan_offset_x
        self._pan_init_offset_y = self.pan_offset_y

    def _on_canvas_b2_motion(self, event):
        dx = event.x - self._pan_start_x
        dy = event.y - self._pan_start_y
        self.pan_offset_x = self._pan_init_offset_x + dx
        self.pan_offset_y = self._pan_init_offset_y + dy
        draw_x0 = getattr(self, "draw_offset_x", 0)
        draw_y0 = getattr(self, "draw_offset_y", 0)
        if getattr(self, "canvas_img_id", None) is not None:
            self.canvas.coords(self.canvas_img_id, draw_x0 + self.pan_offset_x, draw_y0 + self.pan_offset_y)
        self._update_scrollbars()

    def _on_canvas_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    def _render_drawing(self, center_ref: bool = False):
        if not self.pdf_a_path and not self.pdf_b_path:
            return

        try:
            act_ref = self.active_item.get("ref_des") if self.active_item else None
            all_diffs = self.comp_result.get("diff_items", []) if self.comp_result else []

            res_img, meta = render_curtain_drawing_pair(
                pdf_path_a=self.pdf_a_path,
                pdf_path_b=self.pdf_b_path,
                page_idx_a=self.active_page_a,
                page_idx_b=self.active_page_b,
                zoom=self.zoom_level,
                split_ratio=self.split_ratio,
                mode=self.view_mode,
                diff_items=all_diffs,
                active_ref=act_ref,
                rotation=self.rotation_angle
            )
            self.current_meta = meta

            cw = max(200, self.canvas.winfo_width())
            ch = max(200, self.canvas.winfo_height())
            draw_x0 = max(0, (cw - meta["width"]) // 2)
            draw_y0 = max(0, (ch - meta["height"]) // 2)
            self.draw_offset_x = draw_x0
            self.draw_offset_y = draw_y0

            if center_ref and meta.get("active_center"):
                cx, cy = meta["active_center"]
                self.pan_offset_x = (cw // 2) - (draw_x0 + cx)
                self.pan_offset_y = (ch // 2) - (draw_y0 + cy)

            img_x = draw_x0 + self.pan_offset_x
            img_y = draw_y0 + self.pan_offset_y

            self.tk_canvas_img = ImageTk.PhotoImage(res_img)
            if getattr(self, "canvas_img_id", None) is not None:
                self.canvas.coords(self.canvas_img_id, img_x, img_y)
                self.canvas.itemconfig(self.canvas_img_id, image=self.tk_canvas_img)
            else:
                self.canvas.delete("all")
                self.canvas_img_id = self.canvas.create_image(img_x, img_y, image=self.tk_canvas_img, anchor="nw", tags="drawing")
            self._update_scrollbars()
        except Exception as e:
            print(f"Error rendering drawing: {e}")

    # ── Table Population & Selection ──────────────────────────────────────────
    def _on_filter_change(self, choice: str):
        if "tất cả" in choice.lower() or "全部" in choice or "all" in choice.lower():
            self.current_filter = "diff_all"
        elif "thêm" in choice.lower() or "新增" in choice or "added" in choice.lower():
            self.current_filter = "added"
        elif "bớt" in choice.lower() or "删除" in choice or "removed" in choice.lower():
            self.current_filter = "removed"
        elif "đổi" in choice.lower() or "变更" in choice or "changed" in choice.lower():
            self.current_filter = "changed"
        elif "khớp" in choice.lower() or "一致" in choice or "match" in choice.lower():
            self.current_filter = "match"
        self._populate_tree()

    def _on_search_key(self, event=None):
        self.search_query = self.search_entry.get().strip().upper()
        self._populate_tree()

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        if not self.comp_result:
            self.empty_overlay.place(relx=0.5, rely=0.5, anchor="center")
            return

        self.empty_overlay.place_forget()

        all_diffs = self.comp_result.get("diff_items", [])
        all_matches = self.comp_result.get("match_items", [])

        # Filter by active tab
        if self.current_filter == "diff_all":
            items = all_diffs
        elif self.current_filter == "added":
            items = [it for it in all_diffs if it.get("status") == "ADDED"]
        elif self.current_filter == "removed":
            items = [it for it in all_diffs if it.get("status") == "REMOVED"]
        elif self.current_filter == "changed":
            items = [it for it in all_diffs if it.get("status") == "CHANGED"]
        elif self.current_filter == "match":
            items = all_matches
        else:
            items = all_diffs

        # Filter by selected stage if not ALL
        if self.current_stage_id != "ALL":
            target_stage = normalize_process_stage(self.current_stage_id)
            items = [it for it in items if normalize_process_stage(it.get("stage", "")) == target_stage]

        # Apply search query
        q = self.search_query
        if q:
            items = [
                it for it in items
                if q in str(it.get("ref_des", "")).upper() or
                   q in str(it.get("part_a", "")).upper() or
                   q in str(it.get("part_b", "")).upper() or
                   q in str(it.get("rating_a", "")).upper() or
                   q in str(it.get("rating_b", "")).upper() or
                   q in str(it.get("stage", "")).upper()
            ]

        self.lbl_table_count.configure(text=f"{len(items)} dòng")

        for idx, it in enumerate(items, start=1):
            st = it.get("status", "MATCH")
            tag = "tag_added" if st == "ADDED" else ("tag_removed" if st == "REMOVED" else ("tag_changed" if st == "CHANGED" else "tag_match"))

            disp_status = "🟢 LẮP THÊM" if st == "ADDED" else ("🔴 BỎ BỚT" if st == "REMOVED" else ("🟡 ĐỔI TRỊ SỐ" if st == "CHANGED" else "⚪ TRÙNG KHỚP"))

            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                it.get("ref_des", ""),
                it.get("stage", ""),
                disp_status,
                it.get("part_a", ""),
                it.get("rating_a", ""),
                it.get("part_b", ""),
                it.get("rating_b", ""),
                it.get("note", "")
            ))

        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
            self._on_tree_select()

    def _on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals or len(vals) < 3 or not self.comp_result:
            return

        ref_des = vals[1]
        stage = vals[2]
        all_items = self.comp_result.get("all_items", [])
        matched = [it for it in all_items if it.get("ref_des") == ref_des and it.get("stage") == stage]
        if matched:
            self.active_item = matched[0]
            # Check if page matches current stage
            norm_stage = normalize_process_stage(stage)
            cat_match = next((c for c in self.drawing_catalog if c["id"] != "ALL" and normalize_process_stage(c["stage"]) == norm_stage), None)
            if cat_match:
                self.active_page_a = cat_match["page_idx_a"]
                self.active_page_b = cat_match["page_idx_b"]
            self._update_inspector()

    def _update_inspector(self, skip_fullscreen: bool = False):
        if not self.active_item or not self.comp_result:
            return
        it = self.active_item
        ref = it.get("ref_des", "")
        stage = it.get("stage", "SMT")
        status = it.get("status", "CHANGED")

        self.lbl_insp_title.configure(text=f"🔍  {self.t('title_drawing_inspector', ref=ref)}")

        # Update details box
        st_color = "#10B981" if status == "ADDED" else ("#EF4444" if status == "REMOVED" else ("#F59E0B" if status == "CHANGED" else "#3B82F6"))
        self.lbl_det_ref.configure(text=f"Vị trí: {ref}   •   Công đoạn: {stage}   •   Trang: {self.active_page_a + 1}")
        self.lbl_det_status.configure(text=f"Trạng thái: {status}", text_color=st_color)
        self.lbl_det_part_a.configure(text=f"Model A: {it.get('part_a')}   |   {it.get('rating_a')}")
        self.lbl_det_part_b.configure(text=f"Model B: {it.get('part_b')}   |   {it.get('rating_b')}")
        self.lbl_det_note.configure(text=f"Ghi chú: {it.get('note')}")

        # Trigger render and auto-center
        self._render_drawing(center_ref=True)

        # Synchronize with fullscreen modal if currently open
        if not skip_fullscreen and self.fullscreen_window is not None:
            try:
                self.fullscreen_window.set_active_item(it)
            except Exception:
                self.fullscreen_window = None

    # ── Export & Reset ────────────────────────────────────────────────────────
    def _export_ecn(self):
        if not self.comp_result:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fn = filedialog.asksaveasfilename(
            title=self.t("btn_export_model_ecn"),
            initialfile=f"Bao_Cao_ECN_Model_Diff_{ts}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if fn:
            try:
                out = export_model_comparison_excel(self.comp_result, fn)
                messagebox.showinfo("Xuất Thành Công", f"Đã xuất báo cáo ECN thành công:\n{out}")
                self.app.show_toast("Đã xuất báo cáo ECN!")
            except Exception as e:
                messagebox.showerror("Lỗi Xuất File", f"Không thể xuất file: {e}")

    def _reset(self):
        if self.fullscreen_window is not None:
            try:
                self.fullscreen_window.destroy()
            except Exception:
                pass
            self.fullscreen_window = None

        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.pdf_a_path = None
        self.pdf_b_path = None
        self.comp_result = None
        self.active_item = None
        self.drawing_catalog = []
        self.current_stage_id = "ALL"
        self.lbl_model_a_name.configure(text=self.t("model_a_no_file"))
        self.lbl_model_a_sub.configure(text="")
        self.lbl_model_b_name.configure(text=self.t("model_b_no_file"))
        self.lbl_model_b_sub.configure(text="")
        self.lbl_verdict_title.configure(text=self.t("verdict_model_placeholder"), text_color=TEXT_MUTED)
        self.lbl_verdict_sub.configure(text="")
        self.card_verdict.configure(border_color=BORDER_CLR)
        self.tree.delete(*self.tree.get_children())
        self.empty_overlay.place(relx=0.5, rely=0.5, anchor="center")
        self.empty_canvas_overlay.place(relx=0.5, rely=0.5, anchor="center")
        self.canvas.delete("all")
        self.lbl_det_ref.configure(text="Vị trí: ---  |  Công đoạn: ---")
        self.lbl_det_status.configure(text="Trạng thái: ---", text_color=TEXT_MUTED)
        self.lbl_det_part_a.configure(text="Model A: ---")
        self.lbl_det_part_b.configure(text="Model B: ---")
        self.lbl_det_note.configure(text="Đánh giá: ---")
        self.btn_export_model_ecn.configure(state="disabled")
        self.opt_stage_picker.configure(values=["🌐 Tất Cả Công Đoạn (Tự Động)"])
        self.opt_stage_picker.set("🌐 Tất Cả Công Đoạn (Tự Động)")
        self._is_scanning = False
        if getattr(self, "_scan_hud", None) is not None:
            try:
                self._scan_hud.destroy()
            except Exception:
                pass
            self._scan_hud = None
        self.canvas.delete("ai_scan_anim")
        self.canvas.delete("ai_scan_flash")
        clear_curtain_cache()
        self.canvas_img_id = None
        self.rotation_angle = 0
        if hasattr(self, "btn_rotate"):
            self.btn_rotate.configure(text="🔄 Xoay")
        if hasattr(self, "row_models") and not self.row_models.winfo_ismapped():
            self.row_models.pack(fill="x", pady=(0, 6), before=self.card_compare_box)
        self.app._refresh_badges()

    # ── Theme & Language ──────────────────────────────────────────────────────
    def apply_theme(self, is_dark: bool):
        card_bg = BG_CARD[1] if is_dark else BG_CARD[0]
        surf_bg = BG_SURFACE[1] if is_dark else BG_SURFACE[0]
        bdr_clr = BORDER_CLR[1] if is_dark else BORDER_CLR[0]
        self.tree_container.configure(bg="#131929" if is_dark else "#FFFFFF")
        self.canvas.configure(bg="#0B0F1A" if is_dark else "#FFFFFF")

        self.workspace_card.configure(fg_color=card_bg, border_color=bdr_clr)
        self.card_model_a.configure(fg_color=card_bg, border_color=bdr_clr)
        self.card_model_b.configure(fg_color=card_bg, border_color=bdr_clr)
        self.card_compare_box.configure(fg_color=card_bg, border_color=bdr_clr)
        self.inspector_card.configure(fg_color=surf_bg, border_color=bdr_clr)
        self.canvas_container.configure(fg_color=card_bg, border_color=bdr_clr)
        self.det_box.configure(fg_color=card_bg, border_color=bdr_clr)
        self.empty_overlay.configure(fg_color=card_bg)

        # Tags
        self.tree.tag_configure("tag_added", background="#062E1E" if is_dark else "#DCFCE7")
        self.tree.tag_configure("tag_removed", background="#381010" if is_dark else "#FEE2E2")
        self.tree.tag_configure("tag_changed", background="#3B2A05" if is_dark else "#FEF3C7")
        self.tree.tag_configure("tag_match", background="#131929" if is_dark else "#FFFFFF")

    def update_language(self):
        self.lbl_card_a_title.configure(text=self.t("card_model_a_title"))
        self.lbl_card_b_title.configure(text=self.t("card_model_b_title"))
        self.btn_choose_model_a.configure(text=self.t("btn_choose_model_a"))
        self.btn_choose_model_b.configure(text=self.t("btn_choose_model_b"))
        if not self.pdf_a_path:
            self.lbl_model_a_name.configure(text=self.t("model_a_no_file"))
        if not self.pdf_b_path:
            self.lbl_model_b_name.configure(text=self.t("model_b_no_file"))
        self.btn_run_model_comp.configure(text=self.t("btn_run_model_comp"))
        self.btn_export_model_ecn.configure(text=self.t("btn_export_model_ecn"))
        self.btn_reset_model_comp.configure(text=self.t("btn_reset_model_comp"))

        modes = [self.t("disp_mode_table"), self.t("disp_mode_drawing"), self.t("disp_mode_split")]
        curr = self.display_mode
        self.seg_display_mode.configure(values=modes)
        if curr == "table":
            self.seg_display_mode.set(self.t("disp_mode_table"))
        elif curr == "drawing":
            self.seg_display_mode.set(self.t("disp_mode_drawing"))
        else:
            self.seg_display_mode.set(self.t("disp_mode_split"))

        self.btn_fullscreen.configure(text=self.t("btn_fullscreen_drawing"))
        if hasattr(self, "btn_rotate"):
            rot_str = "旋转" if self.app.current_lang == "zh" else ("Rotate" if self.app.current_lang == "en" else "Xoay")
            self.btn_rotate.configure(text=f"🔄 {self.rotation_angle}°" if self.rotation_angle > 0 else f"🔄 {rot_str}")

        for cid, width, anchor in self.col_specs:
            hdr_text = self.t("model_cols").get(cid, cid)
            self.tree.heading(cid, text=hdr_text)



# ─── SERIES BOM COMPARE VIEW (TAB 4) ─────────────────────────────────────────
class SeriesBOMCompareView(ctk.CTkFrame):
    """
    Dedicated view for comparing 2 BOMs of the same model across different series.
    Supports ERP Multi-level BOM PDFs and Excel BOMs.
    Features the QC Focus Checklist, Working Manual PCB Drawing Upload with
    multi-color border highlighting (Green: Added, Red: Removed/DNP, Amber: Modified),
    and Holographic Cyberpunk AI Scan Loading Animation.
    """
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app

        self.file_a = None
        self.file_b = None
        self.file_drawing = None
        self.drawing_page_idx = 0
        self.drawing_pages = []

        self.bom_a_data = None
        self.bom_b_data = None
        self.comparison_result = None

        self.display_mode = "split"  # 'table', 'drawing', 'split'
        self.filter_mode = "focus"
        self.search_query = ""
        self.is_comparing = False

        self.qc_status_map = {}  # loc -> 'OK', 'NG', 'PENDING'
        self.active_loc = None

        # Drawing canvas state
        self.zoom_level = 1.0
        self.rotation_angle = 0
        self.pixel_coords_map = {}
        self.current_annotated_img = None
        self.tk_canvas_img = None
        self.canvas_img_id = None
        self.pan_start_x = 0
        self.pan_start_y = 0

        self._loading_hud = None

        self._build_ui()

    def t(self, key: str, **kwargs) -> str:
        return self.app.t(key, **kwargs)

    def _build_ui(self):
        # ── 1. TOP CONTAINER: 3 UPLOAD CARDS + ACTIONS ───────────────────────
        top_container = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                                     border_width=1, border_color=BORDER_CLR)
        top_container.pack(fill="x", padx=4, pady=(4, 6))

        inner_top = ctk.CTkFrame(top_container, fg_color="transparent")
        inner_top.pack(fill="x", padx=10, pady=8)

        # Card 1: Series A BOM
        self.card_a = ctk.CTkFrame(inner_top, fg_color=BG_SURFACE, corner_radius=8,
                                   border_width=1, border_color=BORDER_CLR)
        self.card_a.pack(side="left", fill="both", expand=True, padx=(0, 4))

        a_head = ctk.CTkFrame(self.card_a, fg_color="transparent")
        a_head.pack(fill="x", padx=8, pady=(6, 2))
        ctk.CTkLabel(a_head, text="1. BOM GỐC (A)", font=("Segoe UI", 10, "bold"),
                     text_color=ACCENT_BLUE).pack(side="left")
        ctk.CTkButton(
            a_head, text="📂 Chọn BOM A", font=("Segoe UI", 9, "bold"),
            height=24, width=95, fg_color=ACCENT_BLUE, hover_color="#1E40AF",
            command=self._select_file_a
        ).pack(side="right")

        self.lbl_file_a_name = ctk.CTkLabel(
            self.card_a, text="Chưa chọn BOM Series A", font=("Segoe UI", 9),
            text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_file_a_name.pack(fill="x", padx=8, pady=(1, 1))

        self.lbl_file_a_meta = ctk.CTkLabel(
            self.card_a, text="Model: — | Series: —", font=("Segoe UI", 8, "bold"),
            text_color=ACCENT_TEAL, anchor="w"
        )
        self.lbl_file_a_meta.pack(fill="x", padx=8, pady=(0, 4))

        # Card 2: Series B BOM
        self.card_b = ctk.CTkFrame(inner_top, fg_color=BG_SURFACE, corner_radius=8,
                                   border_width=1, border_color=BORDER_CLR)
        self.card_b.pack(side="left", fill="both", expand=True, padx=(4, 4))

        b_head = ctk.CTkFrame(self.card_b, fg_color="transparent")
        b_head.pack(fill="x", padx=8, pady=(6, 2))
        ctk.CTkLabel(b_head, text="2. BOM SO SÁNH (B)", font=("Segoe UI", 10, "bold"),
                     text_color=ACCENT_TEAL).pack(side="left")
        ctk.CTkButton(
            b_head, text="📂 Chọn BOM B", font=("Segoe UI", 9, "bold"),
            height=24, width=95, fg_color=ACCENT_TEAL, hover_color="#0D9488",
            command=self._select_file_b
        ).pack(side="right")

        self.lbl_file_b_name = ctk.CTkLabel(
            self.card_b, text="Chưa chọn BOM Series B", font=("Segoe UI", 9),
            text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_file_b_name.pack(fill="x", padx=8, pady=(1, 1))

        self.lbl_file_b_meta = ctk.CTkLabel(
            self.card_b, text="Model: — | Series: —", font=("Segoe UI", 8, "bold"),
            text_color=ACCENT_TEAL, anchor="w"
        )
        self.lbl_file_b_meta.pack(fill="x", padx=8, pady=(0, 4))

        # Card 3: Working Manual PCB Drawing PDF
        self.card_dwg = ctk.CTkFrame(inner_top, fg_color=BG_SURFACE, corner_radius=8,
                                    border_width=1, border_color=BORDER_CLR)
        self.card_dwg.pack(side="left", fill="both", expand=True, padx=(4, 6))

        dwg_head = ctk.CTkFrame(self.card_dwg, fg_color="transparent")
        dwg_head.pack(fill="x", padx=8, pady=(6, 2))
        ctk.CTkLabel(dwg_head, text="3. BẢN VẼ PCB (PDF)", font=("Segoe UI", 10, "bold"),
                     text_color=ACCENT_AMBER).pack(side="left")
        ctk.CTkButton(
            dwg_head, text="📐 Chọn Bản Vẽ", font=("Segoe UI", 9, "bold"),
            height=24, width=95, fg_color=ACCENT_AMBER, hover_color="#B45309",
            command=self._select_file_drawing
        ).pack(side="right")

        self.lbl_file_dwg_name = ctk.CTkLabel(
            self.card_dwg, text="Chưa tải bản vẽ PCB (Tùy chọn)", font=("Segoe UI", 9),
            text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_file_dwg_name.pack(fill="x", padx=8, pady=(1, 1))

        dwg_sub = ctk.CTkFrame(self.card_dwg, fg_color="transparent")
        dwg_sub.pack(fill="x", padx=8, pady=(0, 4))

        ctk.CTkLabel(dwg_sub, text="Trang:", font=("Segoe UI", 8, "bold"), text_color=TEXT_MUTED).pack(side="left", padx=(0, 4))
        self.opt_dwg_page = ctk.CTkOptionMenu(
            dwg_sub, values=["Trang 1"], width=130, height=20, corner_radius=4,
            font=("Segoe UI", 8), command=self._on_drawing_page_selected
        )
        self.opt_dwg_page.pack(side="left")

        # Action Buttons Box
        act_box = ctk.CTkFrame(inner_top, fg_color="transparent")
        act_box.pack(side="right", fill="y", padx=(4, 0))

        self.btn_run_compare = ctk.CTkButton(
            act_box, text="⚡ SO SÁNH & ĐỊNH VỊ", font=("Segoe UI", 11, "bold"),
            height=34, width=155, fg_color=ACCENT_TEAL, hover_color="#0F766E",
            command=self._start_compare
        )
        self.btn_run_compare.pack(fill="x", pady=(1, 3))

        self.btn_export_fai = ctk.CTkButton(
            act_box, text="📥 Xuất Báo Cáo FAI", font=("Segoe UI", 9, "bold"),
            height=24, width=155, fg_color="#1E3A8A", hover_color="#1E40AF",
            state="disabled", command=self._export_excel
        )
        self.btn_export_fai.pack(fill="x", pady=(1, 0))

        # ── 2. VALIDATION & STATS BANNER ─────────────────────────────────────
        self.banner_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.banner_frame.pack(fill="x", padx=4, pady=(0, 4))

        self.lbl_validation = ctk.CTkLabel(
            self.banner_frame, text="💡 Vui lòng chọn 2 file BOM (PDF hoặc Excel) rồi nhấn '⚡ SO SÁNH & ĐỊNH VỊ'.",
            font=("Segoe UI", 9, "italic"), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_validation.pack(fill="x", padx=4)

        self.stats_container = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=8,
                                           border_width=1, border_color=BORDER_CLR)
        self.stats_container.pack(fill="x", padx=4, pady=(0, 6))

        self.kpi_boxes = {}
        kpi_defs = [
            ("total", "🔵 TỔNG VỊ TRÍ", "0", "#3D8EFF"),
            ("matched", "⚪ DÙNG CHUNG", "0 (0%)", "#9E9E9E"),
            ("added", "🟢 THÊM MỚI", "0", "#00E676"),
            ("removed", "🔴 BỎ TRỐNG (DNP)", "0", "#FF5252"),
            ("modified", "🟡 ĐỔI MÃ VẬT TƯ", "0", "#FFAB00"),
            ("focus", "🎯 CẦN KIỂM (QC FOCUS)", "0", "#00C9A7")
        ]

        for k_id, k_title, k_val, k_col in kpi_defs:
            f = ctk.CTkFrame(self.stats_container, fg_color=BG_SURFACE, corner_radius=6)
            f.pack(side="left", fill="both", expand=True, padx=3, pady=4)
            lbl_t = ctk.CTkLabel(f, text=k_title, font=("Segoe UI", 8, "bold"), text_color=TEXT_MUTED)
            lbl_t.pack(pady=(2, 0))
            lbl_v = ctk.CTkLabel(f, text=k_val, font=("Segoe UI", 11, "bold"), text_color=k_col)
            lbl_v.pack(pady=(0, 2))
            self.kpi_boxes[k_id] = (lbl_t, lbl_v, k_col)

        # ── 3. QC FOCUS CHECKLIST HERO CARD & DISPLAY MODE ───────────────────
        self.card_qc_focus = ctk.CTkFrame(
            self, fg_color=("#F0FDFA", "#13232C"), border_width=1.5,
            border_color=ACCENT_TEAL, corner_radius=8
        )
        self.card_qc_focus.pack(fill="x", padx=4, pady=(0, 6))

        qc_head = ctk.CTkFrame(self.card_qc_focus, fg_color="transparent")
        qc_head.pack(fill="x", padx=10, pady=(6, 2))

        ctk.CTkLabel(
            qc_head, text="🎯  BẢNG KIỂM SOÁT LINH KIỆN CẦN CHÚ Ý (QC FOCUS CHECKLIST)",
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_TEAL
        ).pack(side="left")

        # Display Mode Segmented button
        self.seg_display_mode = ctk.CTkSegmentedButton(
            qc_head, values=["📋 Bảng Danh Mục", "📐 Bản Vẽ PCB", "◫ Song Song"],
            height=24, corner_radius=6, font=("Segoe UI", 9, "bold"),
            command=self._on_display_mode_change
        )
        self.seg_display_mode.set("◫ Song Song")
        self.seg_display_mode.pack(side="right", padx=(8, 0))

        self.lbl_qc_progress = ctk.CTkLabel(
            qc_head, text="Tiến độ kiểm tra FAI: 0 / 0 (0%)",
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_AMBER
        )
        self.lbl_qc_progress.pack(side="right", padx=10)

        qc_sub = ctk.CTkFrame(self.card_qc_focus, fg_color="transparent")
        qc_sub.pack(fill="x", padx=10, pady=(0, 3))

        ctk.CTkLabel(
            qc_sub, text="⚡ Click vào dòng linh kiện để định vị trên bản vẽ PCB. Màu xanh: Thêm mới | Đỏ: Bỏ trống (DNP) | Vàng: Đổi mã.",
            font=("Segoe UI", 8, "italic"), text_color=TEXT_MUTED
        ).pack(side="left")

        self.btn_mark_all_ok = ctk.CTkButton(
            qc_sub, text="✓ Đánh dấu tất cả OK", font=("Segoe UI", 8, "bold"),
            height=20, width=115, fg_color=ACCENT_TEAL, hover_color="#0D9488",
            command=self._mark_all_ok
        )
        self.btn_mark_all_ok.pack(side="right", padx=(4, 0))

        self.btn_reset_qc = ctk.CTkButton(
            qc_sub, text="↺ Đặt lại", font=("Segoe UI", 8),
            height=20, width=70, fg_color=BG_SURFACE, hover_color=BG_HOVER,
            text_color=TEXT_PRIMARY, command=self._reset_qc_checks
        )
        self.btn_reset_qc.pack(side="right")

        self.pbar_qc = ctk.CTkProgressBar(
            self.card_qc_focus, height=6, corner_radius=3, progress_color=ACCENT_TEAL
        )
        self.pbar_qc.pack(fill="x", padx=10, pady=(1, 6))
        self.pbar_qc.set(0)

        # ── 4. MAIN SPLIT WORKSPACE: TABLE (LEFT) & DRAWING CANVAS (RIGHT) ───
        self.split_workspace = ctk.CTkFrame(self, fg_color="transparent")
        self.split_workspace.pack(fill="both", expand=True, padx=4, pady=(0, 2))

        # ── LEFT PANEL: TABLE VIEW ───────────────────────────────────────────
        self.table_panel = ctk.CTkFrame(self.split_workspace, fg_color="transparent")
        self.table_panel.pack(side="left", fill="both", expand=True, padx=(0, 3))

        # Filter Bar for table
        tbl_filter_bar = ctk.CTkFrame(self.table_panel, fg_color="transparent")
        tbl_filter_bar.pack(fill="x", pady=(0, 4))

        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            tbl_filter_bar, placeholder_text="🔍 Tìm vị trí (Ref Des), mã LK...",
            width=200, height=28, corner_radius=6, textvariable=self.search_var
        )
        self.search_entry.pack(side="left", padx=(0, 4))
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate_table())

        self.filter_var = ctk.StringVar(value="🎯 Cần chú ý (QC Focus)")
        self.seg_filter = ctk.CTkSegmentedButton(
            tbl_filter_bar,
            values=["🎯 Cần chú ý", "🟢 Thêm", "🔴 Bớt", "🟡 Đổi", "⚪ Khớp", "Tất cả"],
            variable=self.filter_var, height=28, corner_radius=6,
            command=lambda v: self._populate_table()
        )
        self.seg_filter.pack(side="left", fill="x", expand=True)

        table_box = ctk.CTkFrame(self.table_panel, fg_color=BG_CARD, corner_radius=8,
                                 border_width=1, border_color=BORDER_CLR)
        table_box.pack(fill="both", expand=True)

        cols = ("check", "loc", "status", "part_a", "part_b", "action", "spec")
        self.tree = ttk.Treeview(table_box, columns=cols, show="headings", selectmode="browse")

        self.tree.heading("check", text="[QC Kiểm]")
        self.tree.heading("loc", text="Vị Trí (Ref)")
        self.tree.heading("status", text="Phân Loại")
        self.tree.heading("part_a", text="Mã LK (Series A)")
        self.tree.heading("part_b", text="Mã LK (Series B)")
        self.tree.heading("action", text="Chỉ Dẫn Hành Động Cho QC (Action Guide)")
        self.tree.heading("spec", text="Quy Cách (Series B)")

        self.tree.column("check", width=85, anchor="center")
        self.tree.column("loc", width=75, anchor="center")
        self.tree.column("status", width=105, anchor="center")
        self.tree.column("part_a", width=125, anchor="w")
        self.tree.column("part_b", width=125, anchor="w")
        self.tree.column("action", width=300, anchor="w")
        self.tree.column("spec", width=180, anchor="w")

        vsb_tbl = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        hsb_tbl = ttk.Scrollbar(table_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb_tbl.set, xscrollcommand=hsb_tbl.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb_tbl.pack(side="right", fill="y")
        hsb_tbl.pack(side="bottom", fill="x")

        self._apply_tree_tags()
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # ── RIGHT PANEL: PCB DRAWING CANVAS ──────────────────────────────────
        self.drawing_panel = ctk.CTkFrame(self.split_workspace, fg_color=BG_CARD, corner_radius=8,
                                         border_width=1, border_color=BORDER_CLR)
        self.drawing_panel.pack(side="right", fill="both", expand=True, padx=(3, 0))

        # Drawing Toolbar
        dwg_toolbar = ctk.CTkFrame(self.drawing_panel, fg_color=BG_SURFACE, height=32, corner_radius=6)
        dwg_toolbar.pack(fill="x", padx=6, pady=6)

        ctk.CTkLabel(
            dwg_toolbar, text="📐 BẢN VẼ PCB HIGHLIGHT", font=("Segoe UI", 10, "bold"),
            text_color=ACCENT_TEAL
        ).pack(side="left", padx=8)

        # Zoom buttons
        ctk.CTkButton(
            dwg_toolbar, text="🔍 -", width=26, height=22, font=("Consolas", 10, "bold"),
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            command=self._zoom_out
        ).pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(dwg_toolbar, text="100%", font=("Segoe UI", 9, "bold"), width=42)
        self.lbl_zoom.pack(side="left")

        ctk.CTkButton(
            dwg_toolbar, text="🔍 +", width=26, height=22, font=("Consolas", 10, "bold"),
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            command=self._zoom_in
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            dwg_toolbar, text="1:1", width=30, height=22, font=("Segoe UI", 8),
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            command=self._zoom_reset
        ).pack(side="left", padx=2)

        # Rotate button
        self.btn_rotate_dwg = ctk.CTkButton(
            dwg_toolbar, text="🔄 0°", width=50, height=22, font=("Segoe UI", 8, "bold"),
            fg_color=BG_CARD, hover_color=BG_HOVER, text_color=TEXT_PRIMARY,
            command=self._rotate_drawing
        )
        self.btn_rotate_dwg.pack(side="left", padx=4)

        # Active location indicator badge
        self.lbl_spotlight_badge = ctk.CTkLabel(
            dwg_toolbar, text="📍 Chưa chọn linh kiện", font=("Segoe UI", 9, "bold"),
            text_color=TEXT_MUTED
        )
        self.lbl_spotlight_badge.pack(side="left", padx=8)

        # Fullscreen button
        ctk.CTkButton(
            dwg_toolbar, text="⛶ Toàn Màn Hình", width=105, height=22,
            font=("Segoe UI", 9, "bold"), fg_color=ACCENT_BLUE, hover_color="#1E40AF",
            command=self._open_fullscreen_drawing
        ).pack(side="right", padx=6)

        # Canvas Frame
        canvas_box = ctk.CTkFrame(self.drawing_panel, fg_color="#0A0E18", corner_radius=6)
        canvas_box.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        self.canvas = tk.Canvas(
            canvas_box, bg="#0A0E18", highlightthickness=0, cursor="hand2"
        )
        self.h_sb_dwg = ttk.Scrollbar(canvas_box, orient="horizontal", command=self.canvas.xview)
        self.v_sb_dwg = ttk.Scrollbar(canvas_box, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.h_sb_dwg.set, yscrollcommand=self.v_sb_dwg.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_sb_dwg.pack(side="right", fill="y")
        self.h_sb_dwg.pack(side="bottom", fill="x")

        # Canvas pan & zoom bindings
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<MouseWheel>", self._on_canvas_mousewheel)

    # ── THEME & VISUAL TAGS ──────────────────────────────────────────────────
    def _apply_tree_tags(self):
        is_dark = (self.app.current_theme == "dark")
        if is_dark:
            self.tree.tag_configure("tag_added", background="#15362B", foreground="#00E676")
            self.tree.tag_configure("tag_removed", background="#361818", foreground="#FF5252")
            self.tree.tag_configure("tag_modified", background="#362E15", foreground="#FFAB00")
            self.tree.tag_configure("tag_matched", background="#131929", foreground="#7A8BA6")
            self.tree.tag_configure("tag_ok", background="#0E3D2F", foreground="#A3E4D7")
        else:
            self.tree.tag_configure("tag_added", background="#E8F8F5", foreground="#0E6251")
            self.tree.tag_configure("tag_removed", background="#FDEDEC", foreground="#78281F")
            self.tree.tag_configure("tag_modified", background="#FEF9E7", foreground="#7D6608")
            self.tree.tag_configure("tag_matched", background="#FFFFFF", foreground="#212529")
            self.tree.tag_configure("tag_ok", background="#D4EFDF", foreground="#145A32")

    # ── FILE SELECTION & MODEL VALIDATION HANDLERS ───────────────────────────
    def _get_current_bom_model(self) -> str:
        """Returns the base model of the current BOMs (e.g. 'CHA3259AF')."""
        if self.comparison_result and self.comparison_result.get("summary"):
            s = self.comparison_result["summary"]
            m = s.get("model_b") or s.get("model_a")
            if m:
                return m
        if self.file_b:
            mod, _ = sbc.extract_model_and_series(os.path.basename(self.file_b))
            if mod:
                return mod
        if self.file_a:
            mod, _ = sbc.extract_model_and_series(os.path.basename(self.file_a))
            if mod:
                return mod
        return ""

    def _reset_drawing(self):
        """Clears the loaded drawing and resets canvas to empty notice."""
        self.file_drawing = None
        self.drawing_pages = []
        self.drawing_page_idx = 0
        self.current_annotated_img = None
        self.pixel_coords_map = {}
        self.lbl_file_dwg_name.configure(text="Chưa chọn bản vẽ")
        self.lbl_file_dwg_meta.configure(text="PDF Working Manual (PCB)", text_color=TEXT_MUTED)
        self.opt_dwg_page.configure(values=["Trang 1"])
        self.opt_dwg_page.set("Trang 1")
        self.canvas.delete("all")
        self.canvas.create_text(
            300, 200, text="Chưa tải file Bản vẽ Working Manual (PDF)\n\nNhấn '📐 Chọn Bản Vẽ' ở trên để xem bản vẽ PCB\nvới khung viền highlight linh kiện chênh lệch.",
            fill="#64748B", font=("Segoe UI", 11, "bold"), justify="center"
        )

    def _select_file_a(self):
        f = filedialog.askopenfilename(
            title="Chọn file BOM Series A (Gốc)",
            filetypes=[("BOM Files", "*.pdf;*.xlsx;*.xls"), ("PDF Files", "*.pdf"), ("Excel Files", "*.xlsx;*.xls")]
        )
        if not f:
            return
        if os.path.getsize(f) > MAX_FILE_SIZE_BYTES:
            messagebox.showwarning("File quá lớn", f"File '{os.path.basename(f)}' vượt quá 10MB!")
            return
        self.file_a = f
        self.lbl_file_a_name.configure(text=f"{os.path.basename(f)} ({format_file_size(os.path.getsize(f))})")
        model, series = sbc.extract_model_and_series(os.path.basename(f))
        self.lbl_file_a_meta.configure(text=f"Model: {model} | Series: {series or 'Gốc'}")

        # If drawing is already loaded, verify that drawing matches the new BOM model
        if self.file_drawing:
            cur_bom_mod = self._get_current_bom_model()
            is_valid, msg, dinfo = sbc.validate_drawing_against_bom(self.file_drawing, cur_bom_mod)
            if not is_valid:
                messagebox.showwarning(
                    "Bản Vẽ Không Khớp BOM Mới",
                    f"⚠️ Bản vẽ '{os.path.basename(self.file_drawing)}' đã chọn trước đó không thuộc Model '{cur_bom_mod}'.\n\n"
                    f"Hệ thống sẽ đặt lại file bản vẽ. Vui lòng chọn bản vẽ đúng của Model {cur_bom_mod}."
                )
                self._reset_drawing()

    def _select_file_b(self):
        f = filedialog.askopenfilename(
            title="Chọn file BOM Series B (So sánh)",
            filetypes=[("BOM Files", "*.pdf;*.xlsx;*.xls"), ("PDF Files", "*.pdf"), ("Excel Files", "*.xlsx;*.xls")]
        )
        if not f:
            return
        if os.path.getsize(f) > MAX_FILE_SIZE_BYTES:
            messagebox.showwarning("File quá lớn", f"File '{os.path.basename(f)}' vượt quá 10MB!")
            return
        self.file_b = f
        self.lbl_file_b_name.configure(text=f"{os.path.basename(f)} ({format_file_size(os.path.getsize(f))})")
        model, series = sbc.extract_model_and_series(os.path.basename(f))
        self.lbl_file_b_meta.configure(text=f"Model: {model} | Series: {series or 'Mới'}")

        # If drawing is already loaded, verify that drawing matches the new BOM model
        if self.file_drawing:
            cur_bom_mod = self._get_current_bom_model()
            is_valid, msg, dinfo = sbc.validate_drawing_against_bom(self.file_drawing, cur_bom_mod)
            if not is_valid:
                messagebox.showwarning(
                    "Bản Vẽ Không Khớp BOM Mới",
                    f"⚠️ Bản vẽ '{os.path.basename(self.file_drawing)}' đã chọn trước đó không thuộc Model '{cur_bom_mod}'.\n\n"
                    f"Hệ thống sẽ đặt lại file bản vẽ. Vui lòng chọn bản vẽ đúng của Model {cur_bom_mod}."
                )
                self._reset_drawing()

    def _select_file_drawing(self):
        f = filedialog.askopenfilename(
            title="Chọn File Bản Vẽ Working Manual (PDF)",
            filetypes=[("PDF Drawing Manual", "*.pdf"), ("Tất Cả Files", "*.*")]
        )
        if not f:
            return
        if os.path.getsize(f) > MAX_FILE_SIZE_BYTES:
            messagebox.showwarning("File quá lớn", f"Bản vẽ '{os.path.basename(f)}' vượt quá 10MB!")
            return

        bom_model = self._get_current_bom_model()
        is_valid, val_msg, dwg_info = sbc.validate_drawing_against_bom(f, bom_model)

        if not is_valid:
            dwg_disp = dwg_info.get("display_model", "Không xác định")
            messagebox.showerror(
                "Bản Vẽ Không Khớp Model",
                f"❌ Bản vẽ không khớp với Model cần so sánh!\n\n"
                f"• Model của BOM: {bom_model or '(Chưa chọn BOM)'}\n"
                f"• Model trên Bản vẽ: {dwg_disp}\n"
                f"• Chi tiết: {val_msg}\n"
                f"• File bản vẽ: {os.path.basename(f)}\n\n"
                f"Vui lòng tải lên đúng file bản vẽ Working Manual của Model {bom_model}!"
            )
            return

        self.file_drawing = f
        self.lbl_file_dwg_name.configure(text=f"{os.path.basename(f)} ({format_file_size(os.path.getsize(f))})")
        dwg_mod = dwg_info.get("display_model", "")
        self.lbl_file_dwg_meta.configure(
            text=f"✓ Model: {dwg_mod or bom_model} (Hợp lệ)",
            text_color="#00E676"
        )

        # Load drawing pages info
        try:
            self.drawing_pages = sbc.get_drawing_pages_info(f)
            if self.drawing_pages:
                vals = [p["label"] for p in self.drawing_pages]
                self.opt_dwg_page.configure(values=vals)
                self.opt_dwg_page.set(vals[0])
                self.drawing_page_idx = 0
            self._render_drawing()
        except Exception as e:
            print(f"Error loading drawing pages: {e}")

    def _on_drawing_page_selected(self, choice):
        if not self.drawing_pages:
            return
        for p in self.drawing_pages:
            if p["label"] == choice:
                self.drawing_page_idx = p["index"]
                break
        self._render_drawing()

    # ── HOLOGRAPHIC CYBERPUNK HUD LOADING CARD ───────────────────────────────
    def _show_loading_hud(self, title: str = "VIPQC AI — SCANNING & LOCATING PCB COMPONENTS"):
        self._hide_loading_hud()
        is_dark = (self.app.current_theme == "dark")

        self._loading_hud = ctk.CTkFrame(
            self, fg_color="#080F1E" if is_dark else "#0F172A",
            corner_radius=12, border_width=1.5, border_color="#00F0FF"
        )
        self._loading_hud.place(relx=0.5, rely=0.45, anchor="center")

        hud_inner = ctk.CTkFrame(self._loading_hud, fg_color="transparent")
        hud_inner.pack(padx=28, pady=20)

        # Header Row
        h_row = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(h_row, text="⚡", font=("Segoe UI", 16)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            h_row, text=title, font=("Segoe UI", 11, "bold"), text_color="#00F0FF"
        ).pack(side="left", padx=(0, 14))
        ctk.CTkLabel(
            h_row, text="60 FPS • RADAR SCAN", font=("Consolas", 9, "bold"), text_color="#38BDF8"
        ).pack(side="right")

        # Progress Bar Row
        p_row = ctk.CTkFrame(hud_inner, fg_color="transparent")
        p_row.pack(fill="x", pady=(0, 10))

        self.hud_pbar = ctk.CTkProgressBar(
            p_row, width=380, height=10, corner_radius=5,
            fg_color="#1E293B", progress_color="#00F0FF"
        )
        self.hud_pbar.set(0.1)
        self.hud_pbar.pack(side="left", padx=(0, 10))

        self.hud_lbl_pct = ctk.CTkLabel(
            p_row, text="10%", font=("Consolas", 11, "bold"), text_color="#00F0FF", width=42
        )
        self.hud_lbl_pct.pack(side="left")

        # Telemetry Phase Line
        self.hud_lbl_telemetry = ctk.CTkLabel(
            hud_inner, text="[PHASE 1/3] Đang bóc tách danh mục vật tư đa cấp BOM A & BOM B...",
            font=("Consolas", 9), text_color="#94A3B8", anchor="w"
        )
        self.hud_lbl_telemetry.pack(fill="x")

    def _update_loading_hud(self, pct: float, telemetry: str):
        if not self._loading_hud:
            return
        self.hud_pbar.set(pct)
        self.hud_lbl_pct.configure(text=f"{int(pct * 100)}%")
        self.hud_lbl_telemetry.configure(text=telemetry)

    def _hide_loading_hud(self):
        if self._loading_hud:
            try:
                self._loading_hud.destroy()
            except Exception:
                pass
            self._loading_hud = None

    # ── COMPARE ENGINE EXECUTION ─────────────────────────────────────────────
    def _start_compare(self):
        if not self.file_a or not self.file_b:
            messagebox.showwarning("Thiếu file", "Vui lòng chọn đầy đủ 2 file BOM Series A và Series B!")
            return

        # Validate drawing if user selected one
        if self.file_drawing:
            cur_bom_mod = self._get_current_bom_model()
            is_valid, msg, dinfo = sbc.validate_drawing_against_bom(self.file_drawing, cur_bom_mod)
            if not is_valid:
                messagebox.showerror(
                    "Bản Vẽ Không Khớp Model",
                    f"❌ Bản vẽ không khớp với Model cần so sánh!\n\n"
                    f"• Model BOM: {cur_bom_mod}\n"
                    f"• Model Bản vẽ: {dinfo.get('display_model', 'Không xác định')}\n\n"
                    f"Vui lòng tải lên đúng file bản vẽ Working Manual của Model {cur_bom_mod} trước khi so sánh!"
                )
                return

        self.btn_run_compare.configure(state="disabled", text="⏳ Đang phân tích...")
        self.is_comparing = True
        self._show_loading_hud()

        threading.Thread(target=self._run_compare_thread, daemon=True).start()

    def _run_compare_thread(self):
        try:
            time.sleep(0.3)
            self.after(0, lambda: self._update_loading_hud(0.35, "[PHASE 1/3] Đang bóc tách dữ liệu 2 BOM..."))

            bom_a = sbc.parse_any_bom(self.file_a)
            bom_b = sbc.parse_any_bom(self.file_b)

            time.sleep(0.3)
            self.after(0, lambda: self._update_loading_hud(0.70, "[PHASE 2/3] So sánh vị trí linh kiện & phân loại FAI..."))

            res = sbc.compare_series_boms(bom_a, bom_b)

            time.sleep(0.2)
            self.after(0, lambda: self._update_loading_hud(0.95, "[PHASE 3/3] Quét bản vẽ PCB và vẽ khung viền Highlight..."))

            time.sleep(0.3)
            self.after(0, lambda: self._on_compare_finished(res))
        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._on_compare_error(err_msg))

    def _on_compare_finished(self, result):
        self._hide_loading_hud()
        self.comparison_result = result
        self.btn_run_compare.configure(state="normal", text="⚡ SO SÁNH & ĐỊNH VỊ")
        self.btn_export_fai.configure(state="normal")
        self.is_comparing = False

        s = result["summary"]
        v_stat = s["validation_status"]
        v_msg = s["validation_msg"]

        if v_stat == "VALID":
            self.lbl_validation.configure(
                text=f"✅ {v_msg} (Model: {s['model_b']} | Series {s['series_a']} ➔ {s['series_b']})",
                text_color="#00E676"
            )
        else:
            self.lbl_validation.configure(text=f"⚠️ {v_msg}", text_color=ACCENT_AMBER)

        # Update KPI badges
        self.kpi_boxes["total"][1].configure(text=str(s["total_locations"]))
        self.kpi_boxes["matched"][1].configure(text=f"{s['matched_count']} ({s['match_percentage']}%)")
        self.kpi_boxes["added"][1].configure(text=str(s["added_count"]))
        self.kpi_boxes["removed"][1].configure(text=str(s["removed_count"]))
        self.kpi_boxes["modified"][1].configure(text=str(s["modified_count"]))
        self.kpi_boxes["focus"][1].configure(text=str(s["focus_count"]))

        # Update Series headings
        s_a = s["series_a"] or "A"
        s_b = s["series_b"] or "B"
        self.tree.heading("part_a", text=f"Mã LK ({s_a})")
        self.tree.heading("part_b", text=f"Mã LK ({s_b})")
        self.tree.heading("spec", text=f"Quy Cách / Spec ({s_b})")

        self.qc_status_map = {}
        self.filter_var.set("🎯 Cần chú ý")
        self._populate_table()

        # Render Drawing with Border Highlights
        self._render_drawing()

        self.app.set_status(f"Hoàn thành so sánh 2 BOM! {s['focus_count']} linh kiện cần chú ý kiểm tra.")

    def _on_compare_error(self, err_msg):
        self._hide_loading_hud()
        self.btn_run_compare.configure(state="normal", text="⚡ SO SÁNH & ĐỊNH VỊ")
        self.is_comparing = False
        self.lbl_validation.configure(text=f"❌ Lỗi khi so sánh: {err_msg}", text_color="#FF5252")
        messagebox.showerror("Lỗi So Sánh", f"Đã xảy ra lỗi khi phân tích BOM:\n\n{err_msg}")

    # ── DRAWING CANVAS & BORDER HIGHLIGHT RENDERING ──────────────────────────
    def _render_drawing(self):
        if not self.file_drawing or not os.path.exists(self.file_drawing):
            self.canvas.delete("all")
            self.canvas.create_text(
                300, 200, text="Chưa tải file Bản vẽ Working Manual (PDF)\n\nNhấn '📐 Chọn Bản Vẽ' ở trên để xem bản vẽ PCB\nvới khung viền highlight linh kiện chênh lệch.",
                fill="#64748B", font=("Segoe UI", 11, "bold"), justify="center"
            )
            return

        focus_items = self.comparison_result.get("qc_focus_items", []) if self.comparison_result else []

        img, coords_map = sbc.render_annotated_drawing_page(
            self.file_drawing,
            self.drawing_page_idx,
            focus_items,
            active_loc=self.active_loc,
            zoom=self.zoom_level,
            rotation=self.rotation_angle
        )

        if not img:
            return

        self.current_annotated_img = img
        self.pixel_coords_map = coords_map

        cw = max(200, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        draw_x0 = max(0, (cw - img.width) // 2) if img.width < cw else 0
        draw_y0 = max(0, (ch - img.height) // 2) if img.height < ch else 0
        self.draw_offset_x = draw_x0
        self.draw_offset_y = draw_y0

        self.tk_canvas_img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas_img_id = self.canvas.create_image(draw_x0, draw_y0, anchor="nw", image=self.tk_canvas_img)
        max_w = max(cw, img.width + draw_x0)
        max_h = max(ch, img.height + draw_y0)
        self.canvas.configure(scrollregion=(0, 0, max_w, max_h))

        # Center on active_loc if present
        if self.active_loc and self.active_loc in self.pixel_coords_map:
            x0, y0, x1, y1 = self.pixel_coords_map[self.active_loc]
            cx = (x0 + x1) / 2 + self.draw_offset_x
            cy = (y0 + y1) / 2 + self.draw_offset_y
            fx = max(0.0, min(1.0, (cx - cw / 2) / max_w))
            fy = max(0.0, min(1.0, (cy - ch / 2) / max_h))
            self.canvas.xview_moveto(fx)
            self.canvas.yview_moveto(fy)

            self.lbl_spotlight_badge.configure(
                text=f"📍 Tiêu điểm: {self.active_loc}", text_color="#00F0FF"
            )

        # Update fullscreen modal canvas if currently open
        if getattr(self, "_active_fullscreen_canvas", None) is not None:
            try:
                f_can = self._active_fullscreen_canvas
                f_modal = self._active_fullscreen_modal
                tk_f_img = ImageTk.PhotoImage(img)
                f_can.delete("all")
                f_can.create_image(0, 0, anchor="nw", image=tk_f_img)
                f_can.configure(scrollregion=(0, 0, img.width, img.height))
                f_modal._tk_img_ref = tk_f_img
            except Exception:
                pass

    # ── DISPLAY MODE SWITCHING ───────────────────────────────────────────────
    def _on_display_mode_change(self, mode_str):
        if "Bảng" in mode_str:
            self.display_mode = "table"
            self.drawing_panel.pack_forget()
            self.table_panel.pack(fill="both", expand=True)
        elif "Bản Vẽ" in mode_str:
            self.display_mode = "drawing"
            self.table_panel.pack_forget()
            self.drawing_panel.pack(fill="both", expand=True)
            self._render_drawing()
        else:
            self.display_mode = "split"
            self.table_panel.pack_forget()
            self.drawing_panel.pack_forget()
            self.table_panel.pack(side="left", fill="both", expand=True, padx=(0, 3))
            self.drawing_panel.pack(side="right", fill="both", expand=True, padx=(3, 0))
            self._render_drawing()

    # ── DRAWING CONTROLS: ZOOM, PAN, ROTATE, FULLSCREEN ──────────────────────
    def _zoom_in(self):
        self.zoom_level = min(3.0, round(self.zoom_level + 0.2, 1))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing()

    def _zoom_out(self):
        self.zoom_level = max(0.4, round(self.zoom_level - 0.2, 1))
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")
        self._render_drawing()

    def _zoom_reset(self):
        self.zoom_level = 1.0
        self.lbl_zoom.configure(text="100%")
        self._render_drawing()

    def _rotate_drawing(self):
        self.rotation_angle = (self.rotation_angle + 90) % 360
        self.btn_rotate_dwg.configure(text=f"🔄 {self.rotation_angle}°")
        self._render_drawing()

    def _on_canvas_press(self, event):
        self._pan_start_x = event.x
        self._pan_start_y = event.y
        self._pan_has_dragged = False
        self.canvas.config(cursor="fleur")
        self.canvas.scan_mark(event.x, event.y)

    def _on_canvas_drag(self, event):
        dx = abs(event.x - self._pan_start_x)
        dy = abs(event.y - self._pan_start_y)
        if dx > 4 or dy > 4:
            self._pan_has_dragged = True
            self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _on_canvas_release(self, event):
        self.canvas.config(cursor="")
        if getattr(self, "_pan_has_dragged", False):
            # User was panning the canvas, do not trigger component click
            return

        # Single click: check if clicked on a highlighted component box
        draw_x0 = getattr(self, "draw_offset_x", 0)
        draw_y0 = getattr(self, "draw_offset_y", 0)
        canvas_x = self.canvas.canvasx(event.x) - draw_x0
        canvas_y = self.canvas.canvasy(event.y) - draw_y0

        for loc, (bx0, by0, bx1, by1) in self.pixel_coords_map.items():
            if bx0 - 6 <= canvas_x <= bx1 + 6 and by0 - 6 <= canvas_y <= by1 + 6:
                # Component clicked!
                self.active_loc = loc
                # Toggle QC status
                curr = self.qc_status_map.get(loc, "PENDING")
                self.qc_status_map[loc] = "OK" if curr != "OK" else "PENDING"
                self._populate_table()
                # Focus row in Treeview
                if self.tree.exists(loc):
                    self.tree.selection_set(loc)
                    self.tree.see(loc)
                self._render_drawing()
                break

    def _on_canvas_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    def _open_fullscreen_drawing(self):
        if not self.file_drawing:
            messagebox.showinfo("Chưa có bản vẽ", "Vui lòng chọn file Bản vẽ Working Manual trước!")
            return

        modal = ctk.CTkToplevel(self)
        modal.title(f"VIPQC AI — TOÀN MÀN HÌNH BẢN VẼ PCB ({os.path.basename(self.file_drawing)})")
        modal.geometry("1400x900")
        modal.after(100, modal.lift)

        # Fullscreen toolbar
        f_tb = ctk.CTkFrame(modal, fg_color=BG_SURFACE, height=36)
        f_tb.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(
            f_tb, text="📐 BẢN VẼ PCB TOÀN MÀN HÌNH", font=("Segoe UI", 11, "bold"), text_color=ACCENT_TEAL
        ).pack(side="left", padx=10)

        # Zoom in fullscreen
        ctk.CTkButton(
            f_tb, text="🔍 -", width=36, height=24, command=self._zoom_out
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            f_tb, text="🔍 +", width=36, height=24, command=self._zoom_in
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            f_tb, text="🔄 Xoay", width=70, height=24, command=self._rotate_drawing
        ).pack(side="left", padx=4)

        def _close_fullscreen():
            self._active_fullscreen_canvas = None
            self._active_fullscreen_modal = None
            modal.destroy()

        ctk.CTkButton(
            f_tb, text="✕ Đóng", width=70, height=24, fg_color="#E74C3C", command=_close_fullscreen
        ).pack(side="right", padx=10)

        # Fullscreen Canvas
        f_canvas_frame = ctk.CTkFrame(modal, fg_color="#0A0E18")
        f_canvas_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        f_canvas = tk.Canvas(f_canvas_frame, bg="#0A0E18", highlightthickness=0)
        f_vsb = ttk.Scrollbar(f_canvas_frame, orient="vertical", command=f_canvas.yview)
        f_hsb = ttk.Scrollbar(f_canvas_frame, orient="horizontal", command=f_canvas.xview)
        f_canvas.configure(xscrollcommand=f_hsb.set, yscrollcommand=f_vsb.set)

        f_canvas.pack(side="left", fill="both", expand=True)
        f_vsb.pack(side="right", fill="y")
        f_hsb.pack(side="bottom", fill="x")

        # Bind smooth pan & mousewheel in fullscreen
        modal._f_pan_dragged = False
        def _f_press(e):
            modal._f_start_x = e.x
            modal._f_start_y = e.y
            modal._f_pan_dragged = False
            f_canvas.config(cursor="fleur")
            f_canvas.scan_mark(e.x, e.y)

        def _f_drag(e):
            if abs(e.x - getattr(modal, "_f_start_x", e.x)) > 4 or abs(e.y - getattr(modal, "_f_start_y", e.y)) > 4:
                modal._f_pan_dragged = True
                f_canvas.scan_dragto(e.x, e.y, gain=1)

        def _f_release(e):
            f_canvas.config(cursor="")

        f_canvas.bind("<ButtonPress-1>", _f_press)
        f_canvas.bind("<B1-Motion>", _f_drag)
        f_canvas.bind("<ButtonRelease-1>", _f_release)
        f_canvas.bind("<MouseWheel>", lambda e: self._zoom_in() if e.delta > 0 else self._zoom_out())

        self._active_fullscreen_canvas = f_canvas
        self._active_fullscreen_modal = modal
        modal.protocol("WM_DELETE_WINDOW", _close_fullscreen)

        if self.current_annotated_img:
            tk_img = ImageTk.PhotoImage(self.current_annotated_img)
            f_canvas.create_image(0, 0, anchor="nw", image=tk_img)
            f_canvas.configure(scrollregion=(0, 0, self.current_annotated_img.width, self.current_annotated_img.height))
            modal._tk_img_ref = tk_img

    # ── TABLE POPULATION & INTERACTIVE ACTIONS ───────────────────────────────
    def _populate_table(self):
        self.tree.delete(*self.tree.get_children())
        if not self.comparison_result:
            return

        fil = self.filter_var.get()
        query = self.search_var.get().strip().upper()

        if "Cần chú ý" in fil:
            items = self.comparison_result["qc_focus_items"]
        elif "Thêm" in fil:
            items = self.comparison_result["added_items"]
        elif "Bớt" in fil:
            items = self.comparison_result["removed_items"]
        elif "Đổi" in fil:
            items = self.comparison_result["modified_items"]
        elif "Khớp" in fil:
            items = self.comparison_result["matched_items"]
        else:
            items = self.comparison_result["all_items"]

        for item in items:
            loc = item["location"]
            if query and (query not in loc.upper() and query not in item["part_a"].upper() and query not in item["part_b"].upper()):
                continue

            status = item["status"]
            qc_st = self.qc_status_map.get(loc, "PENDING")
            if qc_st == "OK":
                check_display = "✅ ĐÃ DUYỆT"
                tag = "tag_ok"
            elif qc_st == "NG":
                check_display = "❌ LỖI (NG)"
                tag = "tag_removed"
            else:
                check_display = "⏳ Chờ kiểm"
                tag = f"tag_{status.lower()}"

            self.tree.insert("", "end", iid=loc, values=(
                check_display,
                loc,
                item["status_vn"],
                item["part_a"],
                item["part_b"],
                item["action_guide"],
                item["spec_b"] or item["spec_a"]
            ), tags=(tag,))

        self._update_progress()

    def _on_tree_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        # Set active spotlight on drawing
        self.active_loc = item_id

        # Toggle QC check status
        current = self.qc_status_map.get(item_id, "PENDING")
        if current == "PENDING":
            self.qc_status_map[item_id] = "OK"
        elif current == "OK":
            self.qc_status_map[item_id] = "NG"
        else:
            self.qc_status_map[item_id] = "PENDING"

        self._populate_table()

        # Re-render / Spotlight on Drawing
        self._render_drawing()

    def _on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id or not self.comparison_result:
            return

        matched = [i for i in self.comparison_result["all_items"] if i["location"] == item_id]
        if not matched:
            return
        item = matched[0]

        msg = (
            f"VỊ TRÍ: {item['location']}\n"
            f"Trạng thái: {item['status_vn']}\n\n"
            f"Mã linh kiện Series A: {item['part_a']}\n"
            f"Mã linh kiện Series B: {item['part_b']}\n\n"
            f"HƯỚNG DẪN KIỂM TRA QC:\n{item['action_guide']}\n\n"
            f"Thông số kỹ thuật: {item['spec_b'] or item['spec_a']}"
        )
        messagebox.showinfo(f"Chi Tiết Linh Kiện {item_id}", msg)

    def _update_progress(self):
        if not self.comparison_result:
            self.lbl_qc_progress.configure(text="Tiến độ kiểm tra FAI: 0 / 0 (0%)")
            self.pbar_qc.set(0)
            return

        focus_items = self.comparison_result["qc_focus_items"]
        total = len(focus_items)
        if total == 0:
            self.lbl_qc_progress.configure(text="✅ Không có linh kiện khác biệt cần kiểm!")
            self.pbar_qc.set(1.0)
            return

        focus_locs = [i["location"] for i in focus_items]
        checked = sum(1 for loc in focus_locs if self.qc_status_map.get(loc) == "OK")
        pct = int(checked / total * 100)
        self.lbl_qc_progress.configure(text=f"Tiến độ kiểm tra FAI: {checked} / {total} linh kiện ({pct}%)")
        self.pbar_qc.set(checked / total)

    def _mark_all_ok(self):
        if not self.comparison_result:
            return
        for item in self.comparison_result["qc_focus_items"]:
            self.qc_status_map[item["location"]] = "OK"
        self._populate_table()

    def _reset_qc_checks(self):
        if not self.comparison_result:
            return
        self.qc_status_map = {}
        self._populate_table()

    def _export_excel(self):
        if not self.comparison_result:
            return

        s = self.comparison_result["summary"]
        model = s.get("model_b") or s.get("model_a") or "BOM"
        s_a = s.get("series_a") or "A"
        s_b = s.get("series_b") or "B"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"FAI_Checklist_{model}_{s_a}_vs_{s_b}_{timestamp}.xlsx"

        out_path = filedialog.asksaveasfilename(
            title="Lưu Biên Bản Kiểm Tra FAI & So Sánh BOM",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out_path:
            return

        try:
            for item in self.comparison_result["all_items"]:
                loc = item["location"]
                if loc in self.qc_status_map:
                    item["qc_status"] = self.qc_status_map[loc]

            sbe.export_series_bom_report(self.comparison_result, out_path)
            self.app.set_status(f"Đã xuất biên bản FAI thành công: {os.path.basename(out_path)}")

            resp = messagebox.askyesno(
                "Xuất Thành Công",
                f"Đã tạo file báo cáo FAI thành công tại:\n{out_path}\n\nBạn có muốn mở file ngay không?"
            )
            if resp:
                os.startfile(out_path)
        except Exception as e:
            messagebox.showerror("Lỗi Xuất File", f"Không thể xuất file Excel:\n\n{str(e)}")


class BOMExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("VIPQC AI - Working Manual & PCB Drawing Comparator")
        self.geometry("1360x860")
        self.minsize(1180, 740)
        self.configure(fg_color=BG_DEEP)

        # Set Window & Taskbar Icon (from public/Logo.ico & public/Logo.png)
        try:
            ico_file = resource_path(os.path.join("public", "Logo.ico"))
            if os.path.exists(ico_file):
                self.iconbitmap(ico_file)
        except Exception:
            pass

        try:
            png_file = resource_path(os.path.join("public", "Logo.png"))
            if os.path.exists(png_file):
                self._app_logo_pil = Image.open(png_file)
                self._app_logo_tk = ImageTk.PhotoImage(self._app_logo_pil.resize((32, 32), Image.Resampling.LANCZOS))
                self.wm_iconphoto(True, self._app_logo_tk)
        except Exception:
            pass

        # ── State ─────────────────────────────────────────────────────────────
        self.current_lang = "vi"
        self.current_theme = "dark"
        self.current_nav = "extract"
        self.selected_files: list[str] = []
        self.output_dir = ctk.StringVar(value=os.path.abspath("excel_results"))
        self.export_individual = ctk.BooleanVar(value=True)
        self.export_batch = ctk.BooleanVar(value=True)
        self.all_records: list[dict] = []
        self.last_exports: list[str] = []
        self.is_running = False
        self._extract_hud = None
        self.preview_card = None

        self.LANG_MAP = {
            "Tiếng Việt": "vi",
            "中文": "zh",
            "English": "en"
        }

        self.col_specs = [
            ("seq",       48,  "center"),
            ("item_no",   52,  "center"),
            ("process",   120, "center"),
            ("variant",   110, "center"),
            ("part_no",   160, "w"),
            ("sn",        110, "center"),
            ("rating",    170, "w"),
            ("qty",       50,  "center"),
            ("remarks",   200, "w"),
            ("alternates",130, "w"),
            ("page",      50,  "center"),
        ]

        self._build()
        self._apply_theme("dark")

        # Start RGB border animation
        self._rgb_tick = 0
        self._animate_rgb_badges()

        # Native drag & drop hook
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self, func=self._on_drop_files)
            except Exception:
                pass

    def t(self, key: str, **kwargs) -> str:
        """Translate a key according to current_lang with optional kwargs interpolation."""
        val = I18N.get(self.current_lang, I18N["vi"]).get(key, "")
        if kwargs and isinstance(val, str):
            try:
                return val.format(**kwargs)
            except Exception:
                return val
        return val

    # ──────────────────────────────────────────────────────────────────────────
    # UI BUILD
    # ──────────────────────────────────────────────────────────────────────────
    def _build(self):
        # === TOP HEADER BAR ===================================================
        header = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=86)
        header.pack(fill="x")
        header.pack_propagate(False)

        inner_h = ctk.CTkFrame(header, fg_color="transparent")
        inner_h.pack(fill="both", expand=True, padx=20, pady=(12, 14))

        # Logo / title left
        title_box = ctk.CTkFrame(inner_h, fg_color="transparent")
        title_box.pack(side="left")

        # Top Header VIPQC Logo Image
        try:
            png_file = resource_path(os.path.join("public", "Logo.png"))
            if os.path.exists(png_file):
                logo_pil = Image.open(png_file)
                self.ctk_header_logo = ctk.CTkImage(light_image=logo_pil, dark_image=logo_pil, size=(52, 52))
                self.lbl_header_logo = ctk.CTkLabel(title_box, image=self.ctk_header_logo, text="", fg_color="transparent")
                self.lbl_header_logo.pack(side="left", padx=(0, 14))
        except Exception:
            pass

        title_text_box = ctk.CTkFrame(title_box, fg_color="transparent")
        title_text_box.pack(side="left")

        self.lbl_app_title = ctk.CTkLabel(title_text_box, text=self.t("app_title"),
                                          font=FONT_HERO, text_color=TEXT_PRIMARY,
                                          fg_color="transparent")
        self.lbl_app_title.pack(anchor="w")

        self.lbl_app_subtitle = ctk.CTkLabel(title_text_box,
                                             text=self.t("app_subtitle"),
                                             font=("Segoe UI", 10), text_color=TEXT_MUTED,
                                             fg_color="transparent")
        self.lbl_app_subtitle.pack(anchor="w", pady=(2, 0))

        # Controls & Stat badges right (now with ample breathing room!)
        right_header = ctk.CTkFrame(inner_h, fg_color="transparent")
        right_header.pack(side="right")

        # Language dropdown
        self.lang_var = ctk.StringVar(value="Tiếng Việt")
        self.opt_lang = ctk.CTkOptionMenu(
            right_header,
            values=["Tiếng Việt", "中文", "English"],
            variable=self.lang_var,
            width=110, height=32, corner_radius=8,
            font=("Segoe UI", 10, "bold"),
            dropdown_font=("Segoe UI", 10),
            fg_color=BG_SURFACE,
            button_color=BG_HOVER,
            text_color=TEXT_PRIMARY,
            command=self._on_language_change
        )
        self.opt_lang.pack(side="left", padx=(0, 8))

        # Theme toggle (Segmented button)
        self.theme_var = ctk.StringVar(value=self.t("theme_dark"))
        self.seg_theme = ctk.CTkSegmentedButton(
            right_header,
            values=[self.t("theme_dark"), self.t("theme_light")],
            variable=self.theme_var,
            width=145, height=32, corner_radius=8,
            font=("Segoe UI", 10, "bold"),
            fg_color=BG_SURFACE,
            selected_color=ACCENT_TEAL,
            selected_hover_color=ACCENT_TEAL,
            unselected_color=BG_SURFACE,
            unselected_hover_color=BG_HOVER,
            text_color=TEXT_PRIMARY,
            command=self._on_theme_change
        )
        self.seg_theme.pack(side="left", padx=(0, 12))

        # Stat badges with RGB rotating chasing borders (phased offsets 0, 8, 16)
        self.badge_files = StatBadge(right_header, self.t("badge_files"), offset=0)
        self.badge_files.pack(side="left", padx=4)
        self.badge_items = StatBadge(right_header, self.t("badge_items"), offset=8)
        self.badge_items.pack(side="left", padx=4)
        self.badge_qty = StatBadge(right_header, self.t("badge_qty"), offset=16)
        self.badge_qty.pack(side="left", padx=4)

        # Accent border line
        self.accent_line = ctk.CTkFrame(self, fg_color=ACCENT_TEAL, height=2, corner_radius=0)
        self.accent_line.pack(fill="x")

        # === BODY CONTAINER (Sidebar + Main Workspace) ========================
        self.body_container = ctk.CTkFrame(self, fg_color="transparent")
        self.body_container.pack(fill="both", expand=True, padx=16, pady=10)

        # ── LEFT NAVIGATION SIDEBAR ───────────────────────────────────────────
        self.sidebar = ctk.CTkFrame(self.body_container, fg_color=BG_CARD, corner_radius=12,
                                    border_width=1, border_color=BORDER_CLR, width=220)
        self.sidebar.pack(side="left", fill="y", padx=(0, 12))
        self.sidebar.pack_propagate(False)
        self._build_sidebar(self.sidebar)

        # ── MAIN CONTENT WORKSPACE ────────────────────────────────────────────
        self.main_container = ctk.CTkFrame(self.body_container, fg_color="transparent")
        self.main_container.pack(side="right", fill="both", expand=True)

        # ── TAB 1: EXTRACT VIEW ───────────────────────────────────────────────
        self.view_extract = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.view_extract.pack(fill="both", expand=True)

        # ── LEFT PANEL (QUEUE & SETTINGS) ─────────────────────────────────────
        left = ctk.CTkFrame(self.view_extract, fg_color=BG_CARD, corner_radius=12,
                            border_width=1, border_color=BORDER_CLR, width=310)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)
        self._build_left(left)

        # ── RIGHT PANEL (PREVIEW & LOGS) ──────────────────────────────────────
        right = ctk.CTkFrame(self.view_extract, fg_color="transparent")
        right.pack(side="right", fill="both", expand=True)
        self._build_right(right)

        # ── TAB 2: COMPARE VIEW ───────────────────────────────────────────────
        self.view_compare = BOMCompareView(self.main_container, self)
        # (Initially hidden, shown when nav changes to Compare)

        # ── TAB 3: MODEL COMPARE VIEW ─────────────────────────────────────────
        self.view_model_compare = ModelCompareView(self.main_container, self)
        # (Initially hidden, shown when nav changes to Model Compare)

        # ── TAB 4: SERIES BOM COMPARE VIEW ────────────────────────────────────
        self.view_series_bom_compare = SeriesBOMCompareView(self.main_container, self)
        # (Initially hidden, shown when nav changes to Series BOM Compare)

        # === BOTTOM STATUS BAR ================================================
        status_bar = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=34)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        self.lbl_copyright_footer = ctk.CTkLabel(
            status_bar, text=self.t("copyright_footer"),
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_TEAL
        )
        self.lbl_copyright_footer.pack(side="right", padx=16, pady=6)

        self.lbl_status = ctk.CTkLabel(
            status_bar, text=self.t("status_ready"),
            font=FONT_BODY, text_color=TEXT_MUTED, anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True, padx=16, pady=6)

        # Silent background update check after 3.5 seconds
        self.after(3500, self._auto_check_update)

    # ── SIDEBAR INTERNALS ────────────────────────────────────────────────────
    def _build_sidebar(self, parent):
        # Section Label
        self.lbl_sidebar_menu = ctk.CTkLabel(
            parent, text=self.t("sidebar_menu"),
            font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED, anchor="w"
        )
        self.lbl_sidebar_menu.pack(fill="x", padx=14, pady=(14, 8))

        # Navigation Buttons
        self.btn_nav_extract = ctk.CTkButton(
            parent,
            text=self.t("tab_extract"),
            font=("Segoe UI", 11, "bold"),
            height=44,
            corner_radius=8,
            anchor="w",
            fg_color=ACCENT_TEAL,
            text_color=("#FFFFFF", "#0B0F1A"),
            hover_color=("#0F766E", "#00A88C"),
            command=lambda: self._set_nav_active("extract")
        )
        self.btn_nav_extract.pack(fill="x", padx=10, pady=(0, 6))

        self.btn_nav_compare = ctk.CTkButton(
            parent,
            text=self.t("tab_compare"),
            font=("Segoe UI", 11, "bold"),
            height=44,
            corner_radius=8,
            anchor="w",
            fg_color="transparent",
            text_color=TEXT_PRIMARY,
            hover_color=BG_HOVER,
            command=lambda: self._set_nav_active("compare")
        )
        self.btn_nav_compare.pack(fill="x", padx=10, pady=(0, 6))

        self.btn_nav_model_comp = ctk.CTkButton(
            parent,
            text=self.t("tab_model_comp"),
            font=("Segoe UI", 11, "bold"),
            height=44,
            corner_radius=8,
            anchor="w",
            fg_color="transparent",
            text_color=TEXT_PRIMARY,
            hover_color=BG_HOVER,
            command=lambda: self._set_nav_active("model_comp")
        )
        self.btn_nav_model_comp.pack(fill="x", padx=10, pady=(0, 6))

        self.btn_nav_series_bom = ctk.CTkButton(
            parent,
            text=self.t("tab_series_bom"),
            font=("Segoe UI", 11, "bold"),
            height=44,
            corner_radius=8,
            anchor="w",
            fg_color="transparent",
            text_color=TEXT_PRIMARY,
            hover_color=BG_HOVER,
            command=lambda: self._set_nav_active("series_bom")
        )
        self.btn_nav_series_bom.pack(fill="x", padx=10, pady=(0, 10))

        # Subtle divider
        ctk.CTkFrame(parent, fg_color=BORDER_CLR, height=1).pack(fill="x", padx=10, pady=4)

        # Bottom Copyright Card (CĂN GIỮA, NỔI BẬT, VIỀN TEAL ACCENT)
        self.card_copyright = ctk.CTkFrame(
            parent, fg_color=BG_SURFACE, corner_radius=12,
            border_width=1.5, border_color=ACCENT_TEAL
        )
        self.card_copyright.pack(side="bottom", fill="x", padx=10, pady=(0, 14))

        # Sidebar VIPQC Logo Image
        try:
            png_file = resource_path(os.path.join("public", "Logo.png"))
            if os.path.exists(png_file):
                side_logo_pil = Image.open(png_file)
                self.ctk_sidebar_logo = ctk.CTkImage(light_image=side_logo_pil, dark_image=side_logo_pil, size=(46, 46))
                self.lbl_sidebar_logo = ctk.CTkLabel(self.card_copyright, image=self.ctk_sidebar_logo, text="", fg_color="transparent")
                self.lbl_sidebar_logo.pack(pady=(10, 2))
        except Exception:
            pass

        self.lbl_sidebar_c_badge = ctk.CTkLabel(
            self.card_copyright,
            text="🛡️  " + self.t("copyright_title") + "  🛡️",
            font=("Segoe UI", 9, "bold"),
            text_color=ACCENT_TEAL,
            anchor="center",
            justify="center"
        )
        self.lbl_sidebar_c_badge.pack(fill="x", padx=8, pady=(2, 4))

        self.lbl_sidebar_author = ctk.CTkLabel(
            self.card_copyright,
            text=self.t("copyright_text"),
            font=("Segoe UI", 11, "bold"),
            text_color=TEXT_PRIMARY,
            justify="center",
            anchor="center"
        )
        self.lbl_sidebar_author.pack(fill="x", padx=8, pady=(0, 4))

        # Interactive Version Badge & Update Checker Button
        self.btn_sidebar_update = ctk.CTkButton(
            self.card_copyright,
            text=f"🔄 v{updater.CURRENT_VERSION} • Cập nhật",
            font=("Segoe UI", 9, "bold"),
            fg_color=BG_CARD,
            hover_color=BG_HOVER,
            text_color=ACCENT_AMBER,
            corner_radius=6,
            height=26,
            command=self._manual_check_update
        )
        self.btn_sidebar_update.pack(fill="x", padx=10, pady=(0, 10))

    # ── LEFT PANEL INTERNALS ─────────────────────────────────────────────────
    def _build_left(self, parent):
        pad = {"padx": 16}

        # Section label
        self.lbl_sec_queue = ctk.CTkLabel(parent, text=self.t("sec_queue"),
                                          font=("Segoe UI", 11, "bold"),
                                          text_color=ACCENT_TEAL, fg_color="transparent")
        self.lbl_sec_queue.pack(anchor="w", pady=(14, 4), **pad)

        # Add buttons row
        btn_row = ctk.CTkFrame(parent, fg_color="transparent")
        btn_row.pack(fill="x", **pad)

        self.btn_add_f = ctk.CTkButton(btn_row, text=self.t("btn_add_files"),
                                       fg_color=ACCENT_BLUE, hover_color="#2B70D6",
                                       text_color="#FFFFFF",
                                       font=FONT_H2, corner_radius=8, height=36,
                                       command=self._add_files)
        self.btn_add_f.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_add_d = ctk.CTkButton(btn_row, text=self.t("btn_add_folder"),
                                       fg_color=BG_SURFACE, hover_color=BG_HOVER,
                                       text_color=TEXT_PRIMARY,
                                       font=FONT_H2, corner_radius=8, height=36,
                                       command=self._add_folder)
        self.btn_add_d.pack(side="right", fill="x", expand=True, padx=(4, 0))

        # Scrollable file list
        self.file_scroll = ctk.CTkScrollableFrame(
            parent, fg_color=BG_SURFACE, corner_radius=8,
            border_width=1, border_color=BORDER_CLR)
        self.file_scroll.pack(fill="both", expand=True, **pad, pady=8)

        self.lbl_empty_list = ctk.CTkLabel(
            self.file_scroll,
            text=self.t("empty_files"),
            font=FONT_BODY, text_color=TEXT_MUTED, fg_color="transparent")
        self.lbl_empty_list.pack(pady=30)

        # Clear button
        self.btn_clear = ctk.CTkButton(parent, text=self.t("btn_clear_all"),
                                       fg_color="transparent", text_color="#E74C3C",
                                       hover_color=("#FEE2E2", "#2D1010"),
                                       font=("Segoe UI", 10), height=26,
                                       command=self._clear_files)
        self.btn_clear.pack(anchor="e", **pad)

        # Divider
        ctk.CTkFrame(parent, fg_color=BORDER_CLR, height=1).pack(
            fill="x", **pad, pady=8)

        # OUTPUT SETTINGS
        self.lbl_sec_output = ctk.CTkLabel(parent, text=self.t("sec_output"),
                                           font=("Segoe UI", 11, "bold"),
                                           text_color=ACCENT_TEAL, fg_color="transparent")
        self.lbl_sec_output.pack(anchor="w", pady=(0, 6), **pad)

        out_row = ctk.CTkFrame(parent, fg_color="transparent")
        out_row.pack(fill="x", **pad, pady=(0, 6))
        out_entry = ctk.CTkEntry(out_row, textvariable=self.output_dir,
                                 font=FONT_BODY, height=34,
                                 fg_color=BG_SURFACE, border_color=BORDER_CLR)
        out_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkButton(out_row, text="…", width=36, height=34,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY,
                      hover_color=BG_HOVER,
                      command=self._choose_out).pack(side="right")

        self.sw_individual = ctk.CTkSwitch(parent, text=self.t("switch_individual"),
                                           variable=self.export_individual,
                                           font=FONT_BODY, text_color=TEXT_PRIMARY,
                                           progress_color=ACCENT_TEAL, button_color=TEXT_PRIMARY,
                                           fg_color=BG_SURFACE)
        self.sw_individual.pack(anchor="w", **pad, pady=3)

        self.sw_batch = ctk.CTkSwitch(parent, text=self.t("switch_batch"),
                                      variable=self.export_batch,
                                      font=FONT_BODY, text_color=TEXT_PRIMARY,
                                      progress_color=ACCENT_TEAL, button_color=TEXT_PRIMARY,
                                      fg_color=BG_SURFACE)
        self.sw_batch.pack(anchor="w", **pad, pady=3)

        # Divider
        ctk.CTkFrame(parent, fg_color=BORDER_CLR, height=1).pack(
            fill="x", **pad, pady=8)

        # RUN BUTTON
        self.btn_run = ctk.CTkButton(
            parent,
            text=self.t("btn_run"),
            fg_color=ACCENT_TEAL, hover_color=("#0F766E", "#00A88C"),
            text_color=("#FFFFFF", "#0B0F1A"),
            font=("Segoe UI", 13, "bold"), corner_radius=10, height=48,
            command=self._start_thread)
        self.btn_run.pack(fill="x", **pad, pady=(0, 8))

        # Quick access row
        qa = ctk.CTkFrame(parent, fg_color="transparent")
        qa.pack(fill="x", **pad, pady=(0, 14))

        self.btn_open_xl = ctk.CTkButton(
            qa, text=self.t("btn_open_excel"),
            fg_color=BG_SURFACE, text_color=ACCENT_TEAL,
            hover_color=BG_HOVER, font=FONT_H2, height=36,
            state="disabled", command=self._open_excel)
        self.btn_open_xl.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_open_dir = ctk.CTkButton(
            qa, text=self.t("btn_open_folder"),
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY,
            hover_color=BG_HOVER, font=FONT_H2, height=36,
            command=self._open_folder)
        self.btn_open_dir.pack(side="right", fill="x", expand=True, padx=(4, 0))

    # ── RIGHT PANEL INTERNALS ────────────────────────────────────────────────
    def _build_right(self, parent):
        # Preview header
        preview_hdr = ctk.CTkFrame(parent, fg_color="transparent")
        preview_hdr.pack(fill="x", pady=(0, 8))

        self.lbl_sec_preview = ctk.CTkLabel(preview_hdr, text=self.t("sec_preview"),
                                            font=("Segoe UI", 12, "bold"),
                                            text_color=ACCENT_TEAL, fg_color="transparent")
        self.lbl_sec_preview.pack(side="left")

        self.lbl_preview_count = ctk.CTkLabel(preview_hdr, text="",
                                              font=("Segoe UI", 10),
                                              text_color=TEXT_MUTED, fg_color="transparent")
        self.lbl_preview_count.pack(side="left", padx=(10, 0))

        # Search bar
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        self.search_entry = ctk.CTkEntry(preview_hdr, textvariable=self.search_var,
                                         placeholder_text=self.t("search_placeholder"),
                                         font=FONT_BODY, width=360, height=34,
                                         fg_color=BG_SURFACE, border_color=BORDER_CLR)
        self.search_entry.pack(side="right")

        # Process code filter tabs
        self.filter_frame = ctk.CTkFrame(parent, fg_color=BG_CARD,
                                         corner_radius=8, height=38)
        self.filter_frame.pack(fill="x", pady=(0, 8))
        self.filter_frame.pack_propagate(False)
        self._filter_all_btn = None
        self._active_filter = None

        # ── Treeview table ───────────────────────────────────────────────────
        self.preview_card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10,
                                         border_width=1, border_color=BORDER_CLR)
        self.preview_card.pack(fill="both", expand=True)
        preview_card = self.preview_card

        self.tree_frame = tk.Frame(preview_card, bg="#131929")
        self.tree_frame.pack(fill="both", expand=True, padx=2, pady=2)

        self.vsb = tk.Scrollbar(self.tree_frame, orient="vertical")
        self.vsb.pack(side="right", fill="y")
        self.hsb = tk.Scrollbar(self.tree_frame, orient="horizontal")
        self.hsb.pack(side="bottom", fill="x")

        cols = [s[0] for s in self.col_specs]
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show="headings",
                                 style="BOM.Treeview",
                                 yscrollcommand=self.vsb.set,
                                 xscrollcommand=self.hsb.set)
        self.tree.pack(fill="both", expand=True)
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # Configure columns
        for cid, width, anchor in self.col_specs:
            hdr_text = I18N[self.current_lang]["cols"].get(cid, cid)
            self.tree.heading(cid, text=hdr_text, anchor=anchor)
            self.tree.column(cid, width=width, minwidth=40, anchor=anchor)

        # Double-click to copy Part No & Right-click context menu
        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._on_tree_right_click)

        # Empty State Hero Overlay
        self.empty_preview = ctk.CTkFrame(preview_card, fg_color=BG_CARD, corner_radius=14)
        self.empty_preview.place(relx=0.5, rely=0.45, anchor="center")

        self.lbl_empty_icon = ctk.CTkLabel(self.empty_preview, text="📑",
                                           font=("Segoe UI", 48), fg_color="transparent")
        self.lbl_empty_icon.pack(pady=(20, 8))

        self.lbl_empty_title = ctk.CTkLabel(self.empty_preview, text=self.t("empty_preview_title"),
                                            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY,
                                            fg_color="transparent")
        self.lbl_empty_title.pack(padx=28, pady=(0, 6))

        self.lbl_empty_desc = ctk.CTkLabel(self.empty_preview, text=self.t("empty_preview_desc"),
                                           font=("Segoe UI", 11), text_color=TEXT_MUTED,
                                           fg_color="transparent")
        self.lbl_empty_desc.pack(padx=28, pady=(0, 20))

        # Floating Toast popup
        self.toast_frame = ctk.CTkFrame(self, fg_color=ACCENT_TEAL, corner_radius=16)
        self.toast_lbl = ctk.CTkLabel(self.toast_frame, text="",
                                      font=("Segoe UI", 10, "bold"), text_color=("#FFFFFF", "#0B0F1A"))
        self.toast_lbl.pack(padx=16, pady=6)
        self._toast_timer = None

        # ── Log / Progress section ────────────────────────────────────────────
        self.is_log_collapsed = False
        self.log_card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10,
                                     border_width=1, border_color=BORDER_CLR, height=155)
        self.log_card.pack(fill="x", pady=(10, 0))
        self.log_card.pack_propagate(False)

        log_top = ctk.CTkFrame(self.log_card, fg_color="transparent")
        log_top.pack(fill="x", padx=12, pady=(8, 4))

        self.lbl_sec_log = ctk.CTkLabel(log_top, text=self.t("sec_log"),
                                        font=("Segoe UI", 11, "bold"),
                                        text_color=ACCENT_TEAL, fg_color="transparent")
        self.lbl_sec_log.pack(side="left")

        # Collapse toggle button
        self.btn_toggle_log = ctk.CTkButton(
            log_top,
            text=self.t("log_collapse"),
            width=76, height=24, corner_radius=6,
            font=("Segoe UI", 9, "bold"),
            fg_color=BG_SURFACE,
            text_color=TEXT_MUTED,
            hover_color=BG_HOVER,
            command=self._toggle_log
        )
        self.btn_toggle_log.pack(side="right", padx=(8, 0))

        self.progress = AnimatedProgressBar(log_top)
        self.progress.pack(side="right", fill="x", expand=True, padx=(16, 0))

        self.log_box = tk.Text(self.log_card, font=FONT_MONO,
                               relief="flat", borderwidth=0, wrap="word",
                               height=6)
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self._log(self.t("log_ready"), "info")

    # ──────────────────────────────────────────────────────────────────────────
    # THEME & LANGUAGE MANAGEMENT
    # ──────────────────────────────────────────────────────────────────────────
    def _apply_theme(self, mode: str):
        """Applies theme colors to Tkinter and ttk elements that don't auto-update."""
        is_dark = (mode.lower() == "dark")

        tree_bg = "#131929" if is_dark else "#FFFFFF"
        tree_fg = "#E8EDF5" if is_dark else "#0F172A"
        hdr_bg = "#0B0F1A" if is_dark else "#E2E8F0"
        hdr_fg = "#00C9A7" if is_dark else "#0D9488"
        sel_bg = "#232D45" if is_dark else "#BAE6FD"
        sel_fg = "#00C9A7" if is_dark else "#0369A1"
        odd_bg = "#1C2438" if is_dark else "#F8FAFC"
        even_bg = "#131929" if is_dark else "#FFFFFF"
        sb_bg = "#1C2438" if is_dark else "#E2E8F0"

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("BOM.Treeview",
                        background=tree_bg,
                        fieldbackground=tree_bg,
                        foreground=tree_fg,
                        rowheight=28,
                        font=("Segoe UI", 10),
                        borderwidth=0,
                        relief="flat")
        style.configure("BOM.Treeview.Heading",
                        background=hdr_bg,
                        foreground=hdr_fg,
                        font=("Segoe UI", 10, "bold"),
                        relief="flat",
                        padding=5)
        style.map("BOM.Treeview",
                  background=[("selected", sel_bg)],
                  foreground=[("selected", sel_fg)])
        style.map("BOM.Treeview.Heading",
                  background=[("active", odd_bg)])

        self.tree_frame.configure(bg=tree_bg)
        self.tree.tag_configure("odd",  background=odd_bg)
        self.tree.tag_configure("even", background=even_bg)
        self.vsb.configure(bg=sb_bg)
        self.hsb.configure(bg=sb_bg)

        # Log box theme
        log_bg = "#0A0E18" if is_dark else "#F8FAFC"
        log_fg = "#C8D6E5" if is_dark else "#1E293B"
        log_insert = "#00C9A7" if is_dark else "#0D9488"
        log_sel = "#232D45" if is_dark else "#E2E8F0"

        self.log_box.configure(
            bg=log_bg, fg=log_fg,
            insertbackground=log_insert,
            selectbackground=log_sel
        )
        self.log_box.tag_configure("info",    foreground="#5DADE2" if is_dark else "#0284C7")
        self.log_box.tag_configure("success", foreground="#00C9A7" if is_dark else "#0D9488")
        self.log_box.tag_configure("warning", foreground="#FFB830" if is_dark else "#D97706")
        self.log_box.tag_configure("error",   foreground="#E74C3C" if is_dark else "#DC2626")
        self.log_box.tag_configure("bold",    font=("Consolas", 9, "bold"),
                                   foreground="#E8EDF5" if is_dark else "#0F172A")
        self.log_box.tag_configure("muted",   foreground="#7A8BA6" if is_dark else "#64748B")

        # Context menu theme
        menu_bg = "#1C2438" if is_dark else "#FFFFFF"
        menu_fg = "#E8EDF5" if is_dark else "#0F172A"
        menu_act_bg = "#00C9A7" if is_dark else "#0D9488"
        menu_act_fg = "#0B0F1A" if is_dark else "#FFFFFF"
        if hasattr(self, "ctx_menu"):
            try:
                self.ctx_menu.configure(
                    bg=menu_bg, fg=menu_fg,
                    activebackground=menu_act_bg, activeforeground=menu_act_fg,
                    relief="flat", borderwidth=1
                )
            except Exception:
                pass

        # Stat badges theme colors
        if hasattr(self, "badge_files"):
            self.badge_files.set_theme(is_dark)
        if hasattr(self, "badge_items"):
            self.badge_items.set_theme(is_dark)
        if hasattr(self, "badge_qty"):
            self.badge_qty.set_theme(is_dark)

        # Compare view theme
        if hasattr(self, "view_compare"):
            self.view_compare.apply_theme(is_dark)
        if hasattr(self, "view_model_compare"):
            self.view_model_compare.apply_theme(is_dark)
        if hasattr(self, "view_series_bom_compare"):
            self.view_series_bom_compare._apply_tree_tags()

        # Sidebar & Copyright theme
        if hasattr(self, "sidebar"):
            self.sidebar.configure(
                fg_color=BG_CARD[1] if is_dark else BG_CARD[0],
                border_color=BORDER_CLR[1] if is_dark else BORDER_CLR[0]
            )
        if hasattr(self, "card_copyright"):
            self.card_copyright.configure(
                fg_color=BG_SURFACE[1] if is_dark else BG_SURFACE[0],
                border_color=ACCENT_TEAL[1] if is_dark else ACCENT_TEAL[0]
            )
        if hasattr(self, "lbl_sidebar_author"):
            self.lbl_sidebar_author.configure(
                text_color=TEXT_PRIMARY[1] if is_dark else TEXT_PRIMARY[0]
            )
        if hasattr(self, "lbl_copyright_footer"):
            self.lbl_copyright_footer.configure(
                text_color=ACCENT_TEAL[1] if is_dark else ACCENT_TEAL[0]
            )
        if hasattr(self, "btn_nav_extract") and hasattr(self, "btn_nav_compare"):
            self._set_nav_active(self.current_nav)

    def _animate_rgb_badges(self):
        """Continuously rotates the RGB rainbow border around the badges at ~25 FPS."""
        try:
            if not self.winfo_exists():
                return
            self._rgb_tick += 1
            if hasattr(self, "badge_files"):
                self.badge_files.step_animation(self._rgb_tick)
            if hasattr(self, "badge_items"):
                self.badge_items.step_animation(self._rgb_tick)
            if hasattr(self, "badge_qty"):
                self.badge_qty.step_animation(self._rgb_tick)
        except Exception:
            pass
        self.after(40, self._animate_rgb_badges)

    def _on_theme_change(self, choice: str):
        """Callback when user toggles Light/Dark mode."""
        is_light = any(choice.endswith(w) or "Sáng" in choice or "浅色" in choice or "Light" in choice for w in ["Sáng", "浅色", "Light"])
        mode = "light" if is_light else "dark"
        self.current_theme = mode
        ctk.set_appearance_mode(mode)
        self._apply_theme(mode)

    def _on_language_change(self, choice: str):
        """Callback when user selects a language from the dropdown."""
        lang_code = self.LANG_MAP.get(choice, "vi")
        self.current_lang = lang_code
        self._update_ui_language()

    def _set_nav_active(self, mode: str):
        """Switches active view and styles the sidebar buttons accordingly."""
        self.current_nav = mode
        is_dark = (self.current_theme == "dark")

        active_fg = ACCENT_TEAL
        active_txt = ("#FFFFFF", "#0B0F1A")
        inactive_fg = "transparent"
        inactive_txt = TEXT_PRIMARY[1] if is_dark else TEXT_PRIMARY[0]
        hover_col = BG_HOVER[1] if is_dark else BG_HOVER[0]

        if mode == "extract":
            self.btn_nav_extract.configure(fg_color=active_fg, text_color=active_txt, hover_color=("#0F766E", "#00A88C"))
            self.btn_nav_compare.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_model_comp"):
                self.btn_nav_model_comp.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_series_bom"):
                self.btn_nav_series_bom.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "view_model_compare"):
                self.view_model_compare.pack_forget()
            if hasattr(self, "view_series_bom_compare"):
                self.view_series_bom_compare.pack_forget()
            self.view_compare.pack_forget()
            self.view_extract.pack(fill="both", expand=True)
            self._update_badges_for_extract()
            self.set_status("Đã chuyển sang chế độ Trích Xuất BOM.")
        elif mode == "compare":
            self.btn_nav_compare.configure(fg_color=active_fg, text_color=active_txt, hover_color=("#0F766E", "#00A88C"))
            self.btn_nav_extract.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_model_comp"):
                self.btn_nav_model_comp.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_series_bom"):
                self.btn_nav_series_bom.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "view_model_compare"):
                self.view_model_compare.pack_forget()
            if hasattr(self, "view_series_bom_compare"):
                self.view_series_bom_compare.pack_forget()
            self.view_extract.pack_forget()
            self.view_compare.pack(fill="both", expand=True)
            self._update_badges_for_compare()
            self.set_status("Đã chuyển sang chế độ Đối Chiếu BOM.")
        elif mode == "model_comp":
            self.btn_nav_extract.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            self.btn_nav_compare.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_model_comp"):
                self.btn_nav_model_comp.configure(fg_color=active_fg, text_color=active_txt, hover_color=("#0F766E", "#00A88C"))
            if hasattr(self, "btn_nav_series_bom"):
                self.btn_nav_series_bom.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            self.view_extract.pack_forget()
            self.view_compare.pack_forget()
            if hasattr(self, "view_series_bom_compare"):
                self.view_series_bom_compare.pack_forget()
            if hasattr(self, "view_model_compare"):
                self.view_model_compare.pack(fill="both", expand=True)
            self._update_badges_for_model_compare()
            self.set_status("Đã chuyển sang chế độ So Sánh 2 Model & Bản Vẽ PCB.")
        elif mode == "series_bom":
            self.btn_nav_extract.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            self.btn_nav_compare.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_model_comp"):
                self.btn_nav_model_comp.configure(fg_color=inactive_fg, text_color=inactive_txt, hover_color=hover_col)
            if hasattr(self, "btn_nav_series_bom"):
                self.btn_nav_series_bom.configure(fg_color=active_fg, text_color=active_txt, hover_color=("#0F766E", "#00A88C"))
            self.view_extract.pack_forget()
            self.view_compare.pack_forget()
            if hasattr(self, "view_model_compare"):
                self.view_model_compare.pack_forget()
            if hasattr(self, "view_series_bom_compare"):
                self.view_series_bom_compare.pack(fill="both", expand=True)
            self.set_status("Đã chuyển sang chế độ So Sánh 2 BOM (Series).")

    def _on_nav_change(self, choice: str):
        """Compatibility method for switching navigation."""
        is_compare = "compare" in str(choice).lower() or "so sánh" in str(choice).lower() or "比对" in str(choice)
        self._set_nav_active("compare" if is_compare else "extract")

    def _update_badges_for_compare(self):
        """Updates RGB stat badges to reflect BOM Comparison statistics."""
        self.badge_files.set_label(self.t("badge_excel_items"))
        self.badge_items.set_label(self.t("badge_matched_items"))
        self.badge_qty.set_label(self.t("badge_mismatch_items"))

        if hasattr(self, "view_compare") and self.view_compare.comparison_result:
            res = self.view_compare.comparison_result
            self.badge_files.set_value(res.get("total_excel", 0))
            self.badge_items.set_value(res.get("count_matched", 0))
            self.badge_qty.set_value(res.get("total_discrepancies", 0))
        elif hasattr(self, "view_compare") and self.view_compare.excel_data:
            self.badge_files.set_value(self.view_compare.excel_data.get("total_items", 0))
            self.badge_items.set_value(0)
            self.badge_qty.set_value(0)
        else:
            self.badge_files.set_value(0)
            self.badge_items.set_value(0)
            self.badge_qty.set_value(0)

    def _update_badges_for_model_compare(self):
        """Updates RGB stat badges to reflect Model Series Comparison statistics."""
        self.badge_files.set_label(self.t("badge_model_total"))
        self.badge_items.set_label(self.t("badge_model_match"))
        self.badge_qty.set_label(self.t("badge_model_diff"))

        if hasattr(self, "view_model_compare") and self.view_model_compare.comp_result:
            summary = self.view_model_compare.comp_result.get("summary", {})
            self.badge_files.set_value(summary.get("total_positions", 0))
            self.badge_items.set_value(summary.get("match_count", 0))
            self.badge_qty.set_value(summary.get("diff_count", 0))
        else:
            self.badge_files.set_value(0)
            self.badge_items.set_value(0)
            self.badge_qty.set_value(0)

    def _update_badges_for_extract(self):
        """Updates RGB stat badges to reflect Extraction statistics."""
        self.badge_files.set_label(self.t("badge_files"))
        self.badge_items.set_label(self.t("badge_items"))
        self.badge_qty.set_label(self.t("badge_qty"))
        self._refresh_badges()

    def set_status(self, text: str):
        """Updates the text on the bottom status bar."""
        if hasattr(self, "lbl_status"):
            self.lbl_status.configure(text=text)

    def show_toast(self, message: str):
        """Public alias for displaying toast message."""
        self._show_toast(message)

    def _update_ui_language(self):
        """Refreshes all visible UI texts and column headers in the current language."""
        self.lbl_app_title.configure(text=self.t("app_title"))
        self.lbl_app_subtitle.configure(text=self.t("app_subtitle"))

        # Theme toggle options in current language
        dark_txt = self.t("theme_dark")
        light_txt = self.t("theme_light")
        self.seg_theme.configure(values=[dark_txt, light_txt])
        self.theme_var.set(light_txt if self.current_theme == "light" else dark_txt)

        # Sidebar navigation buttons & labels
        if hasattr(self, "lbl_sidebar_menu"):
            self.lbl_sidebar_menu.configure(text=self.t("sidebar_menu"))
        if hasattr(self, "btn_nav_extract"):
            self.btn_nav_extract.configure(text=self.t("tab_extract"))
        if hasattr(self, "btn_nav_compare"):
            self.btn_nav_compare.configure(text=self.t("tab_compare"))
        if hasattr(self, "btn_nav_model_comp"):
            self.btn_nav_model_comp.configure(text=self.t("tab_model_comp"))
        if hasattr(self, "btn_nav_series_bom"):
            self.btn_nav_series_bom.configure(text=self.t("tab_series_bom"))
        if hasattr(self, "lbl_sidebar_c_badge"):
            self.lbl_sidebar_c_badge.configure(text="🛡️  " + self.t("copyright_title") + "  🛡️")
        if hasattr(self, "lbl_sidebar_author"):
            self.lbl_sidebar_author.configure(text=self.t("copyright_text"))
        if hasattr(self, "lbl_copyright_footer"):
            self.lbl_copyright_footer.configure(text=self.t("copyright_footer"))

        # Update badges depending on active view
        if self.current_nav == "compare":
            self._update_badges_for_compare()
        elif self.current_nav == "model_comp":
            self._update_badges_for_model_compare()
        else:
            self._update_badges_for_extract()

        # Update compare view texts
        if hasattr(self, "view_compare"):
            self.view_compare.update_language()
        if hasattr(self, "view_model_compare"):
            self.view_model_compare.update_language()

        # Left panel texts
        self.lbl_sec_queue.configure(text=self.t("sec_queue"))
        self.btn_add_f.configure(text=self.t("btn_add_files"))
        self.btn_add_d.configure(text=self.t("btn_add_folder"))
        self.lbl_empty_list.configure(text=self.t("empty_files"))
        self.btn_clear.configure(text=self.t("btn_clear_all"))
        self.lbl_sec_output.configure(text=self.t("sec_output"))
        self.sw_individual.configure(text=self.t("switch_individual"))
        self.sw_batch.configure(text=self.t("switch_batch"))
        self.btn_run.configure(text=self.t("btn_running") if self.is_running else self.t("btn_run"))
        self.btn_open_xl.configure(text=self.t("btn_open_excel"))
        self.btn_open_dir.configure(text=self.t("btn_open_folder"))

        # Right panel texts
        self.lbl_sec_preview.configure(text=self.t("sec_preview"))
        self.search_entry.configure(placeholder_text=self.t("search_placeholder"))
        self.lbl_sec_log.configure(text=self.t("sec_log"))
        if not self.is_running:
            self.lbl_status.configure(text=self.t("status_ready"))

        if hasattr(self, "lbl_empty_title"):
            self.lbl_empty_title.configure(text=self.t("empty_preview_title"))
            self.lbl_empty_desc.configure(text=self.t("empty_preview_desc"))
        if hasattr(self, "btn_toggle_log"):
            self.btn_toggle_log.configure(text=self.t("log_expand") if self.is_log_collapsed else self.t("log_collapse"))
        if hasattr(self, "lbl_preview_count") and self.all_records:
            cur_rows = len(self.tree.get_children())
            self.lbl_preview_count.configure(text=self.t("preview_count", filtered=cur_rows, total=len(self.all_records)))

        # Table headings
        for cid, _, anchor in self.col_specs:
            hdr_text = I18N[self.current_lang]["cols"].get(cid, cid)
            self.tree.heading(cid, text=hdr_text, anchor=anchor)

        # Filter tabs
        self._rebuild_filter_tabs()

    # ──────────────────────────────────────────────────────────────────────────
    # FILE MANAGEMENT
    # ──────────────────────────────────────────────────────────────────────────
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title=self.t("dlg_select_pdf"),
            filetypes=[("PDF Files", "*.pdf")])
        oversized = []
        for p in paths:
            p = os.path.normpath(p)
            size = os.path.getsize(p) if os.path.exists(p) else 0
            if size > MAX_FILE_SIZE_BYTES:
                oversized.append((os.path.basename(p), format_file_size(size)))
                continue
            if p not in self.selected_files:
                self.selected_files.append(p)
                self._add_file_row(p)
        if oversized:
            names = "\n• " + "\n• ".join(f"{name} ({sz})" for name, sz in oversized)
            messagebox.showwarning(
                self.t("title_file_too_large"),
                self.t("err_batch_too_large", files=names, max=f"{MAX_FILE_SIZE_MB} MB")
            )
            self.show_toast(f"⚠️ Đã bỏ qua {len(oversized)} file vượt quá {MAX_FILE_SIZE_MB}MB!")
        self._refresh_badges()

    def _add_folder(self):
        folder = filedialog.askdirectory(title=self.t("dlg_select_folder"))
        if not folder:
            return
        count = 0
        oversized = []
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".pdf"):
                    p = os.path.normpath(os.path.join(root, f))
                    size = os.path.getsize(p) if os.path.exists(p) else 0
                    if size > MAX_FILE_SIZE_BYTES:
                        oversized.append((f, format_file_size(size)))
                        continue
                    if p not in self.selected_files:
                        self.selected_files.append(p)
                        self._add_file_row(p)
                        count += 1
        self._refresh_badges()
        if oversized:
            names = "\n• " + "\n• ".join(f"{name} ({sz})" for name, sz in oversized[:10])
            if len(oversized) > 10:
                names += f"\n... và {len(oversized) - 10} file khác"
            messagebox.showwarning(
                self.t("title_file_too_large"),
                self.t("err_batch_too_large", files=names, max=f"{MAX_FILE_SIZE_MB} MB")
            )
            self.show_toast(f"⚠️ Đã bỏ qua {len(oversized)} file > {MAX_FILE_SIZE_MB}MB!")
        if count:
            self._log(self.t("log_added_files", count=count, folder=os.path.basename(folder)), "info")

    def _add_file_row(self, path):
        if self.lbl_empty_list.winfo_ismapped():
            self.lbl_empty_list.pack_forget()
        row = FileRow(self.file_scroll, path, on_remove=self._remove_file)
        row.pack(fill="x", pady=2)

    def _remove_file(self, path):
        if path in self.selected_files:
            self.selected_files.remove(path)
        self._refresh_badges()
        if not self.selected_files:
            self.lbl_empty_list.pack(pady=30)

    def _clear_files(self):
        if getattr(self, "is_running", False):
            self.is_running = False
        if getattr(self, "_extract_hud", None) is not None:
            try:
                self._extract_hud.destroy()
            except Exception:
                pass
            self._extract_hud = None
        self.btn_run.configure(state="normal", fg_color=ACCENT_TEAL, text=self.t("btn_run"))

        self.selected_files.clear()
        for w in list(self.file_scroll.winfo_children()):
            if isinstance(w, FileRow):
                w.destroy()
        self.lbl_empty_list.pack(pady=30)
        self.all_records.clear()
        self._populate_tree([])
        self._rebuild_filter_tabs()
        self._refresh_badges()

    def _choose_out(self):
        d = filedialog.askdirectory(title=self.t("dlg_select_out"))
        if d:
            self.output_dir.set(os.path.normpath(d))

    def _refresh_badges(self):
        if self.current_nav == "compare":
            self._update_badges_for_compare()
            return
        elif self.current_nav == "model_comp":
            self._update_badges_for_model_compare()
            return
        self.badge_files.set_value(len(self.selected_files))
        self.badge_items.set_value(len(self.all_records))
        self.badge_qty.set_value(sum(r.get("Qty", 0) for r in self.all_records))

    # ──────────────────────────────────────────────────────────────────────────
    # USER INTERACTIONS (Drag & Drop, Copy, Toast, Context Menu, Collapse)
    # ──────────────────────────────────────────────────────────────────────────
    def _on_drop_files(self, files):
        """Native Windows drag & drop handler for PDF files, Excel files and folders."""
        has_excel = any(
            isinstance(item, (str, bytes)) and (
                (item.decode("utf-8", errors="ignore") if isinstance(item, bytes) else item).lower().endswith((".xlsx", ".xls"))
            ) for item in files
        )

        if self.current_nav == "model_comp":
            self.view_model_compare.handle_drop_files(files)
            return

        if self.current_nav == "compare" or has_excel:
            if self.current_nav != "compare":
                self._set_nav_active("compare")
            self.view_compare.handle_drop_files(files)
            return

        added = 0
        oversized = []
        for item in files:
            if isinstance(item, bytes):
                item = item.decode("utf-8", errors="ignore")
            item = os.path.normpath(item)
            if os.path.isdir(item):
                for root, _, fs in os.walk(item):
                    for f in fs:
                        if f.lower().endswith(".pdf"):
                            p = os.path.normpath(os.path.join(root, f))
                            size = os.path.getsize(p) if os.path.exists(p) else 0
                            if size > MAX_FILE_SIZE_BYTES:
                                oversized.append((f, format_file_size(size)))
                                continue
                            if p not in self.selected_files:
                                self.selected_files.append(p)
                                self._add_file_row(p)
                                added += 1
            elif os.path.isfile(item) and item.lower().endswith(".pdf"):
                size = os.path.getsize(item) if os.path.exists(item) else 0
                if size > MAX_FILE_SIZE_BYTES:
                    oversized.append((os.path.basename(item), format_file_size(size)))
                    continue
                if item not in self.selected_files:
                    self.selected_files.append(item)
                    self._add_file_row(item)
                    added += 1
        self._refresh_badges()
        if oversized:
            names = "\n• " + "\n• ".join(f"{name} ({sz})" for name, sz in oversized[:10])
            if len(oversized) > 10:
                names += f"\n... và {len(oversized) - 10} file khác"
            messagebox.showwarning(
                self.t("title_file_too_large"),
                self.t("err_batch_too_large", files=names, max=f"{MAX_FILE_SIZE_MB} MB")
            )
            self.show_toast(f"⚠️ Đã bỏ qua {len(oversized)} file > {MAX_FILE_SIZE_MB}MB!")
        if added > 0:
            self._log(self.t("log_added_files", count=added, folder="Drag & Drop"), "info")

    def _toggle_log(self):
        """Collapses or expands the bottom processing log to maximize table preview."""
        self.is_log_collapsed = not self.is_log_collapsed
        if self.is_log_collapsed:
            self.log_box.pack_forget()
            self.log_card.configure(height=38)
            self.btn_toggle_log.configure(text=self.t("log_expand"))
        else:
            self.log_card.configure(height=155)
            self.log_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))
            self.btn_toggle_log.configure(text=self.t("log_collapse"))

    def _show_toast(self, message: str):
        """Displays a modern floating toast notification in the preview card."""
        self.toast_lbl.configure(text=message)
        self.toast_frame.place(relx=0.5, rely=0.90, anchor="center")
        self.toast_frame.lift()
        if self._toast_timer:
            self.after_cancel(self._toast_timer)
        self._toast_timer = self.after(2200, lambda: self.toast_frame.place_forget())

    def show_toast(self, message: str):
        self._show_toast(message)

    def _on_tree_double_click(self, event):
        """Double clicking a table row copies the Part No to clipboard."""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        vals = self.tree.item(item_id, "values")
        if len(vals) >= 5:
            part_no = vals[4]
            if part_no:
                self.clipboard_clear()
                self.clipboard_append(str(part_no))
                self._show_toast(self.t("toast_copied", val=part_no))

    def _on_tree_right_click(self, event):
        """Right click context menu for copying row elements or full TSV row."""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        self.tree.selection_set(item_id)
        vals = self.tree.item(item_id, "values")
        if not vals:
            return

        self.ctx_menu.delete(0, "end")
        part_no = vals[4] if len(vals) > 4 else ""
        remarks = vals[8] if len(vals) > 8 else ""
        rating = vals[6] if len(vals) > 6 else ""
        sn = vals[5] if len(vals) > 5 else ""

        def copy_val(v):
            if v:
                self.clipboard_clear()
                self.clipboard_append(str(v))
                self._show_toast(self.t("toast_copied", val=v))

        def copy_row():
            row_txt = "\t".join(str(x) for x in vals)
            self.clipboard_clear()
            self.clipboard_append(row_txt)
            self._show_toast(self.t("toast_copied", val=f"{part_no} (Row)"))

        if part_no:
            self.ctx_menu.add_command(label=f"{self.t('ctx_copy_part')}:  {part_no}",
                                      command=lambda: copy_val(part_no))
        if remarks:
            rem_disp = remarks[:24] + ("…" if len(remarks) > 24 else "")
            self.ctx_menu.add_command(label=f"{self.t('ctx_copy_remark')}:  {rem_disp}",
                                      command=lambda: copy_val(remarks))
        if rating:
            rat_disp = rating[:24] + ("…" if len(rating) > 24 else "")
            self.ctx_menu.add_command(label=f"{self.t('ctx_copy_rating')}:  {rat_disp}",
                                      command=lambda: copy_val(rating))
        if sn:
            self.ctx_menu.add_command(label=f"{self.t('ctx_copy_sn')}:  {sn}",
                                      command=lambda: copy_val(sn))
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label=self.t("ctx_copy_row"), command=copy_row)

        try:
            self.ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx_menu.grab_release()

    # ──────────────────────────────────────────────────────────────────────────
    # PROCESS FILTER TABS
    # ──────────────────────────────────────────────────────────────────────────
    def _rebuild_filter_tabs(self):
        for w in self.filter_frame.winfo_children():
            w.destroy()

        processes = sorted({r["Process"] for r in self.all_records if r.get("Process")})

        def make_filter(p=None):
            def cmd():
                self._active_filter = p
                self._on_search()
            return cmd

        btn_all = ctk.CTkButton(
            self.filter_frame, text=self.t("filter_all"),
            width=65, height=28, corner_radius=6,
            fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
            hover_color=("#0F766E", "#00A88C"), font=FONT_BADGE,
            command=make_filter(None))
        btn_all.pack(side="left", padx=(8, 4), pady=4)

        for proc in processes:
            col = proc_color(proc)
            short = proc.split("AR-")[-1] if "AR-" in proc else proc[:10]
            label = f"{proc[:3]} · {short}"
            btn = ctk.CTkButton(
                self.filter_frame,
                text=label, width=90, height=28, corner_radius=6,
                fg_color=BG_SURFACE,
                text_color=col,
                border_width=1, border_color=col,
                hover_color=BG_HOVER,
                font=FONT_BADGE,
                command=make_filter(proc))
            btn.pack(side="left", padx=4, pady=4)

    # ──────────────────────────────────────────────────────────────────────────
    # TREE / SEARCH
    # ──────────────────────────────────────────────────────────────────────────
    def _populate_tree(self, records):
        self.tree.delete(*self.tree.get_children())
        if records:
            if hasattr(self, "empty_preview"):
                self.empty_preview.place_forget()
        else:
            if hasattr(self, "empty_preview"):
                self.empty_preview.place(relx=0.5, rely=0.45, anchor="center")
                self.empty_preview.lift()

        for i, r in enumerate(records):
            tag = "odd" if i % 2 else "even"
            self.tree.insert("", "end", tags=(tag,), values=(
                r.get("Seq", ""),
                r.get("Item_No", ""),
                r.get("Process", ""),
                r.get("Model_Variant", ""),
                r.get("Part_No", ""),
                r.get("Internal_SN", ""),
                r.get("Rating_Spec", ""),
                r.get("Qty", ""),
                r.get("Remarks_Locations", ""),
                r.get("Alternate_Parts", ""),
                r.get("Source_Page", ""),
            ))

        if hasattr(self, "lbl_preview_count"):
            if self.all_records:
                self.lbl_preview_count.configure(
                    text=self.t("preview_count", filtered=len(records), total=len(self.all_records)))
            else:
                self.lbl_preview_count.configure(text="")

    def _on_search(self, *_):
        q = self.search_var.get().strip().lower()
        proc_filter = self._active_filter

        filtered = self.all_records
        if proc_filter:
            filtered = [r for r in filtered if r.get("Process") == proc_filter]
        if q:
            filtered = [
                r for r in filtered if any(
                    q in str(r.get(k, "")).lower()
                    for k in ("Part_No", "Internal_SN", "Rating_Spec",
                              "Remarks_Locations", "Process", "Model_Variant", "Item_No")
                )
            ]
        self._populate_tree(filtered)

    # ──────────────────────────────────────────────────────────────────────────
    # LOGGING
    # ──────────────────────────────────────────────────────────────────────────
    def _log(self, msg: str, tag: str = "info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}] ", "muted")
        self.log_box.insert("end", f"{msg}\n", tag)
        self.log_box.see("end")

    # ── EXTRACTION THREAD & ANIMATION ─────────────────────────────────────────
    def _start_thread(self):
        if not self.selected_files:
            messagebox.showwarning(self.t("msg_no_files_title"), self.t("msg_no_files_body"))
            return
        if self.is_running:
            return

        self.is_running = True
        self.btn_run.configure(
            state="disabled", fg_color=("#94A3B8", "#1A3A35"),
            text="⚡ Đang trích xuất BOM AI... (0%)"
        )
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if hasattr(self, "empty_preview"):
            self.empty_preview.place_forget()
        self.progress.set(0)

        # Remove previous HUD if any
        if getattr(self, "_extract_hud", None) is not None:
            try:
                self._extract_hud.destroy()
            except Exception:
                pass
            self._extract_hud = None

        # Build Cyberpunk Hologram HUD Card over self.preview_card
        is_dark = (self.current_theme == "dark")
        target_parent = getattr(self, "preview_card", self.tree_frame)
        self._extract_hud = ctk.CTkFrame(
            target_parent, fg_color="#0B132B" if is_dark else "#0F172A",
            corner_radius=12, border_width=1.5, border_color="#00F0FF"
        )
        self._extract_hud.place(relx=0.5, rely=0.42, anchor="center")

        hud_inner = ctk.CTkFrame(self._extract_hud, fg_color="transparent")
        hud_inner.pack(padx=24, pady=16)

        # Row 1: Header + Telemetry
        h_row1 = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row1.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(h_row1, text="⚡", font=("Segoe UI", 16)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            h_row1, text="VIPQC AI BOM EXTRACTION ENGINE",
            font=("Segoe UI", 11, "bold"), text_color="#00F0FF"
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(
            h_row1, text="60 FPS • NEURAL PDF PARSER",
            font=("Consolas", 9, "bold"), text_color="#38BDF8"
        ).pack(side="right")

        # Row 2: Progress bar + Percentage
        h_row2 = ctk.CTkFrame(hud_inner, fg_color="transparent")
        h_row2.pack(fill="x", pady=(0, 8))

        ext_prog = ctk.CTkProgressBar(
            h_row2, width=380, height=10, corner_radius=5,
            fg_color="#1E293B", progress_color="#00F0FF"
        )
        ext_prog.set(0.0)
        ext_prog.pack(side="left", padx=(0, 10))

        lbl_ext_pct = ctk.CTkLabel(
            h_row2, text="0%", font=("Consolas", 11, "bold"), text_color="#00F0FF", width=44
        )
        lbl_ext_pct.pack(side="left")

        # Row 3: Live Terminal Phase
        lbl_ext_log = ctk.CTkLabel(
            hud_inner, text="[01/05] 📡 Khởi tạo VIPQC Neural Vision & PDF Engine...",
            font=("Consolas", 10), text_color="#E2E8F0", anchor="w"
        )
        lbl_ext_log.pack(fill="x")

        # Snapshot configuration on main thread before launching worker
        files_to_process = list(self.selected_files)
        out_dir_path = self.output_dir.get().strip() or "excel_results"
        opt_export_ind = bool(self.export_individual.get())
        opt_export_batch = bool(self.export_batch.get())

        ext_data = {
            "done": False,
            "err": None,
            "records": [],
            "batch_data": [],
            "last_exports": [],
            "files_count": len(files_to_process)
        }

        def safe_log(msg, tag="info"):
            try:
                self.after(0, lambda m=msg, t=tag: self._log(m, t))
            except Exception:
                pass

        def _bg_extract():
            try:
                files = files_to_process
                n = len(files)
                os.makedirs(out_dir_path, exist_ok=True)
                batch_data = []
                last_exports = []
                all_recs = []

                safe_log(self.t("log_start", n=n), "bold")

                for i, fpath in enumerate(files, 1):
                    fname = os.path.basename(fpath)
                    safe_log(self.t("log_analyzing", i=i, n=n, fname=fname), "info")

                    try:
                        meta, flat, mat = extract_full_bom(fpath)
                        batch_data.append((meta, flat, mat))

                        procs = ", ".join(meta.get("processes", [])) or "N/A"
                        safe_log(
                            self.t("log_parsed",
                                   pwb=meta.get("pwb_code") or "N/A",
                                   procs=procs,
                                   items=meta.get("total_items", 0),
                                   qty=meta.get("total_parts_count", 0)),
                            "success")

                        all_recs.extend(flat)

                        if opt_export_ind:
                            base = os.path.splitext(fname)[0]
                            xl = os.path.join(out_dir_path, f"BOM_{base}.xlsx")
                            export_bom_to_excel(meta, flat, mat, xl)
                            last_exports.append(xl)
                            safe_log(self.t("log_exported", name=f"BOM_{base}.xlsx"), "success")

                    except Exception as ex:
                        safe_log(self.t("log_error", fname=fname, err=ex), "error")

                if opt_export_batch and batch_data:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    master_xl = os.path.join(out_dir_path, f"Master_BOM_{ts}.xlsx")
                    export_batch_bom_to_excel(batch_data, master_xl)
                    last_exports.insert(0, master_xl)
                    safe_log(self.t("log_master", name=f"Master_BOM_{ts}.xlsx"), "success")

                ext_data["records"] = all_recs
                ext_data["batch_data"] = batch_data
                ext_data["last_exports"] = last_exports
            except Exception as e:
                ext_data["err"] = e
            finally:
                ext_data["done"] = True

        threading.Thread(target=_bg_extract, daemon=True).start()

        # Animation parameters (3.5 seconds)
        start_time = time.time()
        anim_duration = 3.5

        def _update_extract_anim():
            if not getattr(self, "is_running", False):
                return

            now = time.time()
            elapsed = now - start_time
            t = min(1.0, elapsed / anim_duration)

            ext_prog.set(t)
            self.progress.set(t)
            pct = int(t * 100)
            lbl_ext_pct.configure(text=f"{pct}%")
            self.btn_run.configure(text=f"⚡ Đang trích xuất BOM AI... ({pct}%)")

            lang = getattr(self, "current_lang", "vi")
            if lang == "zh":
                if t < 0.20:
                    lbl_ext_log.configure(text="[01/05] 📡 正在启动 VIPQC 视觉与文档解析引擎...")
                elif t < 0.40:
                    lbl_ext_log.configure(text="[02/05] 📑 解析 Working Manual 表格结构与列坐标...")
                elif t < 0.65:
                    lbl_ext_log.configure(text="[03/05] 🔬 提取 SMD 芯片、IC、电阻与插件元件...")
                elif t < 0.85:
                    lbl_ext_log.configure(text="[04/05] ⚡ 规范化工序代码、位置代码并统计数量...")
                else:
                    lbl_ext_log.configure(text="[05/05] 🎯 正在完成数据导出并生成 Excel 结果！")
            elif lang == "en":
                if t < 0.20:
                    lbl_ext_log.configure(text="[01/05] 📡 Initializing VIPQC Neural Vision & PDF Engine...")
                elif t < 0.40:
                    lbl_ext_log.configure(text="[02/05] 📑 Parsing Working Manual tables & column layout...")
                elif t < 0.65:
                    lbl_ext_log.configure(text="[03/05] 🔬 Extracting SMD chips, ICs, resistors & components...")
                elif t < 0.85:
                    lbl_ext_log.configure(text="[04/05] ⚡ Normalizing Process Codes, Ref Des & aggregating Qty...")
                else:
                    lbl_ext_log.configure(text="[05/05] 🎯 Finalizing extraction & generating Excel report!")
            else:
                if t < 0.20:
                    lbl_ext_log.configure(text="[01/05] 📡 Khởi tạo VIPQC Neural Vision & PDF Engine...")
                elif t < 0.40:
                    lbl_ext_log.configure(text="[02/05] 📑 Phân tích bảng biểu Working Manual & tọa độ cột...")
                elif t < 0.65:
                    lbl_ext_log.configure(text="[03/05] 🔬 Trích xuất linh kiện SMD, IC, Resistor & chân cắm...")
                elif t < 0.85:
                    lbl_ext_log.configure(text="[04/05] ⚡ Chuẩn hóa Process Code, Ref Des & tính tổng Qty...")
                else:
                    lbl_ext_log.configure(text="[05/05] 🎯 Hoàn tất trích xuất & ghi dữ liệu ra tệp Excel!")

            if t < 1.0 or not ext_data["done"]:
                self.after(16, _update_extract_anim)
            else:
                self._finish_extraction(ext_data)

        self.after(20, _update_extract_anim)

    def _finish_extraction(self, ext_data: dict):
        self.is_running = False

        # Pulse border on preview_card
        if hasattr(self, "preview_card") and self.preview_card:
            self.preview_card.configure(border_color="#00F0FF", border_width=2)
            self.after(180, lambda: self.preview_card.configure(border_color=BORDER_CLR, border_width=1))

        # Cleanup HUD
        if getattr(self, "_extract_hud", None) is not None:
            try:
                self._extract_hud.destroy()
            except Exception:
                pass
            self._extract_hud = None

        self.btn_run.configure(state="normal", fg_color=ACCENT_TEAL, text=self.t("btn_run"))

        if ext_data.get("err"):
            self._log(f"Error: {ext_data['err']}", "error")
            messagebox.showerror("Lỗi Trích Xuất", f"Đã xảy ra lỗi khi trích xuất BOM:\n{ext_data['err']}")
            return

        self.all_records = ext_data.get("records", [])
        self.last_exports = ext_data.get("last_exports", [])

        # Populate tree and update UI
        self._populate_tree(self.all_records)
        self._rebuild_filter_tabs()
        self._refresh_badges()

        n = ext_data.get("files_count", len(self.selected_files))
        total_qty = sum(r.get("Qty", 0) for r in self.all_records)
        self._log(self.t("log_finish", rows=len(self.all_records), qty=total_qty), "bold")
        self.lbl_status.configure(text=self.t("status_done", count=len(self.all_records), n=n))

        if self.last_exports:
            self.btn_open_xl.configure(state="normal")

        self.show_toast(f"✅ Đã trích xuất thành công {len(self.all_records)} linh kiện!")

    # ──────────────────────────────────────────────────────────────────────────
    # QUICK ACCESS
    # ──────────────────────────────────────────────────────────────────────────
    def _open_excel(self):
        if self.last_exports and os.path.exists(self.last_exports[0]):
            os.startfile(self.last_exports[0])
        else:
            messagebox.showwarning(self.t("msg_not_found_title"), self.t("msg_not_found_body"))

    def _open_folder(self):
        d = self.output_dir.get().strip() or "excel_results"
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    # ──────────────────────────────────────────────────────────────────────────
    # GITHUB AUTO-UPDATE INTEGRATION
    # ──────────────────────────────────────────────────────────────────────────
    def _auto_check_update(self):
        """Silently checks for new releases on GitHub in a background thread."""
        def _check():
            try:
                res = updater.check_for_updates(timeout=6)
                if res.get("has_update"):
                    self.after(0, lambda: self._show_update_dialog(res))
            except Exception:
                pass

        threading.Thread(target=_check, daemon=True).start()

    def _manual_check_update(self):
        """Triggered when user clicks the version badge in the sidebar footer."""
        self.show_toast("📡 Đang kiểm tra bản cập nhật từ GitHub...")

        def _check():
            try:
                res = updater.check_for_updates(timeout=8)
                if res.get("has_update"):
                    self.after(0, lambda: self._show_update_dialog(res))
                elif res.get("error"):
                    self.after(0, lambda: self.show_toast("⚠️ Không thể kết nối tới GitHub. Vui lòng kiểm tra mạng!"))
                else:
                    self.after(0, lambda: self.show_toast(f"✅ Bạn đang dùng phiên bản mới nhất (v{updater.CURRENT_VERSION})!"))
            except Exception:
                self.after(0, lambda: self.show_toast("⚠️ Lỗi kiểm tra cập nhật."))

        threading.Thread(target=_check, daemon=True).start()

    def _show_update_dialog(self, update_info: dict):
        """Displays the update notification popup."""
        try:
            UpdateDialog(self, update_info)
        except Exception as e:
            print(f"Error opening update dialog: {e}")


# ─── Entry Point ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = BOMExtractorApp()
    app.mainloop()
