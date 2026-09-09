"""
VIPQC AI — Anomaly Report View Module (Full-Width Table & Image Marking Studio)
Modern CustomTkinter GUI featuring:
1. Full-width 18-column Table without cramped sidebar
2. Direct image viewing & interaction from rows
3. Dedicated Image Marking & Zoom Studio:
   - Pan, Zoom in/out, Mousewheel zoom, 100%, Fit
   - Rotate 90 degrees
   - Annotation Tools: Rectangle Box, Oval, Arrow, Freehand Pen, Text Notes
   - Multi-color palette & stroke width
   - Undo (Ctrl+Z) & Clear
   - Save annotated image directly back to Supabase Cloud & Download to PC
4. Action Bar & Context Menu: Add, Edit, Delete, Refresh, Export Excel, Admin lock
"""

import os
import io
import math
import hashlib
import tempfile
import threading
from datetime import datetime
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox, simpledialog
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont, ImageTk

import anomaly_service as svc

# Design Tokens (Light, Dark)
BG_DEEP     = ("#F1F5F9", "#0B0F1A")
BG_CARD     = ("#FFFFFF", "#131929")
BG_SURFACE  = ("#E2E8F0", "#1C2438")
BG_HOVER    = ("#CBD5E1", "#232D45")
ACCENT_TEAL = ("#0D9488", "#00C9A7")
ACCENT_BLUE = ("#2563EB", "#3D8EFF")
ACCENT_AMBER= ("#D97706", "#FFB830")
ACCENT_RED  = ("#DC2626", "#EF4444")
TEXT_PRIMARY= ("#0F172A", "#E8EDF5")
TEXT_MUTED  = ("#64748B", "#7A8BA6")
BORDER_CLR  = ("#CBD5E1", "#2A3650")

FONT_H1     = ("Segoe UI", 15, "bold")
FONT_H2     = ("Segoe UI", 13, "bold")
FONT_BODY   = ("Segoe UI", 12)
FONT_BOLD   = ("Segoe UI", 12, "bold")
FONT_SMALL  = ("Segoe UI", 11)


def safe_after(widget, ms: int, callback):
    """Safely invokes callback on main thread if widget still exists."""
    try:
        if widget and widget.winfo_exists():
            widget.after(ms, callback)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# INTERNATIONALIZATION (I18N) DICTIONARY & NORMALIZERS
# ─────────────────────────────────────────────────────────────────────────────
I18N_ANOMALY = {
    "vi": {
        "search_placeholder": "🔍 Tìm kiếm sản phẩm, công đoạn, lỗi, người phụ trách...",
        "all_proc": "Tất cả công đoạn",
        "proc_options": ["Tất cả công đoạn", "CHA", "SCP", "RAD", "CHP", "SMT", "Lắp Ráp", "KCS / Final QC"],
        "all_prog": "Tất cả tiến độ",
        "prog_options": ["Tất cả tiến độ", "Chưa thực hiện", "Đang thực hiện", "Đã hoàn thành"],
        "btn_admin_locked": "🔒 Admin: Đã Khóa",
        "btn_admin_unlocked": "🔓 Admin: Đã Mở",
        "btn_export": "📊 Xuất Excel",
        "btn_refresh": "🔄 Làm Mới",
        "btn_refreshing": "⏳ Đang tải...",
        "btn_delete": "🗑️ Xóa",
        "btn_edit": "✏️ Sửa",
        "btn_open_marking": "🖼️ Soi & Vẽ Ảnh",
        "btn_add": "➕ Thêm Báo Cáo",
        "tbl_count": "Danh Sách Báo Cáo: {count} bản ghi",
        "hint_label": "💡 Click vào 'Ảnh Lỗi' trên bảng để Soi & Đánh Dấu (Studio). Click đúp vào dòng để Sửa trong Drawer.",
        "ctx_open_marking": "🖼️ Soi & Đánh Dấu Ảnh Lỗi (Studio)",
        "ctx_edit": "✏️ Chỉnh Sửa Báo Cáo",
        "ctx_delete": "🗑️ Xóa Báo Cáo Này",
        "ctx_copy": "📋 Sao Chép Thông Tin Dòng",
        "prog_val_open": "Chưa thực hiện",
        "prog_val_in_progress": "Đang thực hiện",
        "prog_val_done": "Đã hoàn thành",
        "drawer_title_create": "➕ Thêm Báo Cáo Mới",
        "drawer_title_edit": "✏️ Sửa Báo Cáo",
        "sec_inspection": "📌 Thông Tin Kiểm Tra & Số Lượng",
        "lbl_date": "Ngày Tháng (YYYY-MM-DD):",
        "lbl_process": "Công Đoạn:",
        "lbl_product": "Sản Phẩm (Model / PWB): *",
        "ph_product": "Nhập tên model hoặc mã PWB...",
        "lbl_machine": "Máy Móc / Chuyền / Thiết Bị:",
        "ph_machine": "Ví dụ: Line 2, Máy 12...",
        "lbl_tot": "SL Kiểm:",
        "lbl_def": "SL Lỗi:",
        "lbl_rate": "Tỷ Lệ Lỗi:",
        "lbl_resp": "Người Chịu Trách Nhiệm (Manager/Leader):",
        "lbl_pic": "Người Phụ Trách (Inspector/PIC):",
        "lbl_prog": "Tiến Độ Khắc Phục:",
        "drawer_prog_options": ["Chưa thực hiện", "Đang thực hiện", "Đã hoàn thành"],
        "sec_image": "🖼️ Hình Ảnh Minh Họa & Đánh Dấu",
        "img_empty_hint": "[Chưa có ảnh]\nBấm 'Chọn File' hoặc chụp màn hình rồi dán (Ctrl+V)",
        "btn_studio": "🔍 Soi & Đánh Dấu (Studio)",
        "btn_pick": "📁 Chọn",
        "btn_paste": "📋 Dán",
        "sec_defect": "📝 Mô Tả Hiện Tượng & Đối Sách",
        "lbl_desc": "Mô Tả Lỗi (Hiện tượng bất thường):",
        "lbl_cause": "Nguyên Nhân:",
        "lbl_counter": "Biện Pháp Cải Tiến:",
        "lbl_sop": "Tiêu Chuẩn Hóa SOP:",
        "lbl_notes": "Ghi Chú:",
        "btn_close": "Đóng",
        "btn_cancel": "Hủy",
        "btn_save": "💾 Lưu Báo Cáo",
        "saving_txt": "⏳ Đang lưu dữ liệu...",
        "pwd_title": "Xác Thực Quản Trị",
        "pwd_heading": "Nhập Mật Khẩu Quản Trị (QC)",
        "pwd_prompt": "Nhập mật khẩu...",
        "pwd_btn": "Xác Nhận",
        "pwd_err": "❌ Mật khẩu không chính xác! Vui lòng thử lại.",
        "pwd_empty": "Vui lòng nhập mật khẩu!",
        "chpwd_title": "Đổi Mật Khẩu Quản Trị",
        "chpwd_heading": "🔑 Đổi Mật Khẩu Quản Trị",
        "chpwd_old": "Mật khẩu hiện tại",
        "chpwd_new": "Mật khẩu mới (ít nhất 4 ký tự)",
        "chpwd_confirm": "Xác nhận mật khẩu mới",
        "chpwd_save": "💾 Lưu Mật Khẩu",
        "studio_title": "🖼️ Studio Soi & Đánh Dấu Ảnh Lỗi",
        "studio_tools_title": "🎨 CÔNG CỤ VẼ & ĐÁNH DẤU:",
        "studio_pan": "✋ Di chuyển",
        "studio_rect": "🟥 Hộp",
        "studio_oval": "⭕ Tròn",
        "studio_arrow": "➡️ Mũi tên",
        "studio_pen": "✏️ Bút vẽ",
        "studio_text": "🔤 Chữ",
        "studio_undo": "↩️ Hoàn tác",
        "studio_clear": "🗑️ Xóa hết nét",
        "studio_fullscreen": "⛶ Toàn Màn Hình",
        "studio_unfullscreen": "🗗 Thu Nhỏ",
        "studio_rotate": "🔄 Xoay 90° (R)",
        "studio_fit": "🖼️ Vừa Khung",
        "studio_close": "✕ Đóng",
        "studio_download": "📥 Tải Về Máy",
        "studio_save_cloud": "💾 Lưu Lên Báo Cáo",
        "no_img_title": "Chưa có ảnh",
        "no_img_prompt": "Báo cáo này chưa có hình ảnh đính kèm.\nBạn có muốn mở Drawer để thêm ảnh không?",
        "no_img_alert": "Báo cáo này không có hình ảnh đính kèm để xem hoặc đánh dấu.",
        "copied_toast": "📋 Đã sao chép thông tin báo cáo vào Clipboard!",
        "saved_toast": "✅ Đã lưu báo cáo bất thường thành công!",
        "deleted_toast": "🗑️ Đã xóa báo cáo và ảnh trên Cloud thành công!",
        "export_success": "Đã xuất {count} báo cáo ra Excel:\n{path}\n\nBạn có muốn mở file ngay không?",
    },
    "zh": {
        "search_placeholder": "🔍 搜索产品、过程、不良现象、担当者...",
        "all_proc": "全部过程",
        "proc_options": ["全部过程", "CHA", "SCP", "RAD", "CHP", "SMT", "组装", "终检 / Final QC"],
        "all_prog": "全部进度",
        "prog_options": ["全部进度", "未开始", "进行中", "已完成"],
        "btn_admin_locked": "🔒 管理员: 已锁定",
        "btn_admin_unlocked": "🔓 管理员: 已解锁",
        "btn_export": "📊 导出 Excel",
        "btn_refresh": "🔄 刷新",
        "btn_refreshing": "⏳ 正在加载...",
        "btn_delete": "🗑️ 删除",
        "btn_edit": "✏️ 编辑",
        "btn_open_marking": "🖼️ 查看与标记",
        "btn_add": "➕ 新增报告",
        "tbl_count": "异常报告列表: {count} 条记录",
        "hint_label": "💡 点击表格中的'图片'即可在工作室中查看与标记。双击行可在抽屉中编辑。",
        "ctx_open_marking": "🖼️ 查看与标记图片 (工作室)",
        "ctx_edit": "✏️ 编辑此报告",
        "ctx_delete": "🗑️ 删除此报告",
        "ctx_copy": "📋 复制此行信息",
        "prog_val_open": "未开始",
        "prog_val_in_progress": "进行中",
        "prog_val_done": "已完成",
        "drawer_title_create": "➕ 新增异常报告",
        "drawer_title_edit": "✏️ 编辑异常报告",
        "sec_inspection": "📌 检验信息与数量",
        "lbl_date": "日期 (YYYY-MM-DD):",
        "lbl_process": "过程:",
        "lbl_product": "产品 (型号 / PWB): *",
        "ph_product": "输入型号名称或 PWB 代码...",
        "lbl_machine": "设备 / 线体 / 机台:",
        "ph_machine": "例如: Line 2, 机台 12...",
        "lbl_tot": "检查数:",
        "lbl_def": "不良数:",
        "lbl_rate": "不良率:",
        "lbl_resp": "责任人 (主管/组长):",
        "lbl_pic": "担当者 (检验员/PIC):",
        "lbl_prog": "改善进度:",
        "drawer_prog_options": ["未开始", "进行中", "已完成"],
        "sec_image": "🖼️ 现象图片与标记",
        "img_empty_hint": "[暂无图片]\n点击'选择'或截屏后粘贴 (Ctrl+V)",
        "btn_studio": "🔍 查看与标记 (工作室)",
        "btn_pick": "📁 选择",
        "btn_paste": "📋 粘贴",
        "sec_defect": "📝 不良现象描述与改善对策",
        "lbl_desc": "不良现象描述:",
        "lbl_cause": "原因分析:",
        "lbl_counter": "改善对策:",
        "lbl_sop": "SOP 标准化:",
        "lbl_notes": "备注:",
        "btn_close": "关闭",
        "btn_cancel": "取消",
        "btn_save": "💾 保存报告",
        "saving_txt": "⏳ 正在保存数据...",
        "pwd_title": "管理员身份验证",
        "pwd_heading": "请输入管理员密码 (QC)",
        "pwd_prompt": "输入密码...",
        "pwd_btn": "确认",
        "pwd_err": "❌ 密码错误！请重试。",
        "pwd_empty": "请输入密码！",
        "chpwd_title": "修改管理员密码",
        "chpwd_heading": "🔑 修改管理员密码",
        "chpwd_old": "当前密码",
        "chpwd_new": "新密码 (至少4位字符)",
        "chpwd_confirm": "确认新密码",
        "chpwd_save": "💾 保存密码",
        "studio_title": "🖼️ 不良图片查看与标记工作室",
        "studio_tools_title": "🎨 绘制与标记工具:",
        "studio_pan": "✋ 移动",
        "studio_rect": "🟥 矩形",
        "studio_oval": "⭕ 椭圆",
        "studio_arrow": "➡️ 箭头",
        "studio_pen": "✏️ 画笔",
        "studio_text": "🔤 文字",
        "studio_undo": "↩️ 撤销",
        "studio_clear": "🗑️ 清除全部",
        "studio_fullscreen": "⛶ 全屏显示",
        "studio_unfullscreen": "🗗 退出全屏",
        "studio_rotate": "🔄 旋转 90° (R)",
        "studio_fit": "🖼️ 适应窗口",
        "studio_close": "✕ 关闭",
        "studio_download": "📥 下载到电脑",
        "studio_save_cloud": "💾 保存到报告",
        "no_img_title": "暂无图片",
        "no_img_prompt": "该报告暂未上传图片。\n是否打开抽屉进行添加？",
        "no_img_alert": "此报告未附带图片，无法查看或标记。",
        "copied_toast": "📋 已复制报告信息到剪贴板！",
        "saved_toast": "✅ 异常报告保存成功！",
        "deleted_toast": "🗑️ 报告及云端图片删除成功！",
        "export_success": "已成功导出 {count} 条报告至 Excel:\n{path}\n\n是否立即打开文件？",
    },
    "en": {
        "search_placeholder": "🔍 Search product, process, defect, PIC...",
        "all_proc": "All Processes",
        "proc_options": ["All Processes", "CHA", "SCP", "RAD", "CHP", "SMT", "Assembly", "Final QC"],
        "all_prog": "All Progress",
        "prog_options": ["All Progress", "Pending", "In Progress", "Completed"],
        "btn_admin_locked": "🔒 Admin: Locked",
        "btn_admin_unlocked": "🔓 Admin: Unlocked",
        "btn_export": "📊 Export Excel",
        "btn_refresh": "🔄 Refresh",
        "btn_refreshing": "⏳ Loading...",
        "btn_delete": "🗑️ Delete",
        "btn_edit": "✏️ Edit",
        "btn_open_marking": "🖼️ Mark Image",
        "btn_add": "➕ Add Report",
        "tbl_count": "Report List: {count} records",
        "hint_label": "💡 Click 'Image' cell to inspect & mark (Studio). Double-click row to edit in Drawer.",
        "ctx_open_marking": "🖼️ Inspect & Mark Image (Studio)",
        "ctx_edit": "✏️ Edit Report",
        "ctx_delete": "🗑️ Delete Report",
        "ctx_copy": "📋 Copy Row Information",
        "prog_val_open": "Pending",
        "prog_val_in_progress": "In Progress",
        "prog_val_done": "Completed",
        "drawer_title_create": "➕ Add New Report",
        "drawer_title_edit": "✏️ Edit Report",
        "sec_inspection": "📌 Inspection Info & Quantities",
        "lbl_date": "Date (YYYY-MM-DD):",
        "lbl_process": "Process:",
        "lbl_product": "Product (Model / PWB): *",
        "ph_product": "Enter model name or PWB code...",
        "lbl_machine": "Machine / Line / Equipment:",
        "ph_machine": "e.g. Line 2, Machine 12...",
        "lbl_tot": "Total Qty:",
        "lbl_def": "Defect Qty:",
        "lbl_rate": "Defect Rate:",
        "lbl_resp": "Responsible Person (Manager/Leader):",
        "lbl_pic": "PIC (Inspector/PIC):",
        "lbl_prog": "Action Progress:",
        "drawer_prog_options": ["Pending", "In Progress", "Completed"],
        "sec_image": "🖼️ Illustration Image & Marking",
        "img_empty_hint": "[No image]\nClick 'Browse' or paste screenshot (Ctrl+V)",
        "btn_studio": "🔍 Inspect & Mark (Studio)",
        "btn_pick": "📁 Browse",
        "btn_paste": "📋 Paste",
        "sec_defect": "📝 Defect Description & Actions",
        "lbl_desc": "Defect Description:",
        "lbl_cause": "Root Cause:",
        "lbl_counter": "Countermeasures:",
        "lbl_sop": "SOP Standardization:",
        "lbl_notes": "Notes:",
        "btn_close": "Close",
        "btn_cancel": "Cancel",
        "btn_save": "💾 Save Report",
        "saving_txt": "⏳ Saving data...",
        "pwd_title": "Administrator Verification",
        "pwd_heading": "Enter Administrator Password (QC)",
        "pwd_prompt": "Enter password...",
        "pwd_btn": "Confirm",
        "pwd_err": "❌ Incorrect password! Please try again.",
        "pwd_empty": "Please enter password!",
        "chpwd_title": "Change Admin Password",
        "chpwd_heading": "🔑 Change Admin Password",
        "chpwd_old": "Current Password",
        "chpwd_new": "New Password (at least 4 chars)",
        "chpwd_confirm": "Confirm New Password",
        "chpwd_save": "💾 Save Password",
        "studio_title": "🖼️ Defect Image Inspection & Marking Studio",
        "studio_tools_title": "🎨 DRAWING & MARKING TOOLS:",
        "studio_pan": "✋ Pan",
        "studio_rect": "🟥 Box",
        "studio_oval": "⭕ Oval",
        "studio_arrow": "➡️ Arrow",
        "studio_pen": "✏️ Pen",
        "studio_text": "🔤 Text",
        "studio_undo": "↩️ Undo",
        "studio_clear": "🗑️ Clear All",
        "studio_fullscreen": "⛶ Fullscreen",
        "studio_unfullscreen": "🗗 Exit Fullscreen",
        "studio_rotate": "🔄 Rotate 90° (R)",
        "studio_fit": "🖼️ Fit Screen",
        "studio_close": "✕ Close",
        "studio_download": "📥 Download",
        "studio_save_cloud": "💾 Save to Report",
        "no_img_title": "No Image",
        "no_img_prompt": "This report has no attached image.\nWould you like to open Drawer to add one?",
        "no_img_alert": "This report has no image to inspect or mark.",
        "copied_toast": "📋 Copied report details to clipboard!",
        "saved_toast": "✅ Anomaly report saved successfully!",
        "deleted_toast": "🗑️ Report and image deleted from Cloud!",
        "export_success": "Successfully exported {count} reports to Excel:\n{path}\n\nDo you want to open the file now?",
    }
}


def normalize_prog(p_str: str) -> str:
    """Normalizes raw progress string to canonical key: 'done', 'open', or 'in_progress'."""
    s = str(p_str or "").strip().lower()
    if any(k in s for k in ["hoàn thành", "已完成", "complete", "done"]):
        return "done"
    if any(k in s for k in ["chưa", "未开始", "pending", "not"]):
        return "open"
    return "in_progress"


def normalize_proc(p_str: str) -> str:
    """Normalizes process names for multi-language filter comparison."""
    s = str(p_str or "").strip().lower()
    if any(k in s for k in ["lắp ráp", "组装", "assembly"]):
        return "assembly"
    if any(k in s for k in ["kcs", "final qc", "终检"]):
        return "final_qc"
    return s


# ─────────────────────────────────────────────────────────────────────────────
# DIALOG: PASSWORD AUTHENTICATION
# ─────────────────────────────────────────────────────────────────────────────
class PasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent, on_success, lang_code: str = "vi"):
        super().__init__(parent)
        self.lang_code = lang_code
        t = I18N_ANOMALY.get(lang_code, I18N_ANOMALY["vi"])
        self.title(t["pwd_title"])
        self.geometry("400x250")
        self.resizable(False, False)
        self.configure(fg_color=BG_CARD)
        self.transient(parent)
        self.grab_set()

        self.on_success = on_success

        self.update_idletasks()
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(50, (sw - 400) // 2)
            y = max(50, (sh - 250) // 2)
            self.geometry(f"400x250+{x}+{y}")
            self.lift()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

        pad = {"padx": 24}

        lbl_icon = ctk.CTkLabel(self, text="🔒", font=("Segoe UI", 28))
        lbl_icon.pack(pady=(18, 4))

        lbl_title = ctk.CTkLabel(self, text=t["pwd_heading"], font=FONT_H1, text_color=TEXT_PRIMARY)
        lbl_title.pack(pady=(0, 12))

        self.entry_pass = ctk.CTkEntry(self, placeholder_text=t["pwd_prompt"], show="•",
                                       height=38, corner_radius=8, font=FONT_BODY)
        self.entry_pass.pack(fill="x", pady=(0, 8), **pad)
        self.entry_pass.focus()
        self.entry_pass.bind("<Return>", lambda e: self._submit())

        self.lbl_err = ctk.CTkLabel(self, text="", font=FONT_SMALL, text_color=ACCENT_RED)
        self.lbl_err.pack(pady=(0, 10))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", **pad)

        btn_cancel = ctk.CTkButton(btn_row, text=t["btn_cancel"], width=100, height=36, corner_radius=8,
                                   fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                                   command=self.destroy)
        btn_cancel.pack(side="left")

        self.btn_ok = ctk.CTkButton(btn_row, text=t["pwd_btn"], width=140, height=36, corner_radius=8,
                                    fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
                                    font=FONT_H2, command=self._submit)
        self.btn_ok.pack(side="right")

    def _submit(self):
        t = I18N_ANOMALY.get(self.lang_code, I18N_ANOMALY["vi"])
        pwd = self.entry_pass.get().strip()
        if not pwd:
            self.lbl_err.configure(text=t["pwd_empty"])
            return

        if svc.verify_admin_password(pwd):
            self.destroy()
            if self.on_success:
                self.on_success()
        else:
            self.lbl_err.configure(text=t["pwd_err"])
            self.entry_pass.delete(0, "end")


# ─────────────────────────────────────────────────────────────────────────────
# DIALOG: CHANGE PASSWORD
# ─────────────────────────────────────────────────────────────────────────────
class ChangePasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent, lang_code: str = "vi"):
        super().__init__(parent)
        self.lang_code = lang_code
        t = I18N_ANOMALY.get(lang_code, I18N_ANOMALY["vi"])
        self.title(t["chpwd_title"])
        self.geometry("420x330")
        self.resizable(False, False)
        self.configure(fg_color=BG_CARD)
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(50, (sw - 420) // 2)
            y = max(50, (sh - 330) // 2)
            self.geometry(f"420x330+{x}+{y}")
            self.lift()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

        pad = {"padx": 24}

        ctk.CTkLabel(self, text=t["chpwd_heading"], font=FONT_H1, text_color=TEXT_PRIMARY).pack(pady=(20, 14))

        self.old_pass = ctk.CTkEntry(self, placeholder_text=t["chpwd_old"], show="•", height=36, font=FONT_BODY)
        self.old_pass.pack(fill="x", pady=4, **pad)

        self.new_pass = ctk.CTkEntry(self, placeholder_text=t["chpwd_new"], show="•", height=36, font=FONT_BODY)
        self.new_pass.pack(fill="x", pady=4, **pad)

        self.confirm_pass = ctk.CTkEntry(self, placeholder_text=t["chpwd_confirm"], show="•", height=36, font=FONT_BODY)
        self.confirm_pass.pack(fill="x", pady=4, **pad)

        self.lbl_status = ctk.CTkLabel(self, text="", font=FONT_SMALL, text_color=ACCENT_RED)
        self.lbl_status.pack(pady=6)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", pady=(8, 16), **pad)

        ctk.CTkButton(btn_row, text=t["btn_close"], width=100, height=36, corner_radius=8,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self.destroy).pack(side="left")

        ctk.CTkButton(btn_row, text=t["chpwd_save"], width=140, height=36, corner_radius=8,
                      fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
                      font=FONT_H2, command=self._submit).pack(side="right")

    def _submit(self):
        old_p = self.old_pass.get().strip()
        new_p = self.new_pass.get().strip()
        conf_p = self.confirm_pass.get().strip()

        if not old_p or not new_p:
            self.lbl_status.configure(text="Vui lòng điền đầy đủ các trường!", text_color=ACCENT_RED)
            return
        if new_p != conf_p:
            self.lbl_status.configure(text="Mật khẩu mới và xác nhận không khớp!", text_color=ACCENT_RED)
            return

        ok, msg = svc.change_admin_password(old_p, new_p)
        if ok:
            messagebox.showinfo("Thành công", msg)
            self.destroy()
        else:
            self.lbl_status.configure(text=f"❌ {msg}", text_color=ACCENT_RED)


# ─────────────────────────────────────────────────────────────────────────────
# ADVANCED IMAGE MARKING & ZOOM STUDIO (PHÓNG TO, XOAY, ZOOM, MARKING)
# ─────────────────────────────────────────────────────────────────────────────
class AnomalyImageStudioWindow(ctk.CTkToplevel):
    """
    Dedicated Studio for inspecting defect images with high-resolution pan/zoom,
    90-degree rotations, and complete marking/annotation tools (boxes, ovals, arrows, freehand, text).
    Allows saving the annotated image back to Supabase or downloading to local disk.
    """
    def __init__(self, parent, report_data: dict, initial_pil: Image.Image = None, on_image_updated=None, lang_code: str = "vi"):
        super().__init__(parent)
        self.report_data = report_data
        self.on_image_updated = on_image_updated
        self.image_url = report_data.get("image_url", "")
        self.lang_code = lang_code
        t = I18N_ANOMALY.get(lang_code, I18N_ANOMALY["vi"])

        prod = report_data.get("product_name", "N/A")
        proc = report_data.get("process", "")
        self.title(f"{t['studio_title']} — {prod} ({proc})")
        self.geometry("1300x840")
        self.minsize(980, 640)
        self.configure(fg_color=BG_DEEP)
        self.transient(parent)

        # Image state
        self.raw_image_bytes = None
        self.original_pil: Image.Image = initial_pil
        self.rotation_angle = 0  # 0, 90, 180, 270
        self.zoom_level = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self._tk_img = None
        self._initial_fitted = False

        # Annotation state
        # Annotations are stored in (normalized or image-pixel) coordinates of the ROTATED image
        self.annotations: list[dict] = []  # list of dicts: {"type", "coords" or "points", "color", "width", "text"}
        self.current_tool = "pan"          # "pan", "rect", "oval", "arrow", "pen", "text"
        self.current_color = "#EF4444"     # Red by default
        self.current_width = 3
        self.is_drawing = False
        self.drag_start_screen = (0, 0)
        self.current_temp_item = None
        self.pen_current_points = []

        self._build_ui()
        if self.original_pil:
            self.after(50, self._fit_to_screen)
        else:
            self._load_image_async()

    def _build_ui(self):
        t = I18N_ANOMALY.get(self.lang_code, I18N_ANOMALY["vi"])
        # ── TOP TOOLBAR (2 ROWS) ─────────────────────────────────────────────
        bar = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0,
                           border_width=1, border_color=BORDER_CLR)
        bar.pack(fill="x", side="top")

        # ── ROW 1: MARKING TOOLS & STYLES (CÔNG CỤ ĐÁNH DẤU) ────────────────
        row1 = ctk.CTkFrame(bar, fg_color="transparent", height=42)
        row1.pack(fill="x", padx=12, pady=(6, 2))
        row1.pack_propagate(False)

        # Title & Info
        ctk.CTkLabel(row1, text=t["studio_tools_title"], font=FONT_H1, text_color=ACCENT_TEAL).pack(side="left", padx=(2, 10))

        # Tool selector buttons
        self.tool_buttons = {}
        tools = [
            ("pan", t["studio_pan"]),
            ("rect", t["studio_rect"]),
            ("oval", t["studio_oval"]),
            ("arrow", t["studio_arrow"]),
            ("pen", t["studio_pen"]),
            ("text", t["studio_text"]),
        ]

        for t_key, t_label in tools:
            btn = ctk.CTkButton(
                row1, text=t_label, width=82, height=30, corner_radius=6,
                font=FONT_SMALL,
                fg_color=ACCENT_TEAL if t_key == self.current_tool else BG_SURFACE,
                text_color=("#FFFFFF", "#0B0F1A") if t_key == self.current_tool else TEXT_PRIMARY,
                hover_color=BG_HOVER,
                command=lambda k=t_key: self._select_tool(k)
            )
            btn.pack(side="left", padx=2)
            self.tool_buttons[t_key] = btn

        # Separator
        ctk.CTkFrame(row1, fg_color=BORDER_CLR, width=1, height=24).pack(side="left", padx=8)

        # Color picker buttons
        self.color_buttons = {}
        colors = [
            ("#EF4444", "Đỏ"),
            ("#F59E0B", "Vàng"),
            ("#10B981", "Xanh"),
            ("#3B82F6", "Lam"),
            ("#A855F7", "Tím"),
            ("#FFFFFF", "Trắng"),
        ]
        for c_hex, c_name in colors:
            btn = ctk.CTkButton(
                row1, text="", width=24, height=24, corner_radius=12,
                fg_color=c_hex, hover_color=c_hex,
                border_width=2 if c_hex == self.current_color else 0,
                border_color="#FFFFFF" if c_hex != "#FFFFFF" else "#0F172A",
                command=lambda c=c_hex: self._select_color(c)
            )
            btn.pack(side="left", padx=2)
            self.color_buttons[c_hex] = btn

        # Stroke width selector
        ctk.CTkFrame(row1, fg_color=BORDER_CLR, width=1, height=24).pack(side="left", padx=6)
        self.width_buttons = {}
        for w_val, w_label in [(2, "2px"), (4, "4px"), (6, "6px")]:
            w_btn = ctk.CTkButton(
                row1, text=w_label, width=38, height=28, corner_radius=6,
                font=FONT_SMALL,
                fg_color=ACCENT_TEAL if w_val == self.current_width else BG_SURFACE,
                text_color=("#FFFFFF", "#0B0F1A") if w_val == self.current_width else TEXT_PRIMARY,
                hover_color=BG_HOVER,
                command=lambda w=w_val: self._select_width(w)
            )
            w_btn.pack(side="left", padx=1)
            self.width_buttons[w_val] = w_btn

        # Right side of Row 1: Undo and Clear
        ctk.CTkButton(row1, text=t["studio_clear"], width=95, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
                      command=self._clear_annotations).pack(side="right", padx=(2, 0))

        ctk.CTkButton(row1, text=t["studio_undo"], width=85, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._undo).pack(side="right", padx=2)

        # Subtle horizontal divider between Row 1 and Row 2
        ctk.CTkFrame(bar, fg_color=BORDER_CLR, height=1).pack(fill="x", padx=8, pady=2)

        # ── ROW 2: VIEW, FULLSCREEN, ZOOM & ACTIONS (ĐIỀU KHIỂN XEM & LƯU) ───
        row2 = ctk.CTkFrame(bar, fg_color="transparent", height=42)
        row2.pack(fill="x", padx=12, pady=(2, 6))
        row2.pack_propagate(False)

        # Ô PHÓNG TO TOÀN MÀN HÌNH (FULLSCREEN / MAXIMIZE TOGGLE)
        self.btn_fullscreen = ctk.CTkButton(
            row2, text=t["studio_fullscreen"], width=140, height=30, font=FONT_H2,
            fg_color=BG_SURFACE, text_color=ACCENT_TEAL, hover_color=BG_HOVER,
            border_width=1, border_color=ACCENT_TEAL,
            command=self._toggle_fullscreen
        )
        self.btn_fullscreen.pack(side="left", padx=(2, 8))

        # Rotate Button
        ctk.CTkButton(row2, text=t["studio_rotate"], width=105, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._rotate_90).pack(side="left", padx=2)

        # Separator
        ctk.CTkFrame(row2, fg_color=BORDER_CLR, width=1, height=24).pack(side="left", padx=6)

        # Zoom Buttons
        ctk.CTkButton(row2, text="🔍+", width=36, height=30, font=FONT_BOLD,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=lambda: self._zoom_by_factor(1.2)).pack(side="left", padx=1)
        ctk.CTkButton(row2, text="🔍-", width=36, height=30, font=FONT_BOLD,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=lambda: self._zoom_by_factor(0.8)).pack(side="left", padx=1)
        ctk.CTkButton(row2, text="🎯 100%", width=60, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._zoom_to_100).pack(side="left", padx=2)
        ctk.CTkButton(row2, text=t["studio_fit"], width=85, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._fit_to_screen).pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(row2, text="100%", font=FONT_BOLD, text_color=ACCENT_TEAL, width=48)
        self.lbl_zoom.pack(side="left", padx=4)

        # Right side of Row 2: Close, Download, Save Cloud
        ctk.CTkButton(
            row2, text=t["studio_close"], width=75, height=30, font=FONT_SMALL,
            fg_color=BG_SURFACE, text_color=TEXT_MUTED, hover_color=BG_HOVER,
            command=self.destroy
        ).pack(side="right", padx=(2, 0))

        self.btn_download = ctk.CTkButton(
            row2, text=t["studio_download"], width=110, height=30, font=FONT_SMALL,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            command=self._download_to_pc
        )
        self.btn_download.pack(side="right", padx=4)

        self.btn_save_cloud = ctk.CTkButton(
            row2, text=t["studio_save_cloud"], width=155, height=30, font=FONT_H2,
            fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
            hover_color=("#0F766E", "#00A88C"),
            command=self._save_to_cloud
        )
        self.btn_save_cloud.pack(side="right", padx=4)

        # ── INTERACTIVE CANVAS WORKSPACE ──────────────────────────────────────
        self.canvas_frame = ctk.CTkFrame(self, fg_color="black", corner_radius=0)
        self.canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg="#0B0F1A", highlightthickness=0, cursor="hand2")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mouse & Gesture bindings
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)

        # Right click drag to pan anytime
        self.canvas.bind("<ButtonPress-3>", self._start_right_pan)
        self.canvas.bind("<B3-Motion>", self._do_right_pan)

        # Wheel zoom
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        # Hotkeys
        self.bind("<Control-z>", lambda e: self._undo())
        self.bind("<Control-Z>", lambda e: self._undo())
        self.bind("<r>", lambda e: self._rotate_90())
        self.bind("<R>", lambda e: self._rotate_90())
        self.bind("<F11>", lambda e: self._toggle_fullscreen())
        self.bind("<Escape>", lambda e: self._exit_fullscreen_if_active())

    def _toggle_fullscreen(self):
        """Toggles window between Maximized (Full Screen) and Normal state, and auto-fits image."""
        t = I18N_ANOMALY.get(self.lang_code, I18N_ANOMALY["vi"])
        try:
            is_zoomed = self.state() == "zoomed" or bool(self.attributes("-fullscreen"))
            if is_zoomed:
                try:
                    self.attributes("-fullscreen", False)
                except Exception:
                    pass
                self.state("normal")
                self.btn_fullscreen.configure(text=t["studio_fullscreen"], fg_color=BG_SURFACE)
            else:
                self.state("zoomed")
                self.btn_fullscreen.configure(text=t["studio_unfullscreen"], fg_color=ACCENT_TEAL)

            self.after(150, self._fit_to_screen)
        except Exception as e:
            print(f"[Studio] Toggle fullscreen error: {e}")

    def _exit_fullscreen_if_active(self):
        t = I18N_ANOMALY.get(self.lang_code, I18N_ANOMALY["vi"])
        try:
            if self.state() == "zoomed" or bool(self.attributes("-fullscreen")):
                try:
                    self.attributes("-fullscreen", False)
                except Exception:
                    pass
                self.state("normal")
                self.btn_fullscreen.configure(text=t["studio_fullscreen"], fg_color=BG_SURFACE)
                self.after(150, self._fit_to_screen)
        except Exception:
            pass

    def _zoom_to_100(self):
        """Resets zoom level to 100% (1:1 pixel ratio) and centers image."""
        base = self._get_current_base_pil()
        if not base:
            return
        cw = max(self.canvas.winfo_width(), 600)
        ch = max(self.canvas.winfo_height(), 500)
        iw, ih = base.size
        self.zoom_level = 1.0
        self.pan_x = (cw - iw) // 2
        self.pan_y = (ch - ih) // 2
        self._render()

    def _on_canvas_configure(self, event):
        if not self._initial_fitted and self.original_pil:
            if event.width > 50 and event.height > 50:
                self._initial_fitted = True
                self._fit_to_screen()

    # ── IMAGE LOADING ────────────────────────────────────────────────────────
    def _load_image_async(self):
        if not self.image_url:
            self.canvas.delete("all")
            self.canvas.create_text(
                max(self.canvas.winfo_width() // 2, 300),
                max(self.canvas.winfo_height() // 2, 250),
                text="[Báo cáo này chưa có hình ảnh đính kèm]",
                fill="#7A8BA6", font=("Segoe UI", 13)
            )
            return

        self.canvas.delete("all")
        self.canvas.create_text(
            max(self.canvas.winfo_width() // 2, 300),
            max(self.canvas.winfo_height() // 2, 250),
            text="⏳ Đang tải ảnh chất lượng cao từ Cloud...", fill="#00C9A7",
            font=("Segoe UI", 15, "bold"), tags="loading_msg"
        )

        def _fetch():
            try:
                import urllib.request
                req = urllib.request.Request(self.image_url, headers={"User-Agent": "Mozilla/5.0"})
                ssl_ctx = svc.get_ssl_context()
                with urllib.request.urlopen(req, timeout=20, context=ssl_ctx) as resp:
                    raw = resp.read()
                self.raw_image_bytes = raw
                pil_img = Image.open(io.BytesIO(raw))
                self.original_pil = pil_img
                def _done():
                    self.canvas.delete("loading_msg")
                    self._fit_to_screen()
                safe_after(self, 0, _done)
            except Exception as e:
                def _err():
                    self.canvas.delete("all")
                    self.canvas.create_text(
                        max(self.canvas.winfo_width() // 2, 300),
                        max(self.canvas.winfo_height() // 2, 250),
                        text=f"❌ Không thể tải ảnh: {e}", fill="#EF4444", font=("Segoe UI", 12, "bold")
                    )
                safe_after(self, 0, _err)

        threading.Thread(target=_fetch, daemon=True).start()

    def _get_current_base_pil(self) -> Image.Image:
        """Returns the original PIL image rotated by current rotation angle."""
        if not self.original_pil:
            return None
        if self.rotation_angle == 90:
            return self.original_pil.rotate(-90, expand=True)
        elif self.rotation_angle == 180:
            return self.original_pil.rotate(180, expand=True)
        elif self.rotation_angle == 270:
            return self.original_pil.rotate(-270, expand=True)
        return self.original_pil

    def _fit_to_screen(self):
        base = self._get_current_base_pil()
        if not base:
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 50 or ch <= 50:
            self.after(50, self._fit_to_screen)
            return
        iw, ih = base.size
        scale = min((cw - 40) / iw, (ch - 40) / ih, 1.0)
        self.zoom_level = max(scale, 0.05)
        self.pan_x = (cw - int(iw * self.zoom_level)) // 2
        self.pan_y = (ch - int(ih * self.zoom_level)) // 2
        self._render()

    def _zoom_by_factor(self, factor: float):
        base = self._get_current_base_pil()
        if not base:
            return
        cw = max(self.canvas.winfo_width(), 600)
        ch = max(self.canvas.winfo_height(), 500)

        # Center zoom
        center_x = cw / 2
        center_y = ch / 2
        img_cx = (center_x - self.pan_x) / self.zoom_level
        img_cy = (center_y - self.pan_y) / self.zoom_level

        new_zoom = max(0.05, min(8.0, self.zoom_level * factor))
        self.zoom_level = new_zoom
        self.pan_x = center_x - img_cx * new_zoom
        self.pan_y = center_y - img_cy * new_zoom
        self._render()

    def _on_mousewheel(self, event):
        factor = 1.15 if event.delta > 0 else 0.85
        base = self._get_current_base_pil()
        if not base:
            return
        cursor_x = event.x
        cursor_y = event.y
        img_x = (cursor_x - self.pan_x) / self.zoom_level
        img_y = (cursor_y - self.pan_y) / self.zoom_level

        new_zoom = max(0.05, min(8.0, self.zoom_level * factor))
        self.zoom_level = new_zoom
        self.pan_x = cursor_x - img_x * new_zoom
        self.pan_y = cursor_y - img_y * new_zoom
        self._render()

    def _rotate_90(self):
        base = self._get_current_base_pil()
        if not base:
            return
        old_w, old_h = base.size

        # Transform all annotations for clockwise 90 rotation: (x, y) -> (old_h - y, x)
        for ann in self.annotations:
            t = ann.get("type")
            if t in ("rect", "oval"):
                x1, y1, x2, y2 = ann["coords"]
                nx1, ny1 = old_h - y1, x1
                nx2, ny2 = old_h - y2, x2
                ann["coords"] = (min(nx1, nx2), min(ny1, ny2), max(nx1, nx2), max(ny1, ny2))
            elif t == "arrow":
                x1, y1, x2, y2 = ann["coords"]
                ann["coords"] = (old_h - y1, x1, old_h - y2, x2)
            elif t == "pen":
                ann["points"] = [(old_h - y, x) for (x, y) in ann.get("points", [])]
            elif t == "text":
                x, y = ann["coords"]
                ann["coords"] = (old_h - y, x)

        self.rotation_angle = (self.rotation_angle + 90) % 360
        self._fit_to_screen()

    def _select_width(self, width: int):
        self.current_width = width
        for w, btn in self.width_buttons.items():
            if w == width:
                btn.configure(fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"))
            else:
                btn.configure(fg_color=BG_SURFACE, text_color=TEXT_PRIMARY)

    # ── TOOL & COLOR SELECTION ───────────────────────────────────────────────
    def _select_tool(self, tool_key: str):
        self.current_tool = tool_key
        for k, btn in self.tool_buttons.items():
            if k == tool_key:
                btn.configure(fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"))
            else:
                btn.configure(fg_color=BG_SURFACE, text_color=TEXT_PRIMARY)

        # Change cursor accordingly
        if tool_key == "pan":
            self.canvas.configure(cursor="hand2")
        elif tool_key == "text":
            self.canvas.configure(cursor="xterm")
        else:
            self.canvas.configure(cursor="crosshair")

    def _select_color(self, c_hex: str):
        self.current_color = c_hex
        for h, btn in self.color_buttons.items():
            btn.configure(border_width=2 if h == c_hex else 0)

    # ── ANNOTATION RENDERING ─────────────────────────────────────────────────
    def _render(self):
        base = self._get_current_base_pil()
        if not base:
            return

        iw, ih = base.size
        nw = max(10, int(iw * self.zoom_level))
        nh = max(10, int(ih * self.zoom_level))

        resized = base.resize((nw, nh), Image.Resampling.BILINEAR)
        self._tk_img = ImageTk.PhotoImage(resized, master=self.canvas)

        self.canvas.delete("all")
        # Draw background image
        self.canvas.create_image(self.pan_x, self.pan_y, anchor="nw", image=self._tk_img, tags="bg_img")

        # Draw all saved annotations
        for ann in self.annotations:
            self._draw_annotation_on_canvas(ann)

        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")

    def _draw_annotation_on_canvas(self, ann: dict):
        t = ann.get("type")
        col = ann.get("color", "#EF4444")
        w = max(1, int(ann.get("width", 3) * self.zoom_level))

        if t in ("rect", "oval", "arrow"):
            ix1, iy1, ix2, iy2 = ann["coords"]
            sx1 = self.pan_x + ix1 * self.zoom_level
            sy1 = self.pan_y + iy1 * self.zoom_level
            sx2 = self.pan_x + ix2 * self.zoom_level
            sy2 = self.pan_y + iy2 * self.zoom_level

            if t == "rect":
                self.canvas.create_rectangle(sx1, sy1, sx2, sy2, outline=col, width=w, tags="ann")
            elif t == "oval":
                self.canvas.create_oval(sx1, sy1, sx2, sy2, outline=col, width=w, tags="ann")
            elif t == "arrow":
                arrow_sz = max(10, int(14 * self.zoom_level))
                self.canvas.create_line(sx1, sy1, sx2, sy2, fill=col, width=w, arrow=tk.LAST,
                                        arrowshape=(arrow_sz, arrow_sz + 4, max(4, arrow_sz // 2)), tags="ann")

        elif t == "pen":
            screen_pts = []
            for (ix, iy) in ann.get("points", []):
                screen_pts.extend([self.pan_x + ix * self.zoom_level, self.pan_y + iy * self.zoom_level])
            if len(screen_pts) >= 4:
                self.canvas.create_line(*screen_pts, fill=col, width=w, smooth=True, tags="ann")

        elif t == "text":
            ix, iy = ann["coords"]
            sx = self.pan_x + ix * self.zoom_level
            sy = self.pan_y + iy * self.zoom_level
            txt = ann.get("text", "")
            f_sz = max(10, int(13 * self.zoom_level))
            # Text background badge for clarity
            self.canvas.create_rectangle(sx - 2, sy - 2, sx + len(txt) * (f_sz * 0.65) + 6, sy + f_sz + 6,
                                         fill="#000000", outline=col, width=1, tags="ann")
            self.canvas.create_text(sx + 2, sy + 2, text=txt, fill=col, font=("Segoe UI", f_sz, "bold"),
                                    anchor="nw", tags="ann")

    # ── MOUSE DRAWING INTERACTIONS ───────────────────────────────────────────
    def _on_canvas_press(self, event):
        base = self._get_current_base_pil()
        if not base:
            return

        if self.current_tool == "pan":
            self._pan_start_x = event.x
            self._pan_start_y = event.y
            return

        # Convert screen to image coords
        ix = (event.x - self.pan_x) / self.zoom_level
        iy = (event.y - self.pan_y) / self.zoom_level
        self.drag_start_screen = (event.x, event.y)
        self.drag_start_img = (ix, iy)
        self.is_drawing = True

        if self.current_tool == "pen":
            self.pen_current_points = [(ix, iy)]
        elif self.current_tool == "text":
            self.is_drawing = False
            txt = simpledialog.askstring("Ghi Chú Trên Ảnh", "Nhập nội dung ghi chú:", parent=self)
            if txt and txt.strip():
                self.annotations.append({
                    "type": "text",
                    "coords": (ix, iy),
                    "text": txt.strip(),
                    "color": self.current_color,
                    "width": self.current_width
                })
                self._render()

    def _on_canvas_drag(self, event):
        if self.current_tool == "pan":
            dx = event.x - self._pan_start_x
            dy = event.y - self._pan_start_y
            self._pan_start_x = event.x
            self._pan_start_y = event.y
            self.pan_x += dx
            self.pan_y += dy
            self._render()
            return

        if not self.is_drawing:
            return

        # Live preview of current shape
        self.canvas.delete("temp_shape")
        sx1, sy1 = self.drag_start_screen
        sx2, sy2 = event.x, event.y
        col = self.current_color
        w = self.current_width

        if self.current_tool == "rect":
            self.canvas.create_rectangle(sx1, sy1, sx2, sy2, outline=col, width=w, tags="temp_shape")
        elif self.current_tool == "oval":
            self.canvas.create_oval(sx1, sy1, sx2, sy2, outline=col, width=w, tags="temp_shape")
        elif self.current_tool == "arrow":
            self.canvas.create_line(sx1, sy1, sx2, sy2, fill=col, width=w, arrow=tk.LAST, tags="temp_shape")
        elif self.current_tool == "pen":
            ix = (event.x - self.pan_x) / self.zoom_level
            iy = (event.y - self.pan_y) / self.zoom_level
            self.pen_current_points.append((ix, iy))
            screen_pts = []
            for (px, py) in self.pen_current_points:
                screen_pts.extend([self.pan_x + px * self.zoom_level, self.pan_y + py * self.zoom_level])
            if len(screen_pts) >= 4:
                self.canvas.create_line(*screen_pts, fill=col, width=w, tags="temp_shape")

    def _on_canvas_release(self, event):
        if not self.is_drawing:
            return
        self.is_drawing = False
        self.canvas.delete("temp_shape")

        ix1, iy1 = self.drag_start_img
        ix2 = (event.x - self.pan_x) / self.zoom_level
        iy2 = (event.y - self.pan_y) / self.zoom_level

        # Ignore tiny accidental clicks
        dist = math.hypot(event.x - self.drag_start_screen[0], event.y - self.drag_start_screen[1])
        if dist < 6 and self.current_tool != "pen":
            return

        if self.current_tool in ("rect", "oval", "arrow"):
            self.annotations.append({
                "type": self.current_tool,
                "coords": (ix1, iy1, ix2, iy2),
                "color": self.current_color,
                "width": self.current_width
            })
        elif self.current_tool == "pen" and len(self.pen_current_points) > 1:
            self.annotations.append({
                "type": "pen",
                "points": list(self.pen_current_points),
                "color": self.current_color,
                "width": self.current_width
            })
            self.pen_current_points = []

        self._render()

    def _start_right_pan(self, event):
        self._r_pan_start_x = event.x
        self._r_pan_start_y = event.y

    def _do_right_pan(self, event):
        dx = event.x - self._r_pan_start_x
        dy = event.y - self._r_pan_start_y
        self._r_pan_start_x = event.x
        self._r_pan_start_y = event.y
        self.pan_x += dx
        self.pan_y += dy
        self._render()

    def _undo(self):
        if self.annotations:
            self.annotations.pop()
            self._render()

    def _clear_annotations(self):
        if not self.annotations:
            return
        if messagebox.askyesno("Xóa nét vẽ", "Bạn có chắc chắn muốn xóa toàn bộ các nét đánh dấu trên ảnh?"):
            self.annotations.clear()
            self._render()

    # ── EXPORTING / BURNING ANNOTATIONS ON PIL ────────────────────────────────
    def _render_flattened_pil(self) -> Image.Image:
        """Draws all annotations onto the full-resolution rotated PIL image."""
        base = self._get_current_base_pil()
        if not base:
            return None
        out = base.copy().convert("RGB")
        draw = ImageDraw.Draw(out)

        for ann in self.annotations:
            t = ann.get("type")
            col = ann.get("color", "#EF4444")
            w = max(2, int(ann.get("width", 3)))

            if t == "rect":
                x1, y1, x2, y2 = ann["coords"]
                draw.rectangle((min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)), outline=col, width=w)
            elif t == "oval":
                x1, y1, x2, y2 = ann["coords"]
                draw.ellipse((min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)), outline=col, width=w)
            elif t == "arrow":
                x1, y1, x2, y2 = ann["coords"]
                draw.line([(x1, y1), (x2, y2)], fill=col, width=w)
                dx = x2 - x1
                dy = y2 - y1
                length = math.hypot(dx, dy)
                if length > 0:
                    ux, uy = dx / length, dy / length
                    head_len = min(32, max(14, w * 5))
                    px, py = -uy, ux
                    p1 = (x2 - ux * head_len + px * (head_len * 0.5), y2 - uy * head_len + py * (head_len * 0.5))
                    p2 = (x2 - ux * head_len - px * (head_len * 0.5), y2 - uy * head_len - py * (head_len * 0.5))
                    draw.polygon([(x2, y2), p1, p2], fill=col)
            elif t == "pen":
                pts = ann.get("points", [])
                if len(pts) > 1:
                    draw.line(pts, fill=col, width=w, joint="curve")
            elif t == "text":
                x, y = ann["coords"]
                txt = ann.get("text", "")
                try:
                    fnt = ImageFont.load_default()
                except Exception:
                    fnt = None
                # Black badge background
                bb = (x - 2, y - 2, x + len(txt) * 8 + 6, y + 18)
                draw.rectangle(bb, fill="#000000", outline=col, width=1)
                draw.text((x + 2, y + 2), txt, fill=col, font=fnt)

        return out

    def _download_to_pc(self):
        flat = self._render_flattened_pil()
        if not flat:
            return

        prod = self.report_data.get("product_name", "defect")
        proc = self.report_data.get("process", "QC")
        def_fname = f"Marking_{prod}_{proc}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        f = filedialog.asksaveasfilename(
            title="Lưu Ảnh Đã Đánh Dấu",
            initialfile=def_fname,
            defaultextension=".jpg",
            filetypes=[("JPEG Image", "*.jpg"), ("PNG Image", "*.png")]
        )
        if f:
            try:
                flat.save(f, quality=92)
                messagebox.showinfo("Thành công", f"Đã lưu ảnh đánh dấu về máy:\n{f}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu file ảnh:\n{e}")

    def _save_to_cloud(self):
        flat = self._render_flattened_pil()
        if not flat:
            return

        if not messagebox.askyesno("Lưu Lên Cloud", "Bạn có muốn lưu đè ảnh đã đánh dấu này lên Báo Cáo trên Supabase Cloud?"):
            return

        self.btn_save_cloud.configure(state="disabled", text="⏳ Đang lưu...")

        def _do_upload():
            try:
                buf = io.BytesIO()
                flat.save(buf, format="JPEG", quality=88, optimize=True)
                new_bytes = buf.getvalue()

                rep_id = self.report_data["id"]
                # Update report with new image bytes
                updated = svc.update_report(rep_id, self.report_data, new_image_bytes=new_bytes, new_image_name="annotated_defect.jpg")
                safe_after(self, 0, lambda: self._on_cloud_saved_success(updated))
            except Exception as e:
                safe_after(self, 0, lambda: self._on_cloud_saved_error(str(e)))

        threading.Thread(target=_do_upload, daemon=True).start()

    def _on_cloud_saved_success(self, updated_rep):
        self.btn_save_cloud.configure(state="normal", text="💾 Lưu Lên Báo Cáo")
        messagebox.showinfo("Thành công", "✅ Đã cập nhật ảnh đánh dấu lên Cloud Supabase thành công!")
        if self.on_image_updated:
            self.on_image_updated(updated_rep)
        self.destroy()

    def _on_cloud_saved_error(self, err_msg):
        self.btn_save_cloud.configure(state="normal", text="💾 Lưu Lên Báo Cáo")
        messagebox.showerror("Lỗi Lưu Cloud", f"Không thể lưu ảnh lên Cloud:\n{err_msg}")


# ─────────────────────────────────────────────────────────────────────────────
# DRAWER: CREATE / EDIT ANOMALY REPORT (SLIDE-IN PANEL WITH IMAGE PREVIEW)
# ─────────────────────────────────────────────────────────────────────────────
class AnomalyDrawerFrame(ctk.CTkFrame):
    """
    Slide-in Drawer panel on the right side of the main view for creating and editing reports.
    Features rich form controls, auto-calculated defect rate, and a dedicated image preview card.
    """
    def __init__(self, parent, main_view, on_saved=None):
        super().__init__(parent, fg_color=BG_CARD, corner_radius=12,
                         border_width=1, border_color=BORDER_CLR, width=500)
        self.parent = parent
        self.main_view = main_view
        self.on_saved = on_saved
        self.current_lang = getattr(self.main_view, "current_lang", "vi")

        self.is_edit = False
        self.report_data = {}
        self.selected_image_bytes = None
        self.selected_image_name = None
        self.current_pil_image: Image.Image = None
        self._ctk_preview = None

        self._build_ui()

    def _build_ui(self):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.pack_propagate(False)

        # ── DRAWER HEADER ──
        header = ctk.CTkFrame(self, fg_color=BG_SURFACE, height=52, corner_radius=10)
        header.pack(fill="x", padx=10, pady=(10, 6))
        header.pack_propagate(False)

        self.lbl_title = ctk.CTkLabel(header, text=t["drawer_title_create"], font=FONT_H1, text_color=TEXT_PRIMARY)
        self.lbl_title.pack(side="left", padx=16)

        btn_close = ctk.CTkButton(header, text="✕", width=34, height=34, corner_radius=6,
                                  fg_color="transparent", text_color=TEXT_MUTED, hover_color=BG_HOVER,
                                  font=("Segoe UI", 13, "bold"), command=self.close_drawer)
        btn_close.pack(side="right", padx=10)

        # ── SCROLLABLE BODY ──
        self.body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=10, pady=4)

        pad = {"padx": 14, "pady": (3, 3)}

        # --- SECTION 1: THÔNG TIN CHUNG & SỐ LƯỢNG ---
        c1 = ctk.CTkFrame(self.body, fg_color=BG_DEEP, corner_radius=8, border_width=1, border_color=BORDER_CLR)
        c1.pack(fill="x", pady=(0, 10))

        self.lbl_sec1 = ctk.CTkLabel(c1, text=t["sec_inspection"], font=FONT_H2, text_color=ACCENT_TEAL)
        self.lbl_sec1.pack(anchor="w", padx=14, pady=(10, 4))

        self.lbl_f_date = ctk.CTkLabel(c1, text=t["lbl_date"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_date.pack(anchor="w", **pad)
        self.ent_date = ctk.CTkEntry(c1, height=34, font=FONT_BODY)
        self.ent_date.pack(fill="x", **pad)

        self.lbl_f_proc = ctk.CTkLabel(c1, text=t["lbl_process"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_proc.pack(anchor="w", **pad)
        self.opt_process = ctk.CTkOptionMenu(c1, values=["CHA", "SCP", "RAD", "CHP", "SMT", "Lắp Ráp", "KCS / Final QC", "Khác"],
                                             height=34, font=FONT_BODY)
        self.opt_process.pack(fill="x", **pad)

        self.lbl_f_prod = ctk.CTkLabel(c1, text=t["lbl_product"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_prod.pack(anchor="w", **pad)
        self.ent_product = ctk.CTkEntry(c1, height=34, font=FONT_BODY, placeholder_text=t["ph_product"])
        self.ent_product.pack(fill="x", **pad)

        self.lbl_f_mach = ctk.CTkLabel(c1, text=t["lbl_machine"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_mach.pack(anchor="w", **pad)
        self.ent_machine = ctk.CTkEntry(c1, height=34, font=FONT_BODY, placeholder_text=t["ph_machine"])
        self.ent_machine.pack(fill="x", **pad)

        # Quantities row
        q_row = ctk.CTkFrame(c1, fg_color="transparent")
        q_row.pack(fill="x", padx=14, pady=4)

        f1 = ctk.CTkFrame(q_row, fg_color="transparent")
        f1.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.lbl_f_tot = ctk.CTkLabel(f1, text=t["lbl_tot"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_tot.pack(anchor="w")
        self.ent_tot = ctk.CTkEntry(f1, height=34, font=FONT_BODY)
        self.ent_tot.pack(fill="x")
        self.ent_tot.bind("<KeyRelease>", self._calc_rate)

        f2 = ctk.CTkFrame(q_row, fg_color="transparent")
        f2.pack(side="left", fill="x", expand=True, padx=4)
        self.lbl_f_def = ctk.CTkLabel(f2, text=t["lbl_def"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_def.pack(anchor="w")
        self.ent_def = ctk.CTkEntry(f2, height=34, font=FONT_BODY)
        self.ent_def.pack(fill="x")
        self.ent_def.bind("<KeyRelease>", self._calc_rate)

        f3 = ctk.CTkFrame(q_row, fg_color="transparent")
        f3.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.lbl_f_rate = ctk.CTkLabel(f3, text=t["lbl_rate"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_rate.pack(anchor="w")
        self.lbl_rate = ctk.CTkLabel(f3, text="0.00%", height=34, font=FONT_H2,
                                     fg_color=BG_SURFACE, corner_radius=6, text_color=ACCENT_AMBER)
        self.lbl_rate.pack(fill="x")

        self.lbl_f_resp = ctk.CTkLabel(c1, text=t["lbl_resp"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_resp.pack(anchor="w", **pad)
        self.ent_resp = ctk.CTkEntry(c1, height=34, font=FONT_BODY)
        self.ent_resp.pack(fill="x", **pad)

        self.lbl_f_pic = ctk.CTkLabel(c1, text=t["lbl_pic"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_pic.pack(anchor="w", **pad)
        self.ent_pic = ctk.CTkEntry(c1, height=34, font=FONT_BODY)
        self.ent_pic.pack(fill="x", **pad)

        self.lbl_f_prog = ctk.CTkLabel(c1, text=t["lbl_prog"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_prog.pack(anchor="w", **pad)
        self.opt_prog = ctk.CTkOptionMenu(c1, values=t["drawer_prog_options"],
                                          height=34, font=FONT_BODY)
        self.opt_prog.pack(fill="x", padx=14, pady=(3, 12))

        # --- SECTION 2: HÌNH ẢNH MINH HỌA & ĐÁNH DẤU LỖI ---
        c2 = ctk.CTkFrame(self.body, fg_color=BG_DEEP, corner_radius=8, border_width=1, border_color=BORDER_CLR)
        c2.pack(fill="x", pady=(0, 10))

        self.lbl_sec2 = ctk.CTkLabel(c2, text=t["sec_image"], font=FONT_H2, text_color=ACCENT_TEAL)
        self.lbl_sec2.pack(anchor="w", padx=14, pady=(10, 4))

        # Image preview box
        self.box_img = ctk.CTkFrame(c2, fg_color=BG_SURFACE, corner_radius=8, height=190)
        self.box_img.pack(fill="x", padx=14, pady=(4, 6))
        self.box_img.pack_propagate(False)

        self.lbl_img_preview = ctk.CTkLabel(self.box_img, text=t["img_empty_hint"],
                                            font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_img_preview.pack(fill="both", expand=True, padx=8, pady=8)

        # Image control buttons
        btn_img_bar = ctk.CTkFrame(c2, fg_color="transparent")
        btn_img_bar.pack(fill="x", padx=14, pady=(2, 10))

        self.btn_studio = ctk.CTkButton(btn_img_bar, text=t["btn_studio"], height=32, corner_radius=6,
                                        font=FONT_SMALL, fg_color=ACCENT_BLUE, text_color="#FFFFFF",
                                        command=self._open_studio, state="disabled")
        self.btn_studio.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_pick = ctk.CTkButton(btn_img_bar, text=t["btn_pick"], width=68, height=32, corner_radius=6,
                                      font=FONT_SMALL, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                                      command=self._pick_image)
        self.btn_pick.pack(side="left", padx=2)

        self.btn_paste = ctk.CTkButton(btn_img_bar, text=t["btn_paste"], width=65, height=32, corner_radius=6,
                                       font=FONT_SMALL, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                                       command=self._paste_image)
        self.btn_paste.pack(side="left", padx=2)

        self.btn_clear_img = ctk.CTkButton(btn_img_bar, text="✕", width=36, height=32, corner_radius=6,
                                           font=FONT_SMALL, fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
                                           command=self._clear_image)
        self.btn_clear_img.pack(side="left", padx=(2, 0))

        # --- SECTION 3: MÔ TẢ & ĐỐI SÁCH KHẮC PHỤC ---
        c3 = ctk.CTkFrame(self.body, fg_color=BG_DEEP, corner_radius=8, border_width=1, border_color=BORDER_CLR)
        c3.pack(fill="x", pady=(0, 10))

        self.lbl_sec3 = ctk.CTkLabel(c3, text=t["sec_defect"], font=FONT_H2, text_color=ACCENT_TEAL)
        self.lbl_sec3.pack(anchor="w", padx=14, pady=(10, 4))

        self.lbl_f_desc = ctk.CTkLabel(c3, text=t["lbl_desc"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_desc.pack(anchor="w", **pad)
        self.txt_desc = ctk.CTkTextbox(c3, height=65, font=FONT_BODY)
        self.txt_desc.pack(fill="x", **pad)

        self.lbl_f_cause = ctk.CTkLabel(c3, text=t["lbl_cause"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_cause.pack(anchor="w", **pad)
        self.txt_cause = ctk.CTkTextbox(c3, height=65, font=FONT_BODY)
        self.txt_cause.pack(fill="x", **pad)

        self.lbl_f_counter = ctk.CTkLabel(c3, text=t["lbl_counter"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_counter.pack(anchor="w", **pad)
        self.txt_counter = ctk.CTkTextbox(c3, height=65, font=FONT_BODY)
        self.txt_counter.pack(fill="x", **pad)

        self.lbl_f_sop = ctk.CTkLabel(c3, text=t["lbl_sop"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_sop.pack(anchor="w", **pad)
        self.ent_sop = ctk.CTkEntry(c3, height=34, font=FONT_BODY)
        self.ent_sop.pack(fill="x", **pad)

        self.lbl_f_notes = ctk.CTkLabel(c3, text=t["lbl_notes"], font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_f_notes.pack(anchor="w", **pad)
        self.ent_notes = ctk.CTkEntry(c3, height=34, font=FONT_BODY)
        self.ent_notes.pack(fill="x", padx=14, pady=(3, 12))

        # ── DRAWER FOOTER ──
        footer = ctk.CTkFrame(self, fg_color=BG_SURFACE, height=54, corner_radius=10)
        footer.pack(fill="x", side="bottom", padx=10, pady=(4, 10))
        footer.pack_propagate(False)

        self.lbl_saving = ctk.CTkLabel(footer, text="", font=FONT_SMALL, text_color=ACCENT_TEAL)
        self.lbl_saving.pack(side="left", padx=12)

        self.btn_drawer_close = ctk.CTkButton(footer, text=t["btn_close"], width=80, height=36, corner_radius=8,
                                             fg_color=BG_CARD, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                                             command=self.close_drawer)
        self.btn_drawer_close.pack(side="right", padx=(4, 10))

        self.btn_save = ctk.CTkButton(footer, text=t["btn_save"], width=140, height=36, corner_radius=8,
                                      fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
                                      font=FONT_H2, command=self._save_report)
        self.btn_save.pack(side="right", padx=4)

    def update_language(self, lang_code: str):
        """Updates all visible texts in the Drawer according to selected language."""
        self.current_lang = lang_code
        t = I18N_ANOMALY.get(lang_code, I18N_ANOMALY["vi"])

        if self.is_edit:
            rep_id = self.report_data.get("id", "")
            self.lbl_title.configure(text=f"{t['drawer_title_edit']} #{rep_id}")
        else:
            self.lbl_title.configure(text=t["drawer_title_create"])

        self.lbl_sec1.configure(text=t["sec_inspection"])
        self.lbl_f_date.configure(text=t["lbl_date"])
        self.lbl_f_proc.configure(text=t["lbl_process"])
        self.lbl_f_prod.configure(text=t["lbl_product"])
        self.ent_product.configure(placeholder_text=t["ph_product"])
        self.lbl_f_mach.configure(text=t["lbl_machine"])
        self.ent_machine.configure(placeholder_text=t["ph_machine"])
        self.lbl_f_tot.configure(text=t["lbl_tot"])
        self.lbl_f_def.configure(text=t["lbl_def"])
        self.lbl_f_rate.configure(text=t["lbl_rate"])
        self.lbl_f_resp.configure(text=t["lbl_resp"])
        self.lbl_f_pic.configure(text=t["lbl_pic"])
        self.lbl_f_prog.configure(text=t["lbl_prog"])

        cur_prog = self.opt_prog.get()
        p_norm = normalize_prog(cur_prog)
        self.opt_prog.configure(values=t["drawer_prog_options"])
        self.opt_prog.set(t.get(f"prog_val_{p_norm}", cur_prog))

        self.lbl_sec2.configure(text=t["sec_image"])
        if self.current_pil_image is None and (not hasattr(self, "_ctk_preview") or self._ctk_preview is None):
            self.lbl_img_preview.configure(text=t["img_empty_hint"])
        self.btn_studio.configure(text=t["btn_studio"])
        self.btn_pick.configure(text=t["btn_pick"])
        self.btn_paste.configure(text=t["btn_paste"])

        self.lbl_sec3.configure(text=t["sec_defect"])
        self.lbl_f_desc.configure(text=t["lbl_desc"])
        self.lbl_f_cause.configure(text=t["lbl_cause"])
        self.lbl_f_counter.configure(text=t["lbl_counter"])
        self.lbl_f_sop.configure(text=t["lbl_sop"])
        self.lbl_f_notes.configure(text=t["lbl_notes"])

        self.btn_drawer_close.configure(text=t["btn_close"])
        self.btn_save.configure(text=t["btn_save"])

    # ── LOGIC METHODS ──
    def _calc_rate(self, event=None):
        try:
            tot = float(self.ent_tot.get().strip() or 0)
            def_q = float(self.ent_def.get().strip() or 0)
            if tot > 0:
                rate = (def_q / tot) * 100.0
                self.lbl_rate.configure(text=f"{rate:.2f}%")
                if rate > 5.0:
                    self.lbl_rate.configure(text_color=ACCENT_RED)
                elif rate > 1.0:
                    self.lbl_rate.configure(text_color=ACCENT_AMBER)
                else:
                    self.lbl_rate.configure(text_color=ACCENT_TEAL)
            else:
                self.lbl_rate.configure(text="0.00%", text_color=TEXT_MUTED)
        except Exception:
            self.lbl_rate.configure(text="0.00%", text_color=TEXT_MUTED)

    def _show_thumb_from_pil(self, pil_img: Image.Image):
        try:
            thumb = pil_img.copy()
            thumb.thumbnail((320, 180), Image.Resampling.LANCZOS)
            tw, th = thumb.size
            self._ctk_preview = ctk.CTkImage(light_image=thumb, dark_image=thumb, size=(tw, th))
            self.lbl_img_preview.configure(image=self._ctk_preview, text="")
            self.btn_studio.configure(state="normal")
        except Exception as e:
            self.lbl_img_preview.configure(text=f"[Lỗi ảnh: {e}]", image=None)

    def _show_thumb_from_bytes(self, b: bytes):
        try:
            im = Image.open(io.BytesIO(b))
            self.current_pil_image = im
            self._show_thumb_from_pil(im)
        except Exception as e:
            self.lbl_img_preview.configure(text=f"[Lỗi đọc ảnh: {e}]", image=None)

    def _load_remote_thumb(self, url: str):
        self.lbl_img_preview.configure(text="⏳ Đang tải ảnh...", image=None)
        self.btn_studio.configure(state="disabled")

        def _get():
            b = svc.download_image_bytes(url)
            if not b and (url.startswith("http://") or url.startswith("https://")):
                try:
                    import urllib.request
                    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=15, context=svc.get_ssl_context()) as resp:
                        b = resp.read()
                except Exception:
                    b = None
            if b:
                safe_after(self, 0, lambda: self._show_thumb_from_bytes(b))
            else:
                safe_after(self, 0, lambda: self.lbl_img_preview.configure(text="[Không tải được ảnh]"))

        threading.Thread(target=_get, daemon=True).start()

    def _clear_image(self):
        self.selected_image_bytes = None
        self.selected_image_name = None
        self.current_pil_image = None
        self._ctk_preview = None
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.lbl_img_preview.configure(image=None, text=t["img_empty_hint"])
        self.btn_studio.configure(state="disabled")
        if "image_url" in self.report_data:
            self.report_data["image_url"] = None

    def _pick_image(self):
        f = filedialog.askopenfilename(
            title="Chọn Ảnh Lỗi",
            filetypes=[("Hình ảnh", "*.jpg;*.jpeg;*.png;*.bmp;*.webp")]
        )
        if f:
            try:
                with open(f, "rb") as fp:
                    self.selected_image_bytes = fp.read()
                self.selected_image_name = os.path.basename(f)
                im = Image.open(f)
                self.current_pil_image = im
                self._show_thumb_from_pil(im)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file ảnh: {e}", parent=self)

    def _paste_image(self):
        try:
            from PIL import ImageGrab
            im = ImageGrab.grabclipboard()
            if isinstance(im, Image.Image):
                buf = io.BytesIO()
                im.convert("RGB").save(buf, format="JPEG", quality=90)
                self.selected_image_bytes = buf.getvalue()
                self.selected_image_name = f"clipboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                self.current_pil_image = im
                self._show_thumb_from_pil(im)
                if hasattr(self.main_view, "app") and hasattr(self.main_view.app, "show_toast"):
                    self.main_view.app.show_toast("📋 Đã dán ảnh từ Clipboard thành công!")
            else:
                messagebox.showinfo("Clipboard", "Không tìm thấy hình ảnh trong Clipboard!\nHãy chụp màn hình (PrintScreen / Win+Shift+S) rồi dán.", parent=self)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể dán ảnh từ Clipboard: {e}", parent=self)

    def _open_studio(self):
        if not self.current_pil_image and not self.report_data.get("image_url"):
            t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
            messagebox.showinfo(t["no_img_title"], t["no_img_alert"], parent=self)
            return

        def _on_updated(updated_rep):
            if isinstance(updated_rep, dict) and updated_rep.get("image_url"):
                self.report_data["image_url"] = updated_rep.get("image_url")
                self._load_remote_thumb(updated_rep.get("image_url"))
            if self.main_view:
                self.main_view.refresh_data()

        AnomalyImageStudioWindow(self.main_view, self.report_data, initial_pil=self.current_pil_image,
                                 on_image_updated=_on_updated, lang_code=self.current_lang)

    def open_for_create(self):
        self.is_edit = False
        self.report_data = {}
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.lbl_title.configure(text=t["drawer_title_create"])
        self.lbl_saving.configure(text="")
        self.btn_save.configure(state="normal")

        self.ent_date.delete(0, "end")
        self.ent_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.opt_process.set("CHA")
        self.ent_product.delete(0, "end")
        self.ent_machine.delete(0, "end")
        self.ent_tot.delete(0, "end")
        self.ent_tot.insert(0, "0")
        self.ent_def.delete(0, "end")
        self.ent_def.insert(0, "0")
        self._calc_rate()
        self.ent_resp.delete(0, "end")
        self.ent_pic.delete(0, "end")
        self.opt_prog.set(t["prog_val_in_progress"])
        self.txt_desc.delete("1.0", "end")
        self.txt_cause.delete("1.0", "end")
        self.txt_counter.delete("1.0", "end")
        self.ent_sop.delete(0, "end")
        self.ent_notes.delete(0, "end")
        self._clear_image()

        self.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.ent_product.focus()

    def open_for_edit(self, report_data: dict):
        self.is_edit = True
        self.report_data = dict(report_data)
        rep_id = report_data.get("id", "")
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.lbl_title.configure(text=f"{t['drawer_title_edit']} #{rep_id}")
        self.lbl_saving.configure(text="")
        self.btn_save.configure(state="normal")

        self.ent_date.delete(0, "end")
        self.ent_date.insert(0, str(report_data.get("report_date") or datetime.now().strftime("%Y-%m-%d")))
        if report_data.get("process"):
            self.opt_process.set(report_data.get("process"))
        self.ent_product.delete(0, "end")
        self.ent_product.insert(0, str(report_data.get("product_name") or ""))
        self.ent_machine.delete(0, "end")
        self.ent_machine.insert(0, str(report_data.get("machine") or ""))
        self.ent_tot.delete(0, "end")
        self.ent_tot.insert(0, str(report_data.get("total_qty", 0)))
        self.ent_def.delete(0, "end")
        self.ent_def.insert(0, str(report_data.get("defect_qty", 0)))
        self._calc_rate()
        self.ent_resp.delete(0, "end")
        self.ent_resp.insert(0, str(report_data.get("responsible_person") or ""))
        self.ent_pic.delete(0, "end")
        self.ent_pic.insert(0, str(report_data.get("pic") or ""))

        p_raw = report_data.get("progress") or "Đang thực hiện"
        p_norm = normalize_prog(p_raw)
        self.opt_prog.set(t.get(f"prog_val_{p_norm}", p_raw))

        self.txt_desc.delete("1.0", "end")
        self.txt_desc.insert("1.0", str(report_data.get("description") or ""))
        self.txt_cause.delete("1.0", "end")
        self.txt_cause.insert("1.0", str(report_data.get("root_cause") or ""))
        self.txt_counter.delete("1.0", "end")
        self.txt_counter.insert("1.0", str(report_data.get("countermeasures") or ""))
        self.ent_sop.delete(0, "end")
        self.ent_sop.insert(0, str(report_data.get("sop_standard") or ""))
        self.ent_notes.delete(0, "end")
        self.ent_notes.insert(0, str(report_data.get("notes") or ""))

        self.selected_image_bytes = None
        self.selected_image_name = None
        self.current_pil_image = None
        img_url = report_data.get("image_url")
        if img_url:
            self._load_remote_thumb(img_url)
        else:
            self._clear_image()

        self.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.ent_product.focus()

    def close_drawer(self):
        self.grid_forget()

    def _save_report(self):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        prod = self.ent_product.get().strip()
        if not prod:
            messagebox.showwarning("Thiếu thông tin" if self.current_lang == "vi" else ("缺少信息" if self.current_lang == "zh" else "Missing Info"),
                                   "Vui lòng nhập Tên Sản Phẩm (Model/PWB)!" if self.current_lang == "vi" else ("请输入产品名称 (型号/PWB)！" if self.current_lang == "zh" else "Please enter Product Name (Model/PWB)!"),
                                   parent=self)
            self.ent_product.focus()
            return

        try:
            tot_q = int(self.ent_tot.get().strip() or 0)
            def_q = int(self.ent_def.get().strip() or 0)
        except ValueError:
            messagebox.showwarning("Lỗi số lượng" if self.current_lang == "vi" else ("数量格式错误" if self.current_lang == "zh" else "Quantity Error"),
                                   "Số lượng kiểm tra và số lượng lỗi phải là số nguyên!" if self.current_lang == "vi" else ("检查数和不良数必须为整数！" if self.current_lang == "zh" else "Quantities must be integers!"),
                                   parent=self)
            return

        p_choice = self.opt_prog.get().strip()
        p_norm = normalize_prog(p_choice)
        std_prog_map = {"open": "Chưa thực hiện", "in_progress": "Đang thực hiện", "done": "Đã hoàn thành"}

        payload = {
            "report_date": self.ent_date.get().strip() or datetime.now().strftime("%Y-%m-%d"),
            "process": self.opt_process.get().strip(),
            "product_name": prod,
            "machine": self.ent_machine.get().strip(),
            "total_qty": tot_q,
            "defect_qty": def_q,
            "responsible_person": self.ent_resp.get().strip(),
            "pic": self.ent_pic.get().strip(),
            "progress": std_prog_map.get(p_norm, "Đang thực hiện"),
            "description": self.txt_desc.get("1.0", "end-1c").strip(),
            "root_cause": self.txt_cause.get("1.0", "end-1c").strip(),
            "countermeasures": self.txt_counter.get("1.0", "end-1c").strip(),
            "sop_standard": self.ent_sop.get().strip(),
            "notes": self.ent_notes.get().strip(),
        }
        if "image_url" in self.report_data and self.report_data["image_url"] is not None:
            payload["image_url"] = self.report_data["image_url"]

        self.btn_save.configure(state="disabled")
        self.lbl_saving.configure(text=t["saving_txt"])

        def _do_save():
            try:
                if self.is_edit:
                    rep_id = self.report_data["id"]
                    res = svc.update_report(
                        rep_id,
                        payload,
                        new_image_bytes=self.selected_image_bytes,
                        new_image_name=self.selected_image_name
                    )
                else:
                    res = svc.create_report(
                        payload,
                        image_bytes=self.selected_image_bytes,
                        original_image_name=self.selected_image_name
                    )
                safe_after(self, 0, lambda: self._on_save_success(res))
            except Exception as e:
                safe_after(self, 0, lambda: self._on_save_error(str(e)))

        threading.Thread(target=_do_save, daemon=True).start()

    def _on_save_success(self, res):
        self.btn_save.configure(state="normal")
        self.lbl_saving.configure(text="")
        self.close_drawer()
        if self.on_saved:
            self.on_saved(res)

    def _on_save_error(self, err_msg):
        self.btn_save.configure(state="normal")
        self.lbl_saving.configure(text="")
        messagebox.showerror("Lỗi Lưu Dữ Liệu", f"Không thể lưu lên Supabase:\n{err_msg}", parent=self)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN VIEW: ANOMALY REPORT VIEW (FULL-WIDTH TREEVIEW TABLE)
# ─────────────────────────────────────────────────────────────────────────────
class AnomalyReportView(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.current_lang = getattr(app, "current_lang", "vi")

        self.all_reports: list[dict] = []
        self.filtered_reports: list[dict] = []
        self.selected_report: dict = None
        self.is_admin_session = False

        # Thumbnail cache & high-performance image loading
        self._thumb_cache: dict[str, ImageTk.PhotoImage] = {}
        self._thumb_disk_cache_dir = os.path.join(tempfile.gettempdir(), "vipqc_anomaly_thumbs")
        os.makedirs(self._thumb_disk_cache_dir, exist_ok=True)
        self._img_no_photo: ImageTk.PhotoImage = None
        self._img_loading_photo: ImageTk.PhotoImage = None
        self._make_placeholder_images(is_dark=True)

        self._build_ui()
        self.refresh_data()

    def _make_placeholder_images(self, is_dark: bool = True):
        """Generates crisp 46x46 placeholder tiles for rows without images or while loading."""
        tile_bg = (22, 29, 46, 255) if is_dark else (226, 232, 240, 255)
        border_col = (42, 54, 80, 255) if is_dark else (203, 213, 225, 255)
        icon_col = (100, 116, 139, 255) if is_dark else (148, 163, 184, 255)
        teal_col = (13, 148, 136, 255) if is_dark else (15, 118, 110, 255)

        # No Image Placeholder: dark tile with subtle dash
        im_no = Image.new("RGBA", (46, 46), tile_bg)
        d1 = ImageDraw.Draw(im_no)
        d1.rectangle([0, 0, 45, 45], outline=border_col, width=1)
        d1.line([(16, 23), (30, 23)], fill=icon_col, width=2)
        self._img_no_photo = ImageTk.PhotoImage(im_no)

        # Loading Placeholder: dark tile with dots indicator
        im_load = Image.new("RGBA", (46, 46), tile_bg)
        d2 = ImageDraw.Draw(im_load)
        d2.rectangle([0, 0, 45, 45], outline=teal_col, width=1)
        d2.text((12, 14), "...", fill=teal_col)
        self._img_loading_photo = ImageTk.PhotoImage(im_load)

    def _create_thumbnail_image(self, im: Image.Image, size=(46, 46)) -> Image.Image:
        """Crops/fits image into a crisp, polished square thumbnail with subtle border."""
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        im = im.convert("RGBA")
        w, h = im.size
        ratio = min(size[0] / max(w, 1), size[1] / max(h, 1))
        new_w, new_h = max(1, int(w * ratio)), max(1, int(h * ratio))
        im_resized = im.resize((new_w, new_h), Image.Resampling.LANCZOS)

        bg = Image.new("RGBA", size, (28, 36, 56, 255))
        ox = (size[0] - new_w) // 2
        oy = (size[1] - new_h) // 2
        bg.paste(im_resized, (ox, oy), im_resized)

        # Draw subtle border
        draw = ImageDraw.Draw(bg)
        draw.rectangle([0, 0, size[0] - 1, size[1] - 1], outline=(42, 54, 80, 255), width=1)
        return bg

    def _get_or_load_thumb(self, url: str, report_id: str) -> ImageTk.PhotoImage:
        """Returns cached PhotoImage or initiates async download and returns placeholder."""
        if not url:
            return self._img_no_photo

        if url in self._thumb_cache:
            return self._thumb_cache[url]

        # Check local disk cache
        h = hashlib.md5(url.encode("utf-8")).hexdigest()
        disk_path = os.path.join(self._thumb_disk_cache_dir, f"{h}.png")
        if os.path.exists(disk_path):
            try:
                im = Image.open(disk_path)
                photo = ImageTk.PhotoImage(im)
                self._thumb_cache[url] = photo
                return photo
            except Exception:
                pass

        # Async fetch from cloud
        def _fetch():
            try:
                b = svc.download_image_bytes(url)
                if not b and (url.startswith("http://") or url.startswith("https://")):
                    import urllib.request
                    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=15, context=svc.get_ssl_context()) as resp:
                        b = resp.read()
                if not b:
                    return
                orig_im = Image.open(io.BytesIO(b))
                thumb_im = self._create_thumbnail_image(orig_im, (46, 46))
                try:
                    thumb_im.save(disk_path, "PNG")
                except Exception:
                    pass

                def _apply():
                    photo = ImageTk.PhotoImage(thumb_im)
                    self._thumb_cache[url] = photo
                    self._thumb_cache[str(report_id)] = photo
                    try:
                        if self.tree.exists(report_id):
                            self.tree.item(report_id, image=photo)
                    except Exception:
                        pass

                safe_after(self, 0, _apply)
            except Exception:
                pass

        threading.Thread(target=_fetch, daemon=True).start()
        return self._img_loading_photo

    def _build_ui(self):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])

        # ── TOP ACTION & FILTER BAR ──────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                               border_width=1, border_color=BORDER_CLR, height=64)
        top_bar.pack(fill="x", pady=(0, 10))
        top_bar.pack_propagate(False)

        # Search box
        self.ent_search = ctk.CTkEntry(top_bar, placeholder_text=t["search_placeholder"],
                                       width=280, height=36, corner_radius=8, font=FONT_BODY)
        self.ent_search.pack(side="left", padx=(14, 6), pady=14)
        self.ent_search.bind("<KeyRelease>", lambda e: self._apply_filter())

        # Process Filter
        self.opt_filter_proc = ctk.CTkOptionMenu(
            top_bar,
            values=t["proc_options"],
            height=36, corner_radius=8, font=FONT_SMALL, width=130,
            command=lambda v: self._apply_filter()
        )
        self.opt_filter_proc.pack(side="left", padx=4)

        # Progress Filter
        self.opt_filter_prog = ctk.CTkOptionMenu(
            top_bar,
            values=t["prog_options"],
            height=36, corner_radius=8, font=FONT_SMALL, width=130,
            command=lambda v: self._apply_filter()
        )
        self.opt_filter_prog.pack(side="left", padx=4)

        # Right-side action buttons
        self.btn_admin_lock = ctk.CTkButton(
            top_bar, text=t["btn_admin_locked"], width=125, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._toggle_admin_lock
        )
        self.btn_admin_lock.pack(side="right", padx=(4, 14))

        self.btn_export = ctk.CTkButton(
            top_bar, text=t["btn_export"], width=105, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._export_excel
        )
        self.btn_export.pack(side="right", padx=4)

        self.btn_refresh = ctk.CTkButton(
            top_bar, text=t["btn_refresh"], width=95, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self.refresh_data
        )
        self.btn_refresh.pack(side="right", padx=4)

        # Action: Delete selected
        self.btn_delete = ctk.CTkButton(
            top_bar, text=t["btn_delete"], width=80, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._on_delete_report, state="disabled"
        )
        self.btn_delete.pack(side="right", padx=4)

        # Action: Edit selected
        self.btn_edit = ctk.CTkButton(
            top_bar, text=t["btn_edit"], width=80, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._on_edit_report, state="disabled"
        )
        self.btn_edit.pack(side="right", padx=4)

        # Action: Open Image Marking Studio
        self.btn_open_marking = ctk.CTkButton(
            top_bar, text=t["btn_open_marking"], width=120, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=ACCENT_BLUE, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._open_marking_studio, state="disabled"
        )
        self.btn_open_marking.pack(side="right", padx=4)

        # Action: Add Report
        self.btn_add = ctk.CTkButton(
            top_bar, text=t["btn_add"], width=135, height=36, corner_radius=8,
            fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
            font=FONT_H2, command=self._on_add_report
        )
        self.btn_add.pack(side="right", padx=4)

        # ── MAIN CONTENT AREA (TABLE ON LEFT, DRAWER ON RIGHT) ──────────────
        self.content_area = ctk.CTkFrame(self, fg_color="transparent")
        self.content_area.pack(fill="both", expand=True)

        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(1, weight=0)

        self.tbl_card = ctk.CTkFrame(self.content_area, fg_color=BG_CARD, corner_radius=12,
                                     border_width=1, border_color=BORDER_CLR)
        self.tbl_card.grid(row=0, column=0, sticky="nsew")

        # Drawer on right (docked, hidden until opened)
        self.drawer = AnomalyDrawerFrame(self.content_area, self, on_saved=self._on_drawer_saved)

        # Table Header Info Bar
        tbl_info = ctk.CTkFrame(self.tbl_card, fg_color="transparent", height=34)
        tbl_info.pack(fill="x", padx=16, pady=(10, 4))

        self.lbl_tbl_count = ctk.CTkLabel(tbl_info, text=t["tbl_count"].format(count=0),
                                          font=FONT_H2, text_color=TEXT_PRIMARY)
        self.lbl_tbl_count.pack(side="left")

        self.lbl_hint = ctk.CTkLabel(tbl_info,
                                     text=t["hint_label"],
                                     font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_hint.pack(side="right")

        # Treeview with 16 data columns + #0 Thumbnail Image column
        tree_container = tk.Frame(self.tbl_card, bg="#131929")
        tree_container.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Anomaly.Treeview", rowheight=60, font=("Segoe UI", 12))
        style.configure("Anomaly.Treeview.Heading", font=("Segoe UI", 11, "bold"), padding=[6, 8])

        cols = (
            "stt", "date", "process", "product", "machine", "tot", "def", "rate",
            "resp", "pic", "desc", "cause", "counter", "sop", "prog", "notes"
        )
        self.tree = ttk.Treeview(tree_container, columns=cols, show="tree headings",
                                 style="Anomaly.Treeview", selectmode="browse")

        # ── PERMANENT TRILINGUAL COLUMN HEADERS (ENGLISH / TIẾNG VIỆT / 中文) ──
        self.tree.heading("#0", text="Image / Ảnh\n图片", anchor="center")
        self.tree.column("#0", width=120, minwidth=100, anchor="center", stretch=False)

        headers_meta = [
            ("stt",     "No. / STT\n序号",                         90,  "center"),
            ("date",    "Date / Ngày\n日期",                       140, "center"),
            ("process", "Process / Công Đoạn\n工序 (过程)",        170, "center"),
            ("product", "Product / Sản Phẩm\n产品 (Model/PWB)",    260, "center"),
            ("machine", "Machine / Thiết Bị\n设备 (Line/Chuyền)",  200, "center"),
            ("tot",     "Total / SL Kiểm\n检查数",                 135, "center"),
            ("def",     "Defect / SL Lỗi\n不良数",                 135, "center"),
            ("rate",    "Rate / Tỷ Lệ\n不良率 (%)",                140, "center"),
            ("resp",    "Resp. / Chịu TN\n责任人 (Manager)",       185, "center"),
            ("pic",     "PIC / Phụ Trách\n担当者 (Inspector)",     185, "center"),
            ("desc",    "Defect / Hiện Tượng Lỗi\n不良现象描述",   340, "center"),
            ("cause",   "Cause / Nguyên Nhân\n原因分析",           300, "center"),
            ("counter", "Action / Biện Pháp\n改善对策",            320, "center"),
            ("sop",     "SOP / Tiêu Chuẩn\nSOP标准",               160, "center"),
            ("prog",    "Progress / Tiến Độ\n进度状态",            160, "center"),
            ("notes",   "Notes / Ghi Chú\n备注",                   220, "center")
        ]

        for col_id, col_name, col_w, anchor in headers_meta:
            self.tree.heading(col_id, text=col_name, anchor=anchor)
            self.tree.column(col_id, width=col_w, anchor=anchor, minwidth=max(60, col_w - 40), stretch=False)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Treeview Interactions
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self.tree.bind("<Double-Button-1>", self._on_row_double_click)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_cell_click)
        self.tree.bind("<Button-3>", self._show_context_menu)

        # Context Menu
        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.ctx_menu.add_command(label=t["ctx_open_marking"], command=self._open_marking_studio)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label=t["ctx_edit"], command=self._on_edit_report)
        self.ctx_menu.add_command(label=t["ctx_delete"], command=self._on_delete_report)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label=t["ctx_copy"], command=self._copy_row_info)

    def update_language(self, lang_code: str = None):
        """Switches all labels, filters, buttons, context menus and row status to selected language."""
        if lang_code:
            self.current_lang = lang_code
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])

        self.ent_search.configure(placeholder_text=t["search_placeholder"])

        cur_proc = self.opt_filter_proc.get()
        self.opt_filter_proc.configure(values=t["proc_options"])
        if cur_proc in ["Tất cả công đoạn", "全部过程", "All Processes"]:
            self.opt_filter_proc.set(t["all_proc"])

        cur_prog = self.opt_filter_prog.get()
        self.opt_filter_prog.configure(values=t["prog_options"])
        if cur_prog in ["Tất cả tiến độ", "全部进度", "All Progress"]:
            self.opt_filter_prog.set(t["all_prog"])
        else:
            p_norm = normalize_prog(cur_prog)
            self.opt_filter_prog.set(t.get(f"prog_val_{p_norm}", cur_prog))

        self.btn_add.configure(text=t["btn_add"])
        self.btn_open_marking.configure(text=t["btn_open_marking"])
        self.btn_edit.configure(text=t["btn_edit"])
        self.btn_delete.configure(text=t["btn_delete"])
        self.btn_refresh.configure(text=t["btn_refresh"])
        self.btn_export.configure(text=t["btn_export"])
        if self.is_admin_session:
            self.btn_admin_lock.configure(text=t["btn_admin_unlocked"])
        else:
            self.btn_admin_lock.configure(text=t["btn_admin_locked"])

        cur_count = len(self.filtered_reports) if hasattr(self, "filtered_reports") and self.filtered_reports is not None else len(self.all_reports)
        self.lbl_tbl_count.configure(text=t["tbl_count"].format(count=cur_count))
        self.lbl_hint.configure(text=t["hint_label"])

        try:
            self.ctx_menu.entryconfigure(0, label=t["ctx_open_marking"])
            self.ctx_menu.entryconfigure(2, label=t["ctx_edit"])
            self.ctx_menu.entryconfigure(3, label=t["ctx_delete"])
            self.ctx_menu.entryconfigure(5, label=t["ctx_copy"])
        except Exception:
            pass

        if hasattr(self, "drawer"):
            self.drawer.update_language(self.current_lang)

        # Repopulate tree so row status reflects the new language
        if hasattr(self, "filtered_reports") and self.filtered_reports:
            self._populate_tree(self.filtered_reports)
        elif self.all_reports:
            self._populate_tree(self.all_reports)

    # ── DATA FETCH & POPULATE ────────────────────────────────────────────────
    def refresh_data(self):
        """Fetches reports from Supabase asynchronously."""
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.btn_refresh.configure(state="disabled", text=t["btn_refreshing"])

        def _fetch():
            try:
                reps = svc.fetch_all_reports()
                safe_after(self, 0, lambda: self._on_fetch_success(reps))
            except Exception as e:
                safe_after(self, 0, lambda: self._on_fetch_error(str(e)))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_fetch_success(self, reports: list[dict]):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.all_reports = reports
        self.selected_report = None
        self._update_action_buttons_state()
        self.btn_refresh.configure(state="normal", text=t["btn_refresh"])
        self._apply_filter()
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(f"✅ Đồng bộ {len(reports)} báo cáo từ Supabase Cloud." if self.current_lang == "vi" else (
                f"✅ 已从 Supabase 云端同步 {len(reports)} 条报告。" if self.current_lang == "zh" else
                f"✅ Synced {len(reports)} reports from Supabase Cloud."
            ))

    def _on_fetch_error(self, err: str):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.btn_refresh.configure(state="normal", text=t["btn_refresh"])
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(f"⚠️ Lỗi kết nối Cloud: {err[:50]}...")

    def _apply_filter(self):
        q = self.ent_search.get().strip().lower()
        f_proc = self.opt_filter_proc.get()
        f_prog = self.opt_filter_prog.get()

        is_all_proc = f_proc in ["Tất cả công đoạn", "全部过程", "All Processes"]
        is_all_prog = f_prog in ["Tất cả tiến độ", "全部进度", "All Progress"]
        prog_target_norm = normalize_prog(f_prog) if not is_all_prog else None

        res = []
        for r in self.all_reports:
            if not is_all_proc:
                r_proc = str(r.get("process") or "")
                if normalize_proc(r_proc) != normalize_proc(f_proc):
                    continue
            if not is_all_prog:
                r_prog = str(r.get("progress") or "")
                if normalize_prog(r_prog) != prog_target_norm:
                    continue
            if q:
                searchable = f"{r.get('product_name', '')} {r.get('process', '')} {r.get('machine', '')} {r.get('pic', '')} {r.get('responsible_person', '')} {r.get('description', '')} {r.get('root_cause', '')} {r.get('countermeasures', '')}".lower()
                if q not in searchable:
                    continue
            res.append(r)

        self.filtered_reports = res
        self._populate_tree(res)

    def _populate_tree(self, reports: list[dict]):
        self.tree.delete(*self.tree.get_children())
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        self.lbl_tbl_count.configure(text=t["tbl_count"].format(count=len(reports)))

        # Clear selection state if empty
        if not reports:
            self.selected_report = None
            self._update_action_buttons_state()

        for idx, r in enumerate(reports, 1):
            rate_val = r.get("defect_rate")
            rate_str = f"{rate_val:.2f}%" if rate_val is not None else "0.00%"

            tag = "even" if idx % 2 == 0 else "odd"
            prog_raw = str(r.get("progress") or "Đang thực hiện")
            p_norm = normalize_prog(prog_raw)
            if p_norm == "done":
                tag = "tag_done"
            elif p_norm == "open":
                tag = "tag_open"

            prog_display = t.get(f"prog_val_{p_norm}", prog_raw)

            vals = (
                idx,
                r.get("report_date", ""),
                r.get("process", ""),
                r.get("product_name", ""),
                r.get("machine", ""),
                f"{r.get('total_qty', 0):,}",
                f"{r.get('defect_qty', 0):,}",
                rate_str,
                r.get("responsible_person", ""),
                r.get("pic", ""),
                r.get("description", ""),
                r.get("root_cause", ""),
                r.get("countermeasures", ""),
                r.get("sop_standard", ""),
                prog_display,
                r.get("notes", "")
            )
            rep_id = str(r.get("id"))
            thumb_photo = self._get_or_load_thumb(r.get("image_url"), rep_id)
            self.tree.insert("", "end", iid=rep_id, text="", image=thumb_photo, values=vals, tags=(tag,))

    def _on_row_select(self, event):
        selected_ids = self.tree.selection()
        if not selected_ids:
            self.selected_report = None
            self._update_action_buttons_state()
            return
        rep_id = selected_ids[0]
        found = [r for r in self.all_reports if str(r.get("id")) == str(rep_id)]
        if found:
            self.selected_report = found[0]
            self._update_action_buttons_state()

    def _on_row_double_click(self, event):
        """Double clicking a row opens the Drawer to edit the report."""
        if not self.selected_report:
            return
        self._on_edit_report()

    def _on_tree_cell_click(self, event):
        """Clicking directly on the 'Image / Ảnh / 图片' column cell (#0) opens the Image Studio, or prompts to add image in Drawer."""
        col_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        region = self.tree.identify_region(event.x, event.y)
        if row_id and (col_id == "#0" or region == "tree"):
            self.tree.selection_set(row_id)
            self._on_row_select(None)
            found = [r for r in self.all_reports if str(r.get("id")) == str(row_id)]
            if found:
                t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
                if found[0].get("image_url"):
                    self.after(50, self._open_marking_studio)
                else:
                    if messagebox.askyesno(t["no_img_title"], t["no_img_prompt"], parent=self):
                        self._on_edit_report()

    def _update_action_buttons_state(self):
        has_sel = self.selected_report is not None
        self.btn_edit.configure(state="normal" if has_sel else "disabled")
        self.btn_delete.configure(state="normal" if has_sel else "disabled")
        has_img = has_sel and bool(self.selected_report.get("image_url"))
        self.btn_open_marking.configure(state="normal" if has_img else "disabled")

    def _show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self._on_row_select(None)
            try:
                self.ctx_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.ctx_menu.grab_release()

    def _copy_row_info(self):
        if not self.selected_report:
            return
        r = self.selected_report
        line = f"Sản phẩm: {r.get('product_name')} | Công đoạn: {r.get('process')} | Ngày: {r.get('report_date')} | Lỗi: {r.get('description')} | Tỷ lệ: {r.get('defect_rate')}%"
        self.clipboard_clear()
        self.clipboard_append(line)
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(t["copied_toast"])

    # ── OPEN IMAGE MARKING STUDIO ────────────────────────────────────────────
    def _open_marking_studio(self):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        if not self.selected_report or not self.selected_report.get("image_url"):
            messagebox.showinfo(t["no_img_title"], t["no_img_alert"])
            return

        def _on_img_updated(new_rep):
            self.refresh_data()

        AnomalyImageStudioWindow(self, self.selected_report, on_image_updated=_on_img_updated, lang_code=self.current_lang)

    # ── ADMIN & PASSWORD VERIFICATION ────────────────────────────────────────
    def _require_admin(self, callback):
        if self.is_admin_session:
            callback()
        else:
            def _on_ok():
                self.is_admin_session = True
                t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
                self.btn_admin_lock.configure(text=t["btn_admin_unlocked"], fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"))
                callback()
            PasswordDialog(self, on_success=_on_ok, lang_code=self.current_lang)

    def _toggle_admin_lock(self):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        if self.is_admin_session:
            m = messagebox.askyesnocancel(
                t["pwd_title"],
                "Bạn đang ở chế độ Quản Trị.\n\n• Chọn YES để Khóa lại\n• Chọn NO để Đổi Mật Khẩu\n• Chọn CANCEL để đóng" if self.current_lang == "vi" else (
                    "当前处于管理员模式。\n\n• YES: 重新锁定\n• NO: 修改密码\n• CANCEL: 关闭" if self.current_lang == "zh" else
                    "You are in Admin mode.\n\n• YES: Re-lock\n• NO: Change Password\n• CANCEL: Close"
                )
            )
            if m is True:
                self.is_admin_session = False
                self.btn_admin_lock.configure(text=t["btn_admin_locked"], fg_color=BG_SURFACE, text_color=TEXT_PRIMARY)
                messagebox.showinfo(t["pwd_title"], "Đã khóa phiên Quản Trị thành công." if self.current_lang == "vi" else ("已锁定管理员会话。" if self.current_lang == "zh" else "Admin session locked."))
            elif m is False:
                ChangePasswordDialog(self, lang_code=self.current_lang)
        else:
            def _on_ok():
                self.is_admin_session = True
                self.btn_admin_lock.configure(text=t["btn_admin_unlocked"], fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"))
                messagebox.showinfo(t["pwd_title"], "✅ Mở khóa quyền Quản Trị thành công!" if self.current_lang == "vi" else ("✅ 管理员权限解锁成功！" if self.current_lang == "zh" else "✅ Admin unlocked successfully!"))
            PasswordDialog(self, on_success=_on_ok, lang_code=self.current_lang)

    # ── CRUD ACTIONS ─────────────────────────────────────────────────────────
    def _on_add_report(self):
        self.drawer.open_for_create()

    def _on_edit_report(self):
        if not self.selected_report:
            t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
            messagebox.showwarning("Chưa chọn" if self.current_lang == "vi" else ("未选择" if self.current_lang == "zh" else "Not Selected"),
                                   "Vui lòng click chọn một báo cáo trên bảng để chỉnh sửa!" if self.current_lang == "vi" else ("请在表格中选择一条报告进行编辑！" if self.current_lang == "zh" else "Please select a report row to edit!"))
            return
        self.drawer.open_for_edit(self.selected_report)

    def _on_drawer_saved(self, saved_rep):
        if isinstance(saved_rep, dict) and saved_rep.get("image_url"):
            url = saved_rep.get("image_url")
            if hasattr(self.drawer, "current_pil_image") and self.drawer.current_pil_image:
                try:
                    thumb_im = self._create_thumbnail_image(self.drawer.current_pil_image, (46, 46))
                    photo = ImageTk.PhotoImage(thumb_im)
                    self._thumb_cache[url] = photo
                except Exception:
                    pass
        self.refresh_data()
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(t["saved_toast"])

    def _on_delete_report(self):
        if not self.selected_report:
            messagebox.showwarning("Chưa chọn" if self.current_lang == "vi" else ("未选择" if self.current_lang == "zh" else "Not Selected"),
                                   "Vui lòng click chọn một báo cáo trên bảng để xóa!" if self.current_lang == "vi" else ("请在表格中选择一条报告进行删除！" if self.current_lang == "zh" else "Please select a report row to delete!"))
            return

        rep = self.selected_report
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        def _do_del():
            confirm = messagebox.askyesno(
                "Xác Nhận Xóa" if self.current_lang == "vi" else ("确认删除" if self.current_lang == "zh" else "Confirm Delete"),
                f"Bạn có chắc chắn muốn xóa vĩnh viễn báo cáo này?\n\n• Sản phẩm: {rep.get('product_name')}\n• Công đoạn: {rep.get('process')}\n• Ngày: {rep.get('report_date')}" if self.current_lang == "vi" else (
                    f"确定要永久删除此异常报告吗？\n\n• 产品: {rep.get('product_name')}\n• 过程: {rep.get('process')}\n• 日期: {rep.get('report_date')}" if self.current_lang == "zh" else
                    f"Are you sure you want to permanently delete this report?\n\n• Product: {rep.get('product_name')}\n• Process: {rep.get('process')}\n• Date: {rep.get('report_date')}"
                )
            )
            if confirm:
                ok = svc.delete_report(rep["id"], rep.get("image_url"))
                if ok:
                    if hasattr(self, "drawer") and self.drawer.winfo_manager() == "grid":
                        if str(self.drawer.report_data.get("id")) == str(rep.get("id")):
                            self.drawer.close_drawer()
                    self.selected_report = None
                    self._update_action_buttons_state()
                    self.refresh_data()
                    if hasattr(self.app, "show_toast"):
                        self.app.show_toast(t["deleted_toast"])
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa báo cáo trên Supabase!" if self.current_lang == "vi" else ("无法在云端删除此报告！" if self.current_lang == "zh" else "Failed to delete report on Cloud!"))

        self._require_admin(_do_del)

    def _export_excel(self):
        t = I18N_ANOMALY.get(self.current_lang, I18N_ANOMALY["vi"])
        if not self.filtered_reports:
            messagebox.showwarning("Trống" if self.current_lang == "vi" else ("空" if self.current_lang == "zh" else "Empty"),
                                   "Không có bản ghi nào để xuất Excel!" if self.current_lang == "vi" else ("暂无记录可导出 Excel！" if self.current_lang == "zh" else "No records to export!"))
            return

        def_name = f"Bao_Cao_Bat_Thuong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        f = filedialog.asksaveasfilename(
            title="Lưu Báo Cáo Excel" if self.current_lang == "vi" else ("保存 Excel 报告" if self.current_lang == "zh" else "Save Excel Report"),
            initialfile=def_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if f:
            try:
                saved = svc.export_reports_to_excel(self.filtered_reports, f)
                ask_open = messagebox.askyesno(
                    "Xuất Thành Công",
                    f"Đã xuất {len(self.filtered_reports)} báo cáo ra Excel:\n{saved}\n\nBạn có muốn mở file ngay không?"
                )
                if ask_open:
                    os.startfile(saved)
            except Exception as e:
                messagebox.showerror("Lỗi Xuất Excel", f"Không thể xuất file Excel:\n{e}")

    # ── THEME SUPPORT ────────────────────────────────────────────────────────
    def apply_theme(self, is_dark: bool):
        """Applies theme colors to Treeview tags and elements."""
        tree_bg = "#131929" if is_dark else "#FFFFFF"
        tree_fg = "#E8EDF5" if is_dark else "#0F172A"
        hdr_bg = "#0B0F1A" if is_dark else "#E2E8F0"
        hdr_fg = "#00C9A7" if is_dark else "#0D9488"
        sel_bg = "#232D45" if is_dark else "#BAE6FD"
        sel_fg = "#00C9A7" if is_dark else "#0369A1"
        odd_bg = "#1C2438" if is_dark else "#F8FAFC"
        even_bg = "#131929" if is_dark else "#FFFFFF"

        style = ttk.Style()
        style.configure("Anomaly.Treeview",
                        background=tree_bg,
                        fieldbackground=tree_bg,
                        foreground=tree_fg,
                        rowheight=60,
                        font=("Segoe UI", 12),
                        borderwidth=0)
        style.configure("Anomaly.Treeview.Heading",
                        background=hdr_bg,
                        foreground=hdr_fg,
                        font=("Segoe UI", 11, "bold"),
                        padding=[6, 8])
        style.map("Anomaly.Treeview",
                  background=[("selected", sel_bg)],
                  foreground=[("selected", sel_fg)])

        self._make_placeholder_images(is_dark=is_dark)

        self.tree.tag_configure("odd", background=odd_bg)
        self.tree.tag_configure("even", background=even_bg)
        self.tree.tag_configure("tag_done", background="#064E3B" if is_dark else "#DCFCE7",
                                foreground="#A7F3D0" if is_dark else "#166534")
        self.tree.tag_configure("tag_open", background="#7F1D1D" if is_dark else "#FEE2E2",
                                foreground="#FECACA" if is_dark else "#991B1B")
