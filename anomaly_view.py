"""
VIPQC AI — Anomaly Report View Module
Modern CustomTkinter GUI for viewing, creating, updating, deleting,
filtering, and zooming defect images with Supabase Cloud backend.
"""

import os
import io
import threading
from datetime import datetime
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageTk

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

FONT_H1     = ("Segoe UI", 13, "bold")
FONT_H2     = ("Segoe UI", 11, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_SMALL  = ("Segoe UI", 9)


# ─────────────────────────────────────────────────────────────────────────────
# DIALOG: PASSWORD AUTHENTICATION
# ─────────────────────────────────────────────────────────────────────────────
class PasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent, on_success):
        super().__init__(parent)
        self.title("Xác Thực Quản Trị")
        self.geometry("400x250")
        self.resizable(False, False)
        self.configure(fg_color=BG_CARD)
        self.transient(parent)
        self.grab_set()

        self.on_success = on_success

        # Center dialog
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 200
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 125
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        pad = {"padx": 24}

        lbl_icon = ctk.CTkLabel(self, text="🔒", font=("Segoe UI", 28))
        lbl_icon.pack(pady=(18, 4))

        lbl_title = ctk.CTkLabel(self, text="Nhập Mật Khẩu Quản Trị (QC)", font=FONT_H1, text_color=TEXT_PRIMARY)
        lbl_title.pack(pady=(0, 12))

        self.entry_pass = ctk.CTkEntry(self, placeholder_text="Nhập mật khẩu...", show="•",
                                       height=38, corner_radius=8, font=FONT_BODY)
        self.entry_pass.pack(fill="x", pady=(0, 8), **pad)
        self.entry_pass.focus()
        self.entry_pass.bind("<Return>", lambda e: self._submit())

        self.lbl_err = ctk.CTkLabel(self, text="", font=FONT_SMALL, text_color=ACCENT_RED)
        self.lbl_err.pack(pady=(0, 10))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", **pad)

        btn_cancel = ctk.CTkButton(btn_row, text="Hủy", width=100, height=36, corner_radius=8,
                                   fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                                   command=self.destroy)
        btn_cancel.pack(side="left")

        self.btn_ok = ctk.CTkButton(btn_row, text="Xác Nhận", width=140, height=36, corner_radius=8,
                                    fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
                                    font=FONT_H2, command=self._submit)
        self.btn_ok.pack(side="right")

    def _submit(self):
        pwd = self.entry_pass.get().strip()
        if not pwd:
            self.lbl_err.configure(text="Vui lòng nhập mật khẩu!")
            return

        if svc.verify_admin_password(pwd):
            self.destroy()
            if self.on_success:
                self.on_success()
        else:
            self.lbl_err.configure(text="❌ Mật khẩu không chính xác! Vui lòng thử lại.")
            self.entry_pass.delete(0, "end")


# ─────────────────────────────────────────────────────────────────────────────
# DIALOG: CHANGE PASSWORD
# ─────────────────────────────────────────────────────────────────────────────
class ChangePasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Đổi Mật Khẩu Quản Trị")
        self.geometry("420x330")
        self.resizable(False, False)
        self.configure(fg_color=BG_CARD)
        self.transient(parent)
        self.grab_set()

        pad = {"padx": 24}

        ctk.CTkLabel(self, text="🔑 Đổi Mật Khẩu Quản Trị", font=FONT_H1, text_color=TEXT_PRIMARY).pack(pady=(20, 14))

        self.old_pass = ctk.CTkEntry(self, placeholder_text="Mật khẩu hiện tại", show="•", height=36, font=FONT_BODY)
        self.old_pass.pack(fill="x", pady=4, **pad)

        self.new_pass = ctk.CTkEntry(self, placeholder_text="Mật khẩu mới (ít nhất 4 ký tự)", show="•", height=36, font=FONT_BODY)
        self.new_pass.pack(fill="x", pady=4, **pad)

        self.confirm_pass = ctk.CTkEntry(self, placeholder_text="Xác nhận mật khẩu mới", show="•", height=36, font=FONT_BODY)
        self.confirm_pass.pack(fill="x", pady=4, **pad)

        self.lbl_status = ctk.CTkLabel(self, text="", font=FONT_SMALL, text_color=ACCENT_RED)
        self.lbl_status.pack(pady=6)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", pady=(8, 16), **pad)

        ctk.CTkButton(btn_row, text="Đóng", width=100, height=36, corner_radius=8,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self.destroy).pack(side="left")

        ctk.CTkButton(btn_row, text="💾 Lưu Mật Khẩu", width=140, height=36, corner_radius=8,
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
# DIALOG: HIGH-RES IMAGE ZOOM VIEWER
# ─────────────────────────────────────────────────────────────────────────────
class ImageZoomViewer(ctk.CTkToplevel):
    def __init__(self, parent, image_url: str, title: str = "Chi Tiết Ảnh Lỗi"):
        super().__init__(parent)
        self.title(title)
        self.geometry("980x720")
        self.minsize(700, 500)
        self.configure(fg_color=BG_DEEP)
        self.transient(parent)

        self.image_url = image_url
        self.original_pil = None
        self.zoom_level = 1.0

        # Top Bar
        top_bar = ctk.CTkFrame(self, fg_color=BG_CARD, height=48, corner_radius=0)
        top_bar.pack(fill="x")
        top_bar.pack_propagate(False)

        ctk.CTkLabel(top_bar, text=f"🔍 {title}", font=FONT_H2, text_color=TEXT_PRIMARY).pack(side="left", padx=16)

        self.lbl_zoom = ctk.CTkLabel(top_bar, text="100%", font=FONT_BOLD, text_color=ACCENT_TEAL)
        self.lbl_zoom.pack(side="right", padx=12)

        ctk.CTkButton(top_bar, text="Zoom +", width=70, height=30, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY,
                      command=lambda: self._zoom(1.2)).pack(side="right", padx=4)
        ctk.CTkButton(top_bar, text="Zoom -", width=70, height=30, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY,
                      command=lambda: self._zoom(0.8)).pack(side="right", padx=4)
        ctk.CTkButton(top_bar, text="Vừa Khung (Fit)", width=110, height=30, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY,
                      command=self._fit_to_screen).pack(side="right", padx=4)

        # Canvas for Pan & Zoom
        self.canvas = tk.Canvas(self, bg="#0B0F1A", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<ButtonPress-1>", self._start_pan)
        self.canvas.bind("<B1-Motion>", self._do_pan)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        self._pan_x = 0
        self._pan_y = 0
        self._tk_img = None

        self._load_image_async()

    def _load_image_async(self):
        def _fetch():
            try:
                import urllib.request
                req = urllib.request.Request(self.image_url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=15, context=svc.get_ssl_context()) as resp:
                    raw_data = resp.read()
                pil_img = Image.open(io.BytesIO(raw_data))
                self.original_pil = pil_img
                self.after(0, self._fit_to_screen)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Lỗi Tải Ảnh", f"Không thể tải ảnh từ Cloud:\n{e}"))

        threading.Thread(target=_fetch, daemon=True).start()

    def _fit_to_screen(self):
        if not self.original_pil:
            return
        cw = max(self.canvas.winfo_width(), 400)
        ch = max(self.canvas.winfo_height(), 400)
        iw, ih = self.original_pil.size
        scale = min((cw - 40) / iw, (ch - 40) / ih, 1.0)
        self.zoom_level = scale
        self._pan_x = (cw - int(iw * scale)) // 2
        self._pan_y = (ch - int(ih * scale)) // 2
        self._render()

    def _zoom(self, factor):
        if not self.original_pil:
            return
        new_zoom = max(0.1, min(5.0, self.zoom_level * factor))
        self.zoom_level = new_zoom
        self._render()

    def _on_mousewheel(self, event):
        factor = 1.15 if event.delta > 0 else 0.85
        self._zoom(factor)

    def _start_pan(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def _do_pan(self, event):
        dx = event.x - self._drag_start_x
        dy = event.y - self._drag_start_y
        self._drag_start_x = event.x
        self._drag_start_y = event.y
        self._pan_x += dx
        self._pan_y += dy
        self.canvas.delete("img")
        self.canvas.create_image(self._pan_x, self._pan_y, anchor="nw", image=self._tk_img, tags="img")

    def _render(self):
        if not self.original_pil:
            return
        iw, ih = self.original_pil.size
        nw = max(10, int(iw * self.zoom_level))
        nh = max(10, int(ih * self.zoom_level))

        resized = self.original_pil.resize((nw, nh), Image.Resampling.BILINEAR)
        self._tk_img = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.create_image(self._pan_x, self._pan_y, anchor="nw", image=self._tk_img, tags="img")
        self.lbl_zoom.configure(text=f"{int(self.zoom_level * 100)}%")


# ─────────────────────────────────────────────────────────────────────────────
# DIALOG: CREATE / EDIT ANOMALY REPORT (18 COLUMNS)
# ─────────────────────────────────────────────────────────────────────────────
class AnomalyEditDialog(ctk.CTkToplevel):
    def __init__(self, parent, report_data: dict = None, on_saved = None):
        super().__init__(parent)
        self.is_edit = report_data is not None
        self.title("Chỉnh Sửa Báo Cáo Bất Thường" if self.is_edit else "Tạo Báo Cáo Bất Thường Mới")
        self.geometry("900x720")
        self.minsize(780, 600)
        self.configure(fg_color=BG_DEEP)
        self.transient(parent)
        self.grab_set()

        self.report_data = report_data or {}
        self.on_saved = on_saved
        self.selected_image_bytes = None
        self.selected_image_name = None
        self._thumb_tk = None

        self._build_ui()

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, fg_color=BG_CARD, height=54, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        title_text = "✏️ Chỉnh Sửa Báo Cáo Bất Thường" if self.is_edit else "➕ Tạo Báo Cáo Bất Thường Mới"
        ctk.CTkLabel(header, text=title_text, font=FONT_H1, text_color=TEXT_PRIMARY).pack(side="left", padx=20)

        # Scrollable form body
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=12)

        # 2-Column Grid
        col_left = ctk.CTkFrame(scroll, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=BORDER_CLR)
        col_left.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=4)

        col_right = ctk.CTkFrame(scroll, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=BORDER_CLR)
        col_right.pack(side="right", fill="both", expand=True, padx=(8, 0), pady=4)

        pad = {"padx": 16, "pady": (4, 4)}

        # ── LEFT COLUMN (Basic Metrics & Specs) ──
        ctk.CTkLabel(col_left, text="📌 Thông Tin Chung & Số Lượng", font=FONT_H2, text_color=ACCENT_TEAL).pack(anchor="w", padx=16, pady=(12, 6))

        # Date
        ctk.CTkLabel(col_left, text="Ngày Tháng (YYYY-MM-DD):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_date = ctk.CTkEntry(col_left, height=34, font=FONT_BODY)
        self.ent_date.pack(fill="x", **pad)
        self.ent_date.insert(0, str(self.report_data.get("report_date") or datetime.now().strftime("%Y-%m-%d")))

        # Process
        ctk.CTkLabel(col_left, text="Công Đoạn:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.opt_process = ctk.CTkComboBox(col_left, values=["CHA", "SCP", "RAD", "CHP", "SMT", "Lắp Ráp", "KCS / Final QC", "Khác"],
                                          height=34, font=FONT_BODY)
        self.opt_process.pack(fill="x", **pad)
        if self.report_data.get("process"):
            self.opt_process.set(self.report_data.get("process"))

        # Product
        ctk.CTkLabel(col_left, text="Sản Phẩm (Model / PWB):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_product = ctk.CTkEntry(col_left, height=34, font=FONT_BODY)
        self.ent_product.pack(fill="x", **pad)
        self.ent_product.insert(0, str(self.report_data.get("product_name") or ""))

        # Machine
        ctk.CTkLabel(col_left, text="Máy Móc / Chuyền / Thiết Bị:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_machine = ctk.CTkEntry(col_left, height=34, font=FONT_BODY)
        self.ent_machine.pack(fill="x", **pad)
        self.ent_machine.insert(0, str(self.report_data.get("machine") or ""))

        # Quantities & Defect Rate row
        q_frame = ctk.CTkFrame(col_left, fg_color="transparent")
        q_frame.pack(fill="x", padx=16, pady=4)

        # Total Qty
        f_tot = ctk.CTkFrame(q_frame, fg_color="transparent")
        f_tot.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkLabel(f_tot, text="Số Lượng KT:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w")
        self.ent_total_qty = ctk.CTkEntry(f_tot, height=34, font=FONT_BODY)
        self.ent_total_qty.pack(fill="x")
        self.ent_total_qty.insert(0, str(self.report_data.get("total_qty", 0)))
        self.ent_total_qty.bind("<KeyRelease>", self._calc_rate)

        # Defect Qty
        f_def = ctk.CTkFrame(q_frame, fg_color="transparent")
        f_def.pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkLabel(f_def, text="Số Lượng Lỗi:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w")
        self.ent_defect_qty = ctk.CTkEntry(f_def, height=34, font=FONT_BODY)
        self.ent_defect_qty.pack(fill="x")
        self.ent_defect_qty.insert(0, str(self.report_data.get("defect_qty", 0)))
        self.ent_defect_qty.bind("<KeyRelease>", self._calc_rate)

        # Defect Rate (%)
        f_rate = ctk.CTkFrame(q_frame, fg_color="transparent")
        f_rate.pack(side="left", fill="x", expand=True, padx=(4, 0))
        ctk.CTkLabel(f_rate, text="Tỷ Lệ Lỗi (%):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w")
        self.lbl_rate_val = ctk.CTkLabel(f_rate, text="0.00%", height=34, font=FONT_H2,
                                         fg_color=BG_SURFACE, corner_radius=6, text_color=ACCENT_AMBER)
        self.lbl_rate_val.pack(fill="x")

        # Responsible & PIC
        ctk.CTkLabel(col_left, text="Người Chịu Trách Nhiệm (Manager/Leader):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_resp = ctk.CTkEntry(col_left, height=34, font=FONT_BODY)
        self.ent_resp.pack(fill="x", **pad)
        self.ent_resp.insert(0, str(self.report_data.get("responsible_person") or ""))

        ctk.CTkLabel(col_left, text="Người Phụ Trách (Inspector/PIC):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_pic = ctk.CTkEntry(col_left, height=34, font=FONT_BODY)
        self.ent_pic.pack(fill="x", **pad)
        self.ent_pic.insert(0, str(self.report_data.get("pic") or ""))

        # Progress
        ctk.CTkLabel(col_left, text="Tiến Độ:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.opt_progress = ctk.CTkOptionMenu(col_left, values=["Chưa thực hiện", "Đang thực hiện", "Đã hoàn thành"],
                                             height=34, font=FONT_BODY)
        self.opt_progress.pack(fill="x", padx=16, pady=(4, 16))
        if self.report_data.get("progress"):
            self.opt_progress.set(self.report_data.get("progress"))
        else:
            self.opt_progress.set("Đang thực hiện")

        # ── RIGHT COLUMN (Descriptions & Image) ──
        ctk.CTkLabel(col_right, text="📝 Phân Tích & Biện Pháp Khắc Phục", font=FONT_H2, text_color=ACCENT_TEAL).pack(anchor="w", padx=16, pady=(12, 6))

        # Description
        ctk.CTkLabel(col_right, text="Mô Tả Lỗi (Hiện tượng bất thường):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.txt_desc = ctk.CTkTextbox(col_right, height=65, font=FONT_BODY)
        self.txt_desc.pack(fill="x", **pad)
        self.txt_desc.insert("1.0", str(self.report_data.get("description") or ""))

        # Root Cause
        ctk.CTkLabel(col_right, text="Nguyên Nhân:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.txt_cause = ctk.CTkTextbox(col_right, height=65, font=FONT_BODY)
        self.txt_cause.pack(fill="x", **pad)
        self.txt_cause.insert("1.0", str(self.report_data.get("root_cause") or ""))

        # Countermeasures
        ctk.CTkLabel(col_right, text="Các Biện Pháp Cải Tiến:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.txt_counter = ctk.CTkTextbox(col_right, height=65, font=FONT_BODY)
        self.txt_counter.pack(fill="x", **pad)
        self.txt_counter.insert("1.0", str(self.report_data.get("countermeasures") or ""))

        # SOP & Notes
        ctk.CTkLabel(col_right, text="Tiêu Chuẩn Hóa SOP:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_sop = ctk.CTkEntry(col_right, height=34, font=FONT_BODY)
        self.ent_sop.pack(fill="x", **pad)
        self.ent_sop.insert(0, str(self.report_data.get("sop_standard") or ""))

        ctk.CTkLabel(col_right, text="Ghi Chú:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_notes = ctk.CTkEntry(col_right, height=34, font=FONT_BODY)
        self.ent_notes.pack(fill="x", **pad)
        self.ent_notes.insert(0, str(self.report_data.get("notes") or ""))

        # Image section
        ctk.CTkLabel(col_right, text="🖼️ Hình Ảnh Lỗi (Click để chọn hoặc bấm Ctrl+V để dán):",
                     font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", padx=16, pady=(6, 2))

        img_box = ctk.CTkFrame(col_right, fg_color=BG_SURFACE, corner_radius=8, height=110)
        img_box.pack(fill="x", padx=16, pady=(2, 16))
        img_box.pack_propagate(False)

        self.lbl_img_preview = ctk.CTkLabel(img_box, text="[Chưa có ảnh]\nNhấn nút bên cạnh hoặc Ctrl+V",
                                            font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_img_preview.pack(side="left", padx=12, fill="both", expand=True)

        btn_img_row = ctk.CTkFrame(img_box, fg_color="transparent")
        btn_img_row.pack(side="right", padx=10, pady=8)

        ctk.CTkButton(btn_img_row, text="📁 Chọn Ảnh", width=100, height=32, font=FONT_SMALL,
                      command=self._pick_image).pack(pady=3)
        ctk.CTkButton(btn_img_row, text="📋 Dán (Ctrl+V)", width=100, height=32, font=FONT_SMALL,
                      command=self._paste_image).pack(pady=3)
        self.btn_clear_img = ctk.CTkButton(btn_img_row, text="✕ Xóa ảnh", width=100, height=28,
                                           font=FONT_SMALL, fg_color=BG_CARD, text_color=ACCENT_RED,
                                           command=self._clear_image)
        self.btn_clear_img.pack(pady=3)

        # Global clipboard bind
        self.bind("<Control-v>", lambda e: self._paste_image())
        self.bind("<Control-V>", lambda e: self._paste_image())

        # Load existing image if edit
        if self.report_data.get("image_url"):
            self._load_remote_thumb(self.report_data.get("image_url"))

        self._calc_rate()

        # Bottom Bar
        bottom = ctk.CTkFrame(self, fg_color=BG_CARD, height=54, corner_radius=0)
        bottom.pack(fill="x", side="bottom")
        bottom.pack_propagate(False)

        self.lbl_saving = ctk.CTkLabel(bottom, text="", font=FONT_SMALL, text_color=ACCENT_TEAL)
        self.lbl_saving.pack(side="left", padx=20)

        ctk.CTkButton(bottom, text="Hủy", width=100, height=36, corner_radius=8,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self.destroy).pack(side="right", padx=(4, 20))

        self.btn_save = ctk.CTkButton(bottom, text="💾 Lưu Báo Cáo", width=150, height=36, corner_radius=8,
                                      fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
                                      font=FONT_H2, command=self._save)
        self.btn_save.pack(side="right", padx=6)

    def _calc_rate(self, event=None):
        try:
            tot = float(self.ent_total_qty.get().strip() or 0)
            def_q = float(self.ent_defect_qty.get().strip() or 0)
            if tot > 0:
                rate = (def_q / tot) * 100.0
                self.lbl_rate_val.configure(text=f"{rate:.2f}%")
                if rate > 5.0:
                    self.lbl_rate_val.configure(text_color=ACCENT_RED)
                elif rate > 1.0:
                    self.lbl_rate_val.configure(text_color=ACCENT_AMBER)
                else:
                    self.lbl_rate_val.configure(text_color=ACCENT_TEAL)
            else:
                self.lbl_rate_val.configure(text="0.00%", text_color=TEXT_MUTED)
        except Exception:
            self.lbl_rate_val.configure(text="0.00%", text_color=TEXT_MUTED)

    def _pick_image(self):
        f = filedialog.askopenfilename(
            title="Chọn Ảnh Lỗi",
            filetypes=[("Hình ảnh", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Tất cả", "*.*")]
        )
        if f and os.path.exists(f):
            try:
                with open(f, "rb") as fp:
                    self.selected_image_bytes = fp.read()
                self.selected_image_name = os.path.basename(f)
                self._show_thumb_from_bytes(self.selected_image_bytes)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file ảnh:\n{e}")

    def _paste_image(self):
        try:
            from PIL import ImageGrab
            clip_img = ImageGrab.grabclipboard()
            if isinstance(clip_img, Image.Image):
                buf = io.BytesIO()
                clip_img.convert("RGB").save(buf, format="JPEG", quality=90)
                self.selected_image_bytes = buf.getvalue()
                self.selected_image_name = "pasted_defect.jpg"
                self._show_thumb_from_bytes(self.selected_image_bytes)
            elif isinstance(clip_img, list) and clip_img:
                # File path in clipboard
                p = clip_img[0]
                if os.path.isfile(p):
                    with open(p, "rb") as fp:
                        self.selected_image_bytes = fp.read()
                    self.selected_image_name = os.path.basename(p)
                    self._show_thumb_from_bytes(self.selected_image_bytes)
            else:
                messagebox.showinfo("Clipboard", "Không tìm thấy ảnh trong Clipboard (Hãy chụp ảnh màn hình bằng PrintScreen/Snipping Tool rồi thử lại).")
        except Exception as e:
            messagebox.showwarning("Lỗi Clipboard", f"Không thể dán ảnh:\n{e}")

    def _show_thumb_from_bytes(self, b: bytes):
        try:
            im = Image.open(io.BytesIO(b))
            im.thumbnail((140, 95))
            self._thumb_tk = ImageTk.PhotoImage(im)
            self.lbl_img_preview.configure(image=self._thumb_tk, text="")
        except Exception as e:
            self.lbl_img_preview.configure(text="[Ảnh đã chọn]")

    def _load_remote_thumb(self, url: str):
        def _fetch():
            try:
                import urllib.request
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=10, context=svc.get_ssl_context()) as r:
                    raw = r.read()
                im = Image.open(io.BytesIO(raw))
                im.thumbnail((140, 95))
                self._thumb_tk = ImageTk.PhotoImage(im)
                self.after(0, lambda: self.lbl_img_preview.configure(image=self._thumb_tk, text=""))
            except Exception:
                self.after(0, lambda: self.lbl_img_preview.configure(text="[Ảnh trên Cloud]"))

        threading.Thread(target=_fetch, daemon=True).start()

    def _clear_image(self):
        self.selected_image_bytes = None
        self.selected_image_name = None
        self.report_data["image_url"] = None
        self.lbl_img_preview.configure(image="", text="[Đã xóa ảnh]")

    def _save(self):
        # Basic validation
        prod = self.ent_product.get().strip()
        if not prod:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập Tên Sản Phẩm!")
            self.ent_product.focus()
            return

        try:
            tot_q = int(self.ent_total_qty.get().strip() or 0)
            def_q = int(self.ent_defect_qty.get().strip() or 0)
        except ValueError:
            messagebox.showwarning("Lỗi số lượng", "Số lượng kiểm tra và số lượng lỗi phải là số nguyên!")
            return

        payload = {
            "report_date": self.ent_date.get().strip() or datetime.now().strftime("%Y-%m-%d"),
            "process": self.opt_process.get().strip(),
            "product_name": prod,
            "machine": self.ent_machine.get().strip(),
            "total_qty": tot_q,
            "defect_qty": def_q,
            "responsible_person": self.ent_resp.get().strip(),
            "pic": self.ent_pic.get().strip(),
            "progress": self.opt_progress.get().strip(),
            "description": self.txt_desc.get("1.0", "end-1c").strip(),
            "root_cause": self.txt_cause.get("1.0", "end-1c").strip(),
            "countermeasures": self.txt_counter.get("1.0", "end-1c").strip(),
            "sop_standard": self.ent_sop.get().strip(),
            "notes": self.ent_notes.get().strip(),
        }
        if "image_url" in self.report_data:
            payload["image_url"] = self.report_data["image_url"]

        self.btn_save.configure(state="disabled")
        self.lbl_saving.configure(text="⏳ Đang lưu lên Cloud Supabase...")

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
                self.after(0, lambda: self._on_save_success(res))
            except Exception as e:
                self.after(0, lambda: self._on_save_error(str(e)))

        threading.Thread(target=_do_save, daemon=True).start()

    def _on_save_success(self, res):
        self.destroy()
        if self.on_saved:
            self.on_saved(res)

    def _on_save_error(self, err_msg):
        self.btn_save.configure(state="normal")
        self.lbl_saving.configure(text="")
        messagebox.showerror("Lỗi Lưu Dữ Liệu", f"Không thể lưu lên Supabase:\n{err_msg}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN VIEW: ANOMALY REPORT VIEW (CTK FRAME)
# ─────────────────────────────────────────────────────────────────────────────
class AnomalyReportView(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app

        self.all_reports: list[dict] = []
        self.filtered_reports: list[dict] = []
        self.selected_report: dict = None
        self.is_admin_session = False
        self._current_thumb_tk = None

        self._build_ui()
        self.refresh_data()

    def _build_ui(self):
        # ── TOP ACTION & FILTER BAR ──────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                               border_width=1, border_color=BORDER_CLR, height=64)
        top_bar.pack(fill="x", pady=(0, 10))
        top_bar.pack_propagate(False)

        # Search box
        self.ent_search = ctk.CTkEntry(top_bar, placeholder_text="🔍 Tìm kiếm sản phẩm, công đoạn, người phụ trách, lỗi...",
                                       width=300, height=36, corner_radius=8, font=FONT_BODY)
        self.ent_search.pack(side="left", padx=(16, 8), pady=14)
        self.ent_search.bind("<KeyRelease>", lambda e: self._apply_filter())

        # Process Filter
        self.opt_filter_proc = ctk.CTkOptionMenu(
            top_bar,
            values=["Tất cả công đoạn", "CHA", "SCP", "RAD", "CHP", "SMT", "Lắp Ráp", "KCS / Final QC"],
            height=36, corner_radius=8, font=FONT_SMALL,
            command=lambda v: self._apply_filter()
        )
        self.opt_filter_proc.pack(side="left", padx=4)

        # Progress Filter
        self.opt_filter_prog = ctk.CTkOptionMenu(
            top_bar,
            values=["Tất cả tiến độ", "Chưa thực hiện", "Đang thực hiện", "Đã hoàn thành"],
            height=36, corner_radius=8, font=FONT_SMALL,
            command=lambda v: self._apply_filter()
        )
        self.opt_filter_prog.pack(side="left", padx=4)

        # Right-side action buttons
        self.btn_admin_lock = ctk.CTkButton(
            top_bar, text="🔒 Admin: Đã Khóa", width=125, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._toggle_admin_lock
        )
        self.btn_admin_lock.pack(side="right", padx=(4, 16))

        self.btn_export = ctk.CTkButton(
            top_bar, text="📊 Xuất Excel", width=105, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._export_excel
        )
        self.btn_export.pack(side="right", padx=4)

        self.btn_refresh = ctk.CTkButton(
            top_bar, text="🔄 Làm Mới", width=95, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self.refresh_data
        )
        self.btn_refresh.pack(side="right", padx=4)

        self.btn_add = ctk.CTkButton(
            top_bar, text="➕ Thêm Báo Cáo", width=135, height=36, corner_radius=8,
            fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
            font=FONT_H2, command=self._on_add_report
        )
        self.btn_add.pack(side="right", padx=4)

        # ── WORKSPACE SPLIT (Left: Table | Right: Detail Inspection) ─────────
        workspace = ctk.CTkFrame(self, fg_color="transparent")
        workspace.pack(fill="both", expand=True)

        # Left Panel (Table)
        p_left = ctk.CTkFrame(workspace, fg_color=BG_CARD, corner_radius=12,
                              border_width=1, border_color=BORDER_CLR)
        p_left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # Status row above tree
        tbl_info = ctk.CTkFrame(p_left, fg_color="transparent", height=32)
        tbl_info.pack(fill="x", padx=14, pady=(10, 4))

        self.lbl_tbl_count = ctk.CTkLabel(tbl_info, text="Danh Sách Báo Cáo: 0 bản ghi",
                                          font=FONT_H2, text_color=TEXT_PRIMARY)
        self.lbl_tbl_count.pack(side="left")

        # Treeview
        tree_container = tk.Frame(p_left, bg="#131929")
        tree_container.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        cols = ("stt", "date", "process", "product", "machine", "tot", "def", "rate", "pic", "prog", "img")
        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings",
                                 style="Anomaly.Treeview", selectmode="browse")

        self.tree.heading("stt", text="STT")
        self.tree.heading("date", text="Ngày")
        self.tree.heading("process", text="Công Đoạn")
        self.tree.heading("product", text="Sản Phẩm")
        self.tree.heading("machine", text="Máy Móc")
        self.tree.heading("tot", text="SL Kiểm")
        self.tree.heading("def", text="SL Lỗi")
        self.tree.heading("rate", text="Tỷ Lệ (%)")
        self.tree.heading("pic", text="Phụ Trách")
        self.tree.heading("prog", text="Tiến Độ")
        self.tree.heading("img", text="Ảnh")

        self.tree.column("stt", width=42, anchor="center")
        self.tree.column("date", width=85, anchor="center")
        self.tree.column("process", width=80, anchor="center")
        self.tree.column("product", width=120, anchor="w")
        self.tree.column("machine", width=75, anchor="center")
        self.tree.column("tot", width=65, anchor="e")
        self.tree.column("def", width=60, anchor="e")
        self.tree.column("rate", width=70, anchor="e")
        self.tree.column("pic", width=95, anchor="w")
        self.tree.column("prog", width=95, anchor="center")
        self.tree.column("img", width=45, anchor="center")

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)

        # Right Panel (Inspection Card)
        self.p_right = ctk.CTkFrame(workspace, fg_color=BG_CARD, corner_radius=12,
                                    border_width=1, border_color=BORDER_CLR, width=340)
        self.p_right.pack(side="right", fill="y", padx=(0, 0))
        self.p_right.pack_propagate(False)
        self._build_detail_panel(self.p_right)

    def _build_detail_panel(self, parent):
        # Header
        top_d = ctk.CTkFrame(parent, fg_color="transparent")
        top_d.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(top_d, text="🔍 Chi Tiết Báo Cáo", font=FONT_H1, text_color=ACCENT_TEAL).pack(side="left")

        # Image preview box
        self.img_card = ctk.CTkFrame(parent, fg_color=BG_SURFACE, corner_radius=8, height=160)
        self.img_card.pack(fill="x", padx=16, pady=(0, 8))
        self.img_card.pack_propagate(False)

        self.lbl_detail_thumb = ctk.CTkLabel(self.img_card, text="[Chọn 1 báo cáo để xem chi tiết]",
                                             font=FONT_BODY, text_color=TEXT_MUTED)
        self.lbl_detail_thumb.pack(fill="both", expand=True, padx=8, pady=8)
        self.lbl_detail_thumb.bind("<Button-1>", lambda e: self._open_image_viewer())

        self.btn_view_full_img = ctk.CTkButton(
            parent, text="🔍 Phóng To Ảnh (Zoom In)", height=30, font=FONT_SMALL,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            command=self._open_image_viewer
        )
        self.btn_view_full_img.pack(fill="x", padx=16, pady=(0, 8))

        # Details scrollable container
        d_scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        d_scroll.pack(fill="both", expand=True, padx=12, pady=4)

        self.lbl_d_product = ctk.CTkLabel(d_scroll, text="Sản phẩm: --", font=FONT_H2, text_color=TEXT_PRIMARY, anchor="w")
        self.lbl_d_product.pack(fill="x", pady=2)

        self.lbl_d_meta = ctk.CTkLabel(d_scroll, text="Công đoạn: --  |  Máy: --", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w")
        self.lbl_d_meta.pack(fill="x", pady=2)

        self.lbl_d_stats = ctk.CTkLabel(d_scroll, text="SL KT: 0  |  SL Lỗi: 0  |  Tỷ lệ: 0%", font=FONT_BOLD, text_color=ACCENT_AMBER, anchor="w")
        self.lbl_d_stats.pack(fill="x", pady=2)

        self.lbl_d_people = ctk.CTkLabel(d_scroll, text="QL: --  |  PIC: --", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w")
        self.lbl_d_people.pack(fill="x", pady=2)

        ctk.CTkFrame(d_scroll, fg_color=BORDER_CLR, height=1).pack(fill="x", pady=6)

        ctk.CTkLabel(d_scroll, text="Mô tả hiện tượng lỗi:", font=FONT_BOLD, text_color=ACCENT_TEAL, anchor="w").pack(fill="x")
        self.lbl_d_desc = ctk.CTkLabel(d_scroll, text="--", font=FONT_BODY, text_color=TEXT_PRIMARY,
                                       anchor="w", justify="left", wraplength=280)
        self.lbl_d_desc.pack(fill="x", pady=(2, 6))

        ctk.CTkLabel(d_scroll, text="Nguyên nhân:", font=FONT_BOLD, text_color=ACCENT_TEAL, anchor="w").pack(fill="x")
        self.lbl_d_cause = ctk.CTkLabel(d_scroll, text="--", font=FONT_BODY, text_color=TEXT_PRIMARY,
                                        anchor="w", justify="left", wraplength=280)
        self.lbl_d_cause.pack(fill="x", pady=(2, 6))

        ctk.CTkLabel(d_scroll, text="Biện pháp cải tiến:", font=FONT_BOLD, text_color=ACCENT_TEAL, anchor="w").pack(fill="x")
        self.lbl_d_counter = ctk.CTkLabel(d_scroll, text="--", font=FONT_BODY, text_color=TEXT_PRIMARY,
                                          anchor="w", justify="left", wraplength=280)
        self.lbl_d_counter.pack(fill="x", pady=(2, 6))

        self.lbl_d_sop = ctk.CTkLabel(d_scroll, text="SOP: --", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w")
        self.lbl_d_sop.pack(fill="x", pady=2)

        self.lbl_d_notes = ctk.CTkLabel(d_scroll, text="Ghi chú: --", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w")
        self.lbl_d_notes.pack(fill="x", pady=2)

        # Bottom Actions on Selected Item
        act_row = ctk.CTkFrame(parent, fg_color="transparent")
        act_row.pack(fill="x", padx=16, pady=(8, 14))

        self.btn_edit = ctk.CTkButton(
            act_row, text="✏️ Chỉnh Sửa", height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._on_edit_report
        )
        self.btn_edit.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_delete = ctk.CTkButton(
            act_row, text="🗑️ Xóa", height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._on_delete_report
        )
        self.btn_delete.pack(side="right", fill="x", expand=True, padx=(4, 0))

    # ── DATA FETCH & POPULATE ────────────────────────────────────────────────
    def refresh_data(self):
        """Fetches reports from Supabase asynchronously."""
        self.btn_refresh.configure(state="disabled", text="⏳ Đang tải...")

        def _fetch():
            try:
                reps = svc.fetch_all_reports()
                self.after(0, lambda: self._on_fetch_success(reps))
            except Exception as e:
                self.after(0, lambda: self._on_fetch_error(str(e)))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_fetch_success(self, reports: list[dict]):
        self.all_reports = reports
        self.btn_refresh.configure(state="normal", text="🔄 Làm Mới")
        self._apply_filter()
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(f"✅ Đã đồng bộ {len(reports)} báo cáo từ Cloud Supabase!")

    def _on_fetch_error(self, err: str):
        self.btn_refresh.configure(state="normal", text="🔄 Làm Mới")
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(f"⚠️ Lỗi kết nối Cloud: {err[:50]}...")

    def _apply_filter(self):
        q = self.ent_search.get().strip().lower()
        f_proc = self.opt_filter_proc.get()
        f_prog = self.opt_filter_prog.get()

        res = []
        for r in self.all_reports:
            # Process filter
            if f_proc != "Tất cả công đoạn" and str(r.get("process")).lower() != f_proc.lower():
                continue
            # Progress filter
            if f_prog != "Tất cả tiến độ" and str(r.get("progress")).lower() != f_prog.lower():
                continue
            # Text search filter
            if q:
                searchable = f"{r.get('product_name', '')} {r.get('process', '')} {r.get('machine', '')} {r.get('pic', '')} {r.get('responsible_person', '')} {r.get('description', '')}".lower()
                if q not in searchable:
                    continue
            res.append(r)

        self.filtered_reports = res
        self._populate_tree(res)

    def _populate_tree(self, reports: list[dict]):
        self.tree.delete(*self.tree.get_children())
        self.lbl_tbl_count.configure(text=f"Danh Sách Báo Cáo: {len(reports)} bản ghi")

        for idx, r in enumerate(reports, 1):
            rate_val = r.get("defect_rate")
            rate_str = f"{rate_val:.2f}%" if rate_val is not None else "0.00%"
            has_img = "📷" if r.get("image_url") else ""

            tag = "even" if idx % 2 == 0 else "odd"
            prog = str(r.get("progress") or "")
            if "hoàn thành" in prog.lower():
                tag = "tag_done"
            elif "chưa" in prog.lower():
                tag = "tag_open"

            vals = (
                idx,
                r.get("report_date", ""),
                r.get("process", ""),
                r.get("product_name", ""),
                r.get("machine", ""),
                r.get("total_qty", 0),
                r.get("defect_qty", 0),
                rate_str,
                r.get("pic", ""),
                r.get("progress", "Đang thực hiện"),
                has_img
            )
            self.tree.insert("", "end", iid=str(r.get("id")), values=vals, tags=(tag,))

    def _on_row_select(self, event):
        selected_ids = self.tree.selection()
        if not selected_ids:
            return
        rep_id = selected_ids[0]
        found = [r for r in self.all_reports if str(r.get("id")) == str(rep_id)]
        if found:
            self.selected_report = found[0]
            self._display_detail(self.selected_report)

    def _display_detail(self, r: dict):
        self.lbl_d_product.configure(text=f"Sản phẩm: {r.get('product_name', '--')}")
        self.lbl_d_meta.configure(text=f"Công đoạn: {r.get('process', '--')}  |  Máy: {r.get('machine', '--')}")

        rate_val = r.get("defect_rate", 0.0)
        rate_str = f"{rate_val:.2f}%" if rate_val is not None else "0.00%"
        self.lbl_d_stats.configure(text=f"SL KT: {r.get('total_qty', 0):,}  |  SL Lỗi: {r.get('defect_qty', 0):,}  |  Tỷ lệ: {rate_str}")

        self.lbl_d_people.configure(text=f"QL: {r.get('responsible_person', '--')}  |  PIC: {r.get('pic', '--')}")
        self.lbl_d_desc.configure(text=r.get("description") or "--")
        self.lbl_d_cause.configure(text=r.get("root_cause") or "--")
        self.lbl_d_counter.configure(text=r.get("countermeasures") or "--")
        self.lbl_d_sop.configure(text=f"SOP: {r.get('sop_standard', '--')}")
        self.lbl_d_notes.configure(text=f"Ghi chú: {r.get('notes', '--')}")

        img_url = r.get("image_url")
        if img_url:
            self.lbl_detail_thumb.configure(image="", text="⏳ Đang tải ảnh...")
            self.btn_view_full_img.configure(state="normal")
            self._fetch_thumb_async(img_url)
        else:
            self.lbl_detail_thumb.configure(image="", text="[Báo cáo này không có ảnh]")
            self.btn_view_full_img.configure(state="disabled")

    def _fetch_thumb_async(self, url: str):
        def _fetch():
            try:
                import urllib.request
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=10, context=svc.get_ssl_context()) as resp:
                    raw = resp.read()
                im = Image.open(io.BytesIO(raw))
                im.thumbnail((260, 140))
                self._current_thumb_tk = ImageTk.PhotoImage(im)
                self.after(0, lambda: self.lbl_detail_thumb.configure(image=self._current_thumb_tk, text=""))
            except Exception:
                self.after(0, lambda: self.lbl_detail_thumb.configure(text="[Ảnh trên Cloud - Click để mở]"))

        threading.Thread(target=_fetch, daemon=True).start()

    def _open_image_viewer(self):
        if self.selected_report and self.selected_report.get("image_url"):
            ImageZoomViewer(self, self.selected_report["image_url"],
                            title=f"Ảnh Lỗi: {self.selected_report.get('product_name')} ({self.selected_report.get('process')})")

    # ── ADMIN & PASSWORD VERIFICATION ────────────────────────────────────────
    def _require_admin(self, callback):
        """Executes callback if already unlocked in session, otherwise prompts for password."""
        if self.is_admin_session:
            callback()
        else:
            def _on_ok():
                self.is_admin_session = True
                self.btn_admin_lock.configure(text="🔓 Admin: Đã Mở", fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"))
                callback()
            PasswordDialog(self, on_success=_on_ok)

    def _toggle_admin_lock(self):
        if self.is_admin_session:
            # Offer to lock or change password
            m = messagebox.askyesnocancel("Quản Trị", "Bạn đang ở chế độ Quản Trị.\n\n• Chọn YES để Khóa lại\n• Chọn NO để Đổi Mật Khẩu\n• Chọn CANCEL để đóng")
            if m is True:
                self.is_admin_session = False
                self.btn_admin_lock.configure(text="🔒 Admin: Đã Khóa", fg_color=BG_SURFACE, text_color=TEXT_PRIMARY)
                messagebox.showinfo("Đã khóa", "Đã khóa phiên Quản Trị thành công.")
            elif m is False:
                ChangePasswordDialog(self)
        else:
            def _on_ok():
                self.is_admin_session = True
                self.btn_admin_lock.configure(text="🔓 Admin: Đã Mở", fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"))
                messagebox.showinfo("Mở Khóa", "✅ Mở khóa quyền Quản Trị thành công!")
            PasswordDialog(self, on_success=_on_ok)

    # ── CRUD ACTIONS ─────────────────────────────────────────────────────────
    def _on_add_report(self):
        self._require_admin(lambda: AnomalyEditDialog(self, on_saved=self._on_report_created))

    def _on_report_created(self, new_rep):
        self.refresh_data()
        if hasattr(self.app, "show_toast"):
            self.app.show_toast("✅ Đã thêm báo cáo bất thường mới lên Supabase!")

    def _on_edit_report(self):
        if not self.selected_report:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một báo cáo trên bảng để chỉnh sửa!")
            return
        self._require_admin(lambda: AnomalyEditDialog(self, report_data=self.selected_report, on_saved=self._on_report_updated))

    def _on_report_updated(self, updated_rep):
        self.refresh_data()
        if hasattr(self.app, "show_toast"):
            self.app.show_toast(f"✅ Đã cập nhật báo cáo #{updated_rep.get('id')} thành công!")

    def _on_delete_report(self):
        if not self.selected_report:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một báo cáo trên bảng để xóa!")
            return

        rep = self.selected_report
        def _do_del():
            confirm = messagebox.askyesno(
                "Xác Nhận Xóa",
                f"Bạn có chắc chắn muốn xóa vĩnh viễn báo cáo này?\n\n• Sản phẩm: {rep.get('product_name')}\n• Công đoạn: {rep.get('process')}\n• Ngày: {rep.get('report_date')}"
            )
            if confirm:
                ok = svc.delete_report(rep["id"], rep.get("image_url"))
                if ok:
                    self.selected_report = None
                    self.refresh_data()
                    if hasattr(self.app, "show_toast"):
                        self.app.show_toast("🗑️ Đã xóa báo cáo và ảnh trên Cloud thành công!")
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa báo cáo trên Supabase!")

        self._require_admin(_do_del)

    def _export_excel(self):
        if not self.filtered_reports:
            messagebox.showwarning("Trống", "Không có bản ghi nào để xuất Excel!")
            return

        def_name = f"Bao_Cao_Bat_Thuong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        f = filedialog.asksaveasfilename(
            title="Lưu Báo Cáo Excel",
            initialfile=def_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if f:
            try:
                saved = svc.export_reports_to_excel(self.filtered_reports, f)
                ask_open = messagebox.askyesno("Xuất Thành Công", f"Đã xuất {len(self.filtered_reports)} báo cáo ra Excel:\n{saved}\n\nBạn có muốn mở file ngay không?")
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
                        rowheight=28,
                        font=("Segoe UI", 10),
                        borderwidth=0)
        style.configure("Anomaly.Treeview.Heading",
                        background=hdr_bg,
                        foreground=hdr_fg,
                        font=("Segoe UI", 9, "bold"),
                        padding=4)
        style.map("Anomaly.Treeview",
                  background=[("selected", sel_bg)],
                  foreground=[("selected", sel_fg)])

        self.tree.tag_configure("odd",  background=odd_bg)
        self.tree.tag_configure("even", background=even_bg)
        self.tree.tag_configure("tag_done", background="#064E3B" if is_dark else "#DCFCE7",
                                foreground="#A7F3D0" if is_dark else "#166534")
        self.tree.tag_configure("tag_open", background="#7F1D1D" if is_dark else "#FEE2E2",
                                foreground="#FECACA" if is_dark else "#991B1B")
