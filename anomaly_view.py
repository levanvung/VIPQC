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

FONT_H1     = ("Segoe UI", 13, "bold")
FONT_H2     = ("Segoe UI", 11, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_SMALL  = ("Segoe UI", 9)


def safe_after(widget, ms: int, callback):
    """Safely invokes callback on main thread if widget still exists."""
    try:
        if widget and widget.winfo_exists():
            widget.after(ms, callback)
    except Exception:
        pass


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
# ADVANCED IMAGE MARKING & ZOOM STUDIO (PHÓNG TO, XOAY, ZOOM, MARKING)
# ─────────────────────────────────────────────────────────────────────────────
class AnomalyImageStudioWindow(ctk.CTkToplevel):
    """
    Dedicated Studio for inspecting defect images with high-resolution pan/zoom,
    90-degree rotations, and complete marking/annotation tools (boxes, ovals, arrows, freehand, text).
    Allows saving the annotated image back to Supabase or downloading to local disk.
    """
    def __init__(self, parent, report_data: dict, initial_pil: Image.Image = None, on_image_updated=None):
        super().__init__(parent)
        self.report_data = report_data
        self.on_image_updated = on_image_updated
        self.image_url = report_data.get("image_url", "")

        prod = report_data.get("product_name", "Không rõ")
        proc = report_data.get("process", "")
        self.title(f"🖼️ Studio Soi & Đánh Dấu Ảnh Lỗi — {prod} ({proc})")
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
        # ── TOP TOOLBAR (2 ROWS) ─────────────────────────────────────────────
        bar = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0,
                           border_width=1, border_color=BORDER_CLR)
        bar.pack(fill="x", side="top")

        # ── ROW 1: MARKING TOOLS & STYLES (CÔNG CỤ ĐÁNH DẤU) ────────────────
        row1 = ctk.CTkFrame(bar, fg_color="transparent", height=42)
        row1.pack(fill="x", padx=12, pady=(6, 2))
        row1.pack_propagate(False)

        # Title & Info
        ctk.CTkLabel(row1, text="🎨 CÔNG CỤ VẼ & ĐÁNH DẤU:", font=FONT_H1, text_color=ACCENT_TEAL).pack(side="left", padx=(2, 10))

        # Tool selector buttons
        self.tool_buttons = {}
        tools = [
            ("pan", "✋ Di chuyển"),
            ("rect", "🟥 Hộp"),
            ("oval", "⭕ Tròn"),
            ("arrow", "➡️ Mũi tên"),
            ("pen", "✏️ Bút vẽ"),
            ("text", "🔤 Chữ"),
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
        ctk.CTkButton(row1, text="🗑️ Xóa hết nét", width=95, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
                      command=self._clear_annotations).pack(side="right", padx=(2, 0))

        ctk.CTkButton(row1, text="↩️ Hoàn tác", width=85, height=30, font=FONT_SMALL,
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
            row2, text="⛶ Toàn Màn Hình", width=140, height=30, font=FONT_H2,
            fg_color=BG_SURFACE, text_color=ACCENT_TEAL, hover_color=BG_HOVER,
            border_width=1, border_color=ACCENT_TEAL,
            command=self._toggle_fullscreen
        )
        self.btn_fullscreen.pack(side="left", padx=(2, 8))

        # Rotate Button
        ctk.CTkButton(row2, text="🔄 Xoay 90° (R)", width=105, height=30, font=FONT_SMALL,
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
        ctk.CTkButton(row2, text="🖼️ Vừa Khung", width=85, height=30, font=FONT_SMALL,
                      fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._fit_to_screen).pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(row2, text="100%", font=FONT_BOLD, text_color=ACCENT_TEAL, width=48)
        self.lbl_zoom.pack(side="left", padx=4)

        # Right side of Row 2: Close, Download, Save Cloud
        ctk.CTkButton(
            row2, text="✕ Đóng", width=75, height=30, font=FONT_SMALL,
            fg_color=BG_SURFACE, text_color=TEXT_MUTED, hover_color=BG_HOVER,
            command=self.destroy
        ).pack(side="right", padx=(2, 0))

        self.btn_download = ctk.CTkButton(
            row2, text="📥 Tải Về Máy", width=110, height=30, font=FONT_SMALL,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            command=self._download_to_pc
        )
        self.btn_download.pack(side="right", padx=4)

        self.btn_save_cloud = ctk.CTkButton(
            row2, text="💾 Lưu Lên Báo Cáo", width=155, height=30, font=FONT_H2,
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
        try:
            is_zoomed = self.state() == "zoomed" or bool(self.attributes("-fullscreen"))
            if is_zoomed:
                try:
                    self.attributes("-fullscreen", False)
                except Exception:
                    pass
                self.state("normal")
                self.btn_fullscreen.configure(text="⛶ Toàn Màn Hình", fg_color=BG_SURFACE)
            else:
                self.state("zoomed")
                self.btn_fullscreen.configure(text="🗗 Thu Nhỏ Cửa Sổ", fg_color=ACCENT_TEAL)

            self.after(150, self._fit_to_screen)
        except Exception as e:
            print(f"[Studio] Toggle fullscreen error: {e}")

    def _exit_fullscreen_if_active(self):
        try:
            if self.state() == "zoomed" or bool(self.attributes("-fullscreen")):
                try:
                    self.attributes("-fullscreen", False)
                except Exception:
                    pass
                self.state("normal")
                self.btn_fullscreen.configure(text="⛶ Toàn Màn Hình", fg_color=BG_SURFACE)
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

        self.is_edit = False
        self.report_data = {}
        self.selected_image_bytes = None
        self.selected_image_name = None
        self.current_pil_image: Image.Image = None
        self._ctk_preview = None

        self._build_ui()

    def _build_ui(self):
        self.pack_propagate(False)

        # ── DRAWER HEADER ──
        header = ctk.CTkFrame(self, fg_color=BG_SURFACE, height=52, corner_radius=10)
        header.pack(fill="x", padx=10, pady=(10, 6))
        header.pack_propagate(False)

        self.lbl_title = ctk.CTkLabel(header, text="➕ Thêm Báo Cáo Mới", font=FONT_H1, text_color=TEXT_PRIMARY)
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

        ctk.CTkLabel(c1, text="📌 Thông Tin Kiểm Tra & Số Lượng", font=FONT_H2, text_color=ACCENT_TEAL).pack(anchor="w", padx=14, pady=(10, 4))

        ctk.CTkLabel(c1, text="Ngày Tháng (YYYY-MM-DD):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_date = ctk.CTkEntry(c1, height=34, font=FONT_BODY)
        self.ent_date.pack(fill="x", **pad)

        ctk.CTkLabel(c1, text="Công Đoạn:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.opt_process = ctk.CTkOptionMenu(c1, values=["CHA", "SCP", "RAD", "CHP", "SMT", "Lắp Ráp", "KCS / Final QC", "Khác"],
                                             height=34, font=FONT_BODY)
        self.opt_process.pack(fill="x", **pad)

        ctk.CTkLabel(c1, text="Sản Phẩm (Model / PWB): *", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_product = ctk.CTkEntry(c1, height=34, font=FONT_BODY, placeholder_text="Nhập tên model hoặc mã PWB...")
        self.ent_product.pack(fill="x", **pad)

        ctk.CTkLabel(c1, text="Máy Móc / Chuyền / Thiết Bị:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_machine = ctk.CTkEntry(c1, height=34, font=FONT_BODY, placeholder_text="Ví dụ: Line 2, Máy 12...")
        self.ent_machine.pack(fill="x", **pad)

        # Quantities row
        q_row = ctk.CTkFrame(c1, fg_color="transparent")
        q_row.pack(fill="x", padx=14, pady=4)

        f1 = ctk.CTkFrame(q_row, fg_color="transparent")
        f1.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkLabel(f1, text="SL Kiểm:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w")
        self.ent_tot = ctk.CTkEntry(f1, height=34, font=FONT_BODY)
        self.ent_tot.pack(fill="x")
        self.ent_tot.bind("<KeyRelease>", self._calc_rate)

        f2 = ctk.CTkFrame(q_row, fg_color="transparent")
        f2.pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkLabel(f2, text="SL Lỗi:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w")
        self.ent_def = ctk.CTkEntry(f2, height=34, font=FONT_BODY)
        self.ent_def.pack(fill="x")
        self.ent_def.bind("<KeyRelease>", self._calc_rate)

        f3 = ctk.CTkFrame(q_row, fg_color="transparent")
        f3.pack(side="left", fill="x", expand=True, padx=(4, 0))
        ctk.CTkLabel(f3, text="Tỷ Lệ Lỗi:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w")
        self.lbl_rate = ctk.CTkLabel(f3, text="0.00%", height=34, font=FONT_H2,
                                     fg_color=BG_SURFACE, corner_radius=6, text_color=ACCENT_AMBER)
        self.lbl_rate.pack(fill="x")

        ctk.CTkLabel(c1, text="Người Chịu Trách Nhiệm (Manager/Leader):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_resp = ctk.CTkEntry(c1, height=34, font=FONT_BODY)
        self.ent_resp.pack(fill="x", **pad)

        ctk.CTkLabel(c1, text="Người Phụ Trách (Inspector/PIC):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_pic = ctk.CTkEntry(c1, height=34, font=FONT_BODY)
        self.ent_pic.pack(fill="x", **pad)

        ctk.CTkLabel(c1, text="Tiến Độ Khắc Phục:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.opt_prog = ctk.CTkOptionMenu(c1, values=["Chưa thực hiện", "Đang thực hiện", "Đã hoàn thành"],
                                          height=34, font=FONT_BODY)
        self.opt_prog.pack(fill="x", padx=14, pady=(3, 12))

        # --- SECTION 2: HÌNH ẢNH MINH HỌA & ĐÁNH DẤU LỖI ---
        c2 = ctk.CTkFrame(self.body, fg_color=BG_DEEP, corner_radius=8, border_width=1, border_color=BORDER_CLR)
        c2.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(c2, text="🖼️ Hình Ảnh Minh Họa & Đánh Dấu", font=FONT_H2, text_color=ACCENT_TEAL).pack(anchor="w", padx=14, pady=(10, 4))

        # Image preview box
        self.box_img = ctk.CTkFrame(c2, fg_color=BG_SURFACE, corner_radius=8, height=190)
        self.box_img.pack(fill="x", padx=14, pady=(4, 6))
        self.box_img.pack_propagate(False)

        self.lbl_img_preview = ctk.CTkLabel(self.box_img, text="[Chưa có ảnh]\nBấm 'Chọn File' hoặc chụp màn hình rồi dán (Ctrl+V)",
                                            font=FONT_SMALL, text_color=TEXT_MUTED)
        self.lbl_img_preview.pack(fill="both", expand=True, padx=8, pady=8)

        # Image control buttons
        btn_img_bar = ctk.CTkFrame(c2, fg_color="transparent")
        btn_img_bar.pack(fill="x", padx=14, pady=(2, 10))

        self.btn_studio = ctk.CTkButton(btn_img_bar, text="🔍 Soi & Đánh Dấu (Studio)", height=32, corner_radius=6,
                                        font=FONT_SMALL, fg_color=ACCENT_BLUE, text_color="#FFFFFF",
                                        command=self._open_studio, state="disabled")
        self.btn_studio.pack(side="left", fill="x", expand=True, padx=(0, 4))

        ctk.CTkButton(btn_img_bar, text="📁 Chọn", width=68, height=32, corner_radius=6,
                      font=FONT_SMALL, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._pick_image).pack(side="left", padx=2)

        ctk.CTkButton(btn_img_bar, text="📋 Dán", width=65, height=32, corner_radius=6,
                      font=FONT_SMALL, fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self._paste_image).pack(side="left", padx=2)

        self.btn_clear_img = ctk.CTkButton(btn_img_bar, text="✕", width=36, height=32, corner_radius=6,
                                           font=FONT_SMALL, fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
                                           command=self._clear_image)
        self.btn_clear_img.pack(side="left", padx=(2, 0))

        # --- SECTION 3: MÔ TẢ & ĐỐI SÁCH KHẮC PHỤC ---
        c3 = ctk.CTkFrame(self.body, fg_color=BG_DEEP, corner_radius=8, border_width=1, border_color=BORDER_CLR)
        c3.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(c3, text="📝 Mô Tả Hiện Tượng & Đối Sách", font=FONT_H2, text_color=ACCENT_TEAL).pack(anchor="w", padx=14, pady=(10, 4))

        ctk.CTkLabel(c3, text="Mô Tả Lỗi (Hiện tượng bất thường):", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.txt_desc = ctk.CTkTextbox(c3, height=65, font=FONT_BODY)
        self.txt_desc.pack(fill="x", **pad)

        ctk.CTkLabel(c3, text="Nguyên Nhân:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.txt_cause = ctk.CTkTextbox(c3, height=65, font=FONT_BODY)
        self.txt_cause.pack(fill="x", **pad)

        ctk.CTkLabel(c3, text="Biện Pháp Cải Tiến:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.txt_counter = ctk.CTkTextbox(c3, height=65, font=FONT_BODY)
        self.txt_counter.pack(fill="x", **pad)

        ctk.CTkLabel(c3, text="Tiêu Chuẩn Hóa SOP:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_sop = ctk.CTkEntry(c3, height=34, font=FONT_BODY)
        self.ent_sop.pack(fill="x", **pad)

        ctk.CTkLabel(c3, text="Ghi Chú:", font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", **pad)
        self.ent_notes = ctk.CTkEntry(c3, height=34, font=FONT_BODY)
        self.ent_notes.pack(fill="x", padx=14, pady=(3, 12))

        # ── DRAWER FOOTER ──
        footer = ctk.CTkFrame(self, fg_color=BG_SURFACE, height=54, corner_radius=10)
        footer.pack(fill="x", side="bottom", padx=10, pady=(4, 10))
        footer.pack_propagate(False)

        self.lbl_saving = ctk.CTkLabel(footer, text="", font=FONT_SMALL, text_color=ACCENT_TEAL)
        self.lbl_saving.pack(side="left", padx=12)

        ctk.CTkButton(footer, text="Đóng", width=80, height=36, corner_radius=8,
                      fg_color=BG_CARD, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
                      command=self.close_drawer).pack(side="right", padx=(4, 10))

        self.btn_save = ctk.CTkButton(footer, text="💾 Lưu Báo Cáo", width=140, height=36, corner_radius=8,
                                      fg_color=ACCENT_TEAL, text_color=("#FFFFFF", "#0B0F1A"),
                                      font=FONT_H2, command=self._save_report)
        self.btn_save.pack(side="right", padx=4)

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
            self.lbl_img_preview.configure(text=f"[Ảnh đã chọn: {len(b)} bytes]", image=None)

    def _load_remote_thumb(self, url: str):
        self.lbl_img_preview.configure(text="⏳ Đang tải ảnh...", image=None)
        self.btn_studio.configure(state="disabled")

        def _fetch():
            try:
                import urllib.request
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                ssl_ctx = svc.get_ssl_context()
                with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as r:
                    raw = r.read()
                im = Image.open(io.BytesIO(raw))
                self.current_pil_image = im
                safe_after(self, 0, lambda: self._show_thumb_from_pil(im))
            except Exception as e:
                safe_after(self, 0, lambda: self.lbl_img_preview.configure(text=f"[Không tải được ảnh: {e}]", image=None))

        threading.Thread(target=_fetch, daemon=True).start()

    def _clear_image(self):
        self.selected_image_bytes = None
        self.selected_image_name = None
        self.current_pil_image = None
        self.report_data["image_url"] = None
        self._ctk_preview = None
        self.lbl_img_preview.configure(image=None, text="[Chưa có ảnh]\nBấm 'Chọn File' hoặc chụp màn hình rồi dán (Ctrl+V)")
        self.btn_studio.configure(state="disabled")

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
                p = clip_img[0]
                if os.path.isfile(p):
                    with open(p, "rb") as fp:
                        self.selected_image_bytes = fp.read()
                    self.selected_image_name = os.path.basename(p)
                    self._show_thumb_from_bytes(self.selected_image_bytes)
            else:
                messagebox.showinfo("Clipboard", "Không tìm thấy ảnh trong Clipboard (Hãy chụp ảnh màn hình bằng Snipping Tool rồi thử lại).")
        except Exception as e:
            messagebox.showwarning("Lỗi Clipboard", f"Không thể dán ảnh:\n{e}")

    def _open_studio(self):
        if not self.current_pil_image and not self.report_data.get("image_url"):
            messagebox.showwarning("Chưa có ảnh", "Vui lòng chọn hoặc dán ảnh trước khi mở Studio!")
            return

        def _on_updated(updated_rep):
            if isinstance(updated_rep, dict) and updated_rep.get("image_url"):
                self.report_data["image_url"] = updated_rep.get("image_url")
                self._load_remote_thumb(updated_rep.get("image_url"))
            if self.main_view:
                self.main_view.refresh_data()

        AnomalyImageStudioWindow(self.main_view, self.report_data, initial_pil=self.current_pil_image, on_image_updated=_on_updated)

    def open_for_create(self):
        self.is_edit = False
        self.report_data = {}
        self.lbl_title.configure(text="➕ Thêm Báo Cáo Mới")
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
        self.opt_prog.set("Đang thực hiện")
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
        self.lbl_title.configure(text=f"✏️ Sửa Báo Cáo #{rep_id}")
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
        if report_data.get("progress"):
            self.opt_prog.set(report_data.get("progress"))
        else:
            self.opt_prog.set("Đang thực hiện")
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
        prod = self.ent_product.get().strip()
        if not prod:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập Tên Sản Phẩm (Model/PWB)!", parent=self)
            self.ent_product.focus()
            return

        try:
            tot_q = int(self.ent_tot.get().strip() or 0)
            def_q = int(self.ent_def.get().strip() or 0)
        except ValueError:
            messagebox.showwarning("Lỗi số lượng", "Số lượng kiểm tra và số lượng lỗi phải là số nguyên!", parent=self)
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
            "progress": self.opt_prog.get().strip(),
            "description": self.txt_desc.get("1.0", "end-1c").strip(),
            "root_cause": self.txt_cause.get("1.0", "end-1c").strip(),
            "countermeasures": self.txt_counter.get("1.0", "end-1c").strip(),
            "sop_standard": self.ent_sop.get().strip(),
            "notes": self.ent_notes.get().strip(),
        }
        if "image_url" in self.report_data and self.report_data["image_url"] is not None:
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
                import urllib.request
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                ssl_ctx = svc.get_ssl_context()
                with urllib.request.urlopen(req, timeout=12, context=ssl_ctx) as r:
                    raw = r.read()
                orig_im = Image.open(io.BytesIO(raw))
                thumb_im = self._create_thumbnail_image(orig_im, (46, 46))
                try:
                    thumb_im.save(disk_path, "PNG")
                except Exception:
                    pass

                def _apply():
                    photo = ImageTk.PhotoImage(thumb_im)
                    self._thumb_cache[url] = photo
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
        # ── TOP ACTION & FILTER BAR ──────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                               border_width=1, border_color=BORDER_CLR, height=64)
        top_bar.pack(fill="x", pady=(0, 10))
        top_bar.pack_propagate(False)

        # Search box
        self.ent_search = ctk.CTkEntry(top_bar, placeholder_text="🔍 Tìm kiếm sản phẩm, công đoạn, lỗi, người phụ trách...",
                                       width=280, height=36, corner_radius=8, font=FONT_BODY)
        self.ent_search.pack(side="left", padx=(14, 6), pady=14)
        self.ent_search.bind("<KeyRelease>", lambda e: self._apply_filter())

        # Process Filter
        self.opt_filter_proc = ctk.CTkOptionMenu(
            top_bar,
            values=["Tất cả công đoạn", "CHA", "SCP", "RAD", "CHP", "SMT", "Lắp Ráp", "KCS / Final QC"],
            height=36, corner_radius=8, font=FONT_SMALL, width=130,
            command=lambda v: self._apply_filter()
        )
        self.opt_filter_proc.pack(side="left", padx=4)

        # Progress Filter
        self.opt_filter_prog = ctk.CTkOptionMenu(
            top_bar,
            values=["Tất cả tiến độ", "Chưa thực hiện", "Đang thực hiện", "Đã hoàn thành"],
            height=36, corner_radius=8, font=FONT_SMALL, width=130,
            command=lambda v: self._apply_filter()
        )
        self.opt_filter_prog.pack(side="left", padx=4)

        # Right-side action buttons
        self.btn_admin_lock = ctk.CTkButton(
            top_bar, text="🔒 Admin: Đã Khóa", width=125, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._toggle_admin_lock
        )
        self.btn_admin_lock.pack(side="right", padx=(4, 14))

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

        # Action: Delete selected
        self.btn_delete = ctk.CTkButton(
            top_bar, text="🗑️ Xóa", width=80, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=ACCENT_RED, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._on_delete_report, state="disabled"
        )
        self.btn_delete.pack(side="right", padx=4)

        # Action: Edit selected
        self.btn_edit = ctk.CTkButton(
            top_bar, text="✏️ Sửa", width=80, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=TEXT_PRIMARY, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._on_edit_report, state="disabled"
        )
        self.btn_edit.pack(side="right", padx=4)

        # Action: Open Image Marking Studio
        self.btn_open_marking = ctk.CTkButton(
            top_bar, text="🖼️ Soi & Vẽ Ảnh", width=120, height=36, corner_radius=8,
            fg_color=BG_SURFACE, text_color=ACCENT_BLUE, hover_color=BG_HOVER,
            font=FONT_SMALL, command=self._open_marking_studio, state="disabled"
        )
        self.btn_open_marking.pack(side="right", padx=4)

        # Action: Add Report
        self.btn_add = ctk.CTkButton(
            top_bar, text="➕ Thêm Báo Cáo", width=135, height=36, corner_radius=8,
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

        self.lbl_tbl_count = ctk.CTkLabel(tbl_info, text="Danh Sách Báo Cáo: 0 bản ghi",
                                          font=FONT_H2, text_color=TEXT_PRIMARY)
        self.lbl_tbl_count.pack(side="left")

        lbl_hint = ctk.CTkLabel(tbl_info,
                                text="💡 Click vào 'Ảnh Lỗi' trên bảng để Soi & Đánh Dấu (Studio). Click đúp vào dòng để Sửa trong Drawer.",
                                font=FONT_SMALL, text_color=TEXT_MUTED)
        lbl_hint.pack(side="right")

        # Treeview with 16 data columns + #0 Thumbnail Image column
        tree_container = tk.Frame(self.tbl_card, bg="#131929")
        tree_container.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        cols = (
            "stt", "date", "process", "product", "machine", "tot", "def", "rate",
            "resp", "pic", "desc", "cause", "counter", "sop", "prog", "notes"
        )
        self.tree = ttk.Treeview(tree_container, columns=cols, show="tree headings",
                                 style="Anomaly.Treeview", selectmode="browse")

        self.tree.heading("#0", text="Ảnh Lỗi", anchor="center")
        self.tree.column("#0", width=75, minwidth=65, anchor="center", stretch=False)

        headers_meta = [
            ("stt", "STT", 50, "center"),
            ("date", "Ngày Tháng", 105, "center"),
            ("process", "Công Đoạn", 95, "center"),
            ("product", "Sản Phẩm (Model/PWB)", 170, "center"),
            ("machine", "Máy Móc / Line", 115, "center"),
            ("tot", "SL Kiểm", 90, "center"),
            ("def", "SL Lỗi", 85, "center"),
            ("rate", "Tỷ Lệ (%)", 90, "center"),
            ("resp", "Người Chịu TN", 140, "center"),
            ("pic", "Người Phụ Trách", 130, "center"),
            ("desc", "Mô Tả Hiện Tượng Lỗi", 280, "center"),
            ("cause", "Nguyên Nhân", 250, "center"),
            ("counter", "Biện Pháp Cải Tiến", 280, "center"),
            ("sop", "Tiêu Chuẩn SOP", 125, "center"),
            ("prog", "Tiến Độ", 125, "center"),
            ("notes", "Ghi Chú", 180, "center")
        ]

        for col_id, col_name, col_w, col_align in headers_meta:
            self.tree.heading(col_id, text=col_name, anchor="center")
            self.tree.column(col_id, width=col_w, anchor="center", minwidth=45)

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
        self.ctx_menu.add_command(label="🖼️ Soi & Đánh Dấu Ảnh Lỗi (Studio)", command=self._open_marking_studio)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="✏️ Chỉnh Sửa Báo Cáo", command=self._on_edit_report)
        self.ctx_menu.add_command(label="🗑️ Xóa Báo Cáo Này", command=self._on_delete_report)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="📋 Sao Chép Thông Tin Dòng", command=self._copy_row_info)

    # ── DATA FETCH & POPULATE ────────────────────────────────────────────────
    def refresh_data(self):
        """Fetches reports from Supabase asynchronously."""
        self.btn_refresh.configure(state="disabled", text="⏳ Đang tải...")

        def _fetch():
            try:
                reps = svc.fetch_all_reports()
                safe_after(self, 0, lambda: self._on_fetch_success(reps))
            except Exception as e:
                safe_after(self, 0, lambda: self._on_fetch_error(str(e)))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_fetch_success(self, reports: list[dict]):
        self.all_reports = reports
        self.selected_report = None
        self._update_action_buttons_state()
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
            if f_proc != "Tất cả công đoạn" and str(r.get("process")).lower() != f_proc.lower():
                continue
            if f_prog != "Tất cả tiến độ" and str(r.get("progress")).lower() != f_prog.lower():
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
        self.lbl_tbl_count.configure(text=f"Danh Sách Báo Cáo: {len(reports)} bản ghi")

        # Clear selection state if empty
        if not reports:
            self.selected_report = None
            self._update_action_buttons_state()

        for idx, r in enumerate(reports, 1):
            rate_val = r.get("defect_rate")
            rate_str = f"{rate_val:.2f}%" if rate_val is not None else "0.00%"

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
                f"{r.get('total_qty', 0):,}",
                f"{r.get('defect_qty', 0):,}",
                rate_str,
                r.get("responsible_person", ""),
                r.get("pic", ""),
                r.get("description", ""),
                r.get("root_cause", ""),
                r.get("countermeasures", ""),
                r.get("sop_standard", ""),
                r.get("progress", "Đang thực hiện"),
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
        """Clicking directly on the 'Ảnh Lỗi' column cell (#0) opens the Image Studio, or prompts to add image in Drawer."""
        col_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        region = self.tree.identify_region(event.x, event.y)
        if row_id and (col_id == "#0" or region == "tree"):
            self.tree.selection_set(row_id)
            self._on_row_select(None)
            found = [r for r in self.all_reports if str(r.get("id")) == str(row_id)]
            if found:
                if found[0].get("image_url"):
                    self.after(50, self._open_marking_studio)
                else:
                    if messagebox.askyesno("Chưa có ảnh", "Báo cáo này chưa có hình ảnh đính kèm.\nBạn có muốn mở Drawer để thêm ảnh không?", parent=self):
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
        if hasattr(self.app, "show_toast"):
            self.app.show_toast("📋 Đã sao chép thông tin báo cáo vào Clipboard!")

    # ── OPEN IMAGE MARKING STUDIO ────────────────────────────────────────────
    def _open_marking_studio(self):
        if not self.selected_report or not self.selected_report.get("image_url"):
            messagebox.showinfo("Không có ảnh", "Báo cáo này không có hình ảnh đính kèm để xem hoặc đánh dấu.")
            return

        def _on_img_updated(new_rep):
            self.refresh_data()

        AnomalyImageStudioWindow(self, self.selected_report, on_image_updated=_on_img_updated)

    # ── ADMIN & PASSWORD VERIFICATION ────────────────────────────────────────
    def _require_admin(self, callback):
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
            m = messagebox.askyesnocancel(
                "Quản Trị",
                "Bạn đang ở chế độ Quản Trị.\n\n• Chọn YES để Khóa lại\n• Chọn NO để Đổi Mật Khẩu\n• Chọn CANCEL để đóng"
            )
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
        self.drawer.open_for_create()

    def _on_edit_report(self):
        if not self.selected_report:
            messagebox.showwarning("Chưa chọn", "Vui lòng click chọn một báo cáo trên bảng để chỉnh sửa!")
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
        if hasattr(self.app, "show_toast"):
            self.app.show_toast("✅ Đã lưu báo cáo bất thường thành công!")

    def _on_delete_report(self):
        if not self.selected_report:
            messagebox.showwarning("Chưa chọn", "Vui lòng click chọn một báo cáo trên bảng để xóa!")
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
                    if hasattr(self, "drawer") and self.drawer.winfo_manager() == "grid":
                        if str(self.drawer.report_data.get("id")) == str(rep.get("id")):
                            self.drawer.close_drawer()
                    self.selected_report = None
                    self._update_action_buttons_state()
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
                        rowheight=56,
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

        self._make_placeholder_images(is_dark=is_dark)

        self.tree.tag_configure("odd", background=odd_bg)
        self.tree.tag_configure("even", background=even_bg)
        self.tree.tag_configure("tag_done", background="#064E3B" if is_dark else "#DCFCE7",
                                foreground="#A7F3D0" if is_dark else "#166534")
        self.tree.tag_configure("tag_open", background="#7F1D1D" if is_dark else "#FEE2E2",
                                foreground="#FECACA" if is_dark else "#991B1B")
