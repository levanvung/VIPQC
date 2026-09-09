import sys
sys.stdout.reconfigure(encoding="utf-8")
import py_compile

# We read the new code cleanly from a separate file to avoid escape sequence issues
new_class_code = """class SeriesBOMCompareView(ctk.CTkFrame):
    \"\"\"
    Dedicated view for comparing 2 BOMs of the same model across different series.
    Supports ERP Multi-level BOM PDFs and Excel BOMs.
    Features the QC Focus Checklist, Interactive FAI Verification,
    and Holographic Cyberpunk AI Loading Animation.
    \"\"\"
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app

        self.file_a = None
        self.file_b = None

        self.bom_a_data = None
        self.bom_b_data = None
        self.comparison_result = None

        self.filter_mode = "focus"
        self.search_query = ""
        self.is_comparing = False

        self.qc_status_map = {}  # loc -> 'OK', 'NG', 'PENDING'
        self.active_loc = None

        self._loading_hud = None

        self._build_ui()

    def t(self, key: str, **kwargs) -> str:
        return self.app.t(key, **kwargs)

    def _build_ui(self):
        # ── 1. TOP CONTAINER: 2 UPLOAD CARDS + ACTIONS ───────────────────────
        top_container = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=12,
                                     border_width=1, border_color=BORDER_CLR)
        top_container.pack(fill="x", padx=4, pady=(4, 6))

        inner_top = ctk.CTkFrame(top_container, fg_color="transparent")
        inner_top.pack(fill="x", padx=10, pady=8)

        # Card 1: Series A BOM
        self.card_a = ctk.CTkFrame(inner_top, fg_color=BG_SURFACE, corner_radius=8,
                                   border_width=1, border_color=BORDER_CLR)
        self.card_a.pack(side="left", fill="both", expand=True, padx=(0, 6))

        a_head = ctk.CTkFrame(self.card_a, fg_color="transparent")
        a_head.pack(fill="x", padx=8, pady=(6, 2))
        ctk.CTkLabel(a_head, text="1. BOM GỐC (SERIES A)", font=("Segoe UI", 10, "bold"),
                     text_color=ACCENT_BLUE).pack(side="left")
        ctk.CTkButton(
            a_head, text="📂 Chọn BOM A", font=("Segoe UI", 9, "bold"),
            height=24, width=105, fg_color=ACCENT_BLUE, hover_color="#1E40AF",
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
        self.card_b.pack(side="left", fill="both", expand=True, padx=(6, 8))

        b_head = ctk.CTkFrame(self.card_b, fg_color="transparent")
        b_head.pack(fill="x", padx=8, pady=(6, 2))
        ctk.CTkLabel(b_head, text="2. BOM SO SÁNH (SERIES B)", font=("Segoe UI", 10, "bold"),
                     text_color=ACCENT_TEAL).pack(side="left")
        ctk.CTkButton(
            b_head, text="📂 Chọn BOM B", font=("Segoe UI", 9, "bold"),
            height=24, width=105, fg_color=ACCENT_TEAL, hover_color="#0D9488",
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

        # Action Buttons Box
        act_box = ctk.CTkFrame(inner_top, fg_color="transparent")
        act_box.pack(side="right", fill="y", padx=(4, 0))

        self.btn_run_compare = ctk.CTkButton(
            act_box, text="⚡ SO SÁNH 2 BOM", font=("Segoe UI", 11, "bold"),
            height=34, width=160, fg_color=ACCENT_TEAL, hover_color="#0F766E",
            command=self._start_compare
        )
        self.btn_run_compare.pack(fill="x", pady=(1, 3))

        self.btn_export_fai = ctk.CTkButton(
            act_box, text="📥 Xuất Báo Cáo FAI", font=("Segoe UI", 9, "bold"),
            height=24, width=160, fg_color="#1E3A8A", hover_color="#1E40AF",
            state="disabled", command=self._export_excel
        )
        self.btn_export_fai.pack(fill="x", pady=(1, 0))

        # ── 2. VALIDATION & STATS BANNER ─────────────────────────────────────
        self.banner_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.banner_frame.pack(fill="x", padx=4, pady=(0, 4))

        self.lbl_validation = ctk.CTkLabel(
            self.banner_frame, text="💡 Vui lòng chọn 2 file BOM (PDF hoặc Excel) rồi nhấn '⚡ SO SÁNH 2 BOM'.",
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

        # ── 3. QC FOCUS CHECKLIST HERO CARD ──────────────────────────────────
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

        self.lbl_qc_progress = ctk.CTkLabel(
            qc_head, text="Tiến độ kiểm tra FAI: 0 / 0 (0%)",
            font=("Segoe UI", 10, "bold"), text_color=ACCENT_AMBER
        )
        self.lbl_qc_progress.pack(side="right", padx=10)

        qc_sub = ctk.CTkFrame(self.card_qc_focus, fg_color="transparent")
        qc_sub.pack(fill="x", padx=10, pady=(0, 3))

        ctk.CTkLabel(
            qc_sub, text="⚡ Click vào từng dòng để đổi trạng thái kiểm tra (⏳ Chờ kiểm ➔ ✅ ĐÃ DUYỆT ➔ ❌ LỖI NG). Màu xanh: Thêm | Đỏ: Bớt (DNP) | Vàng: Đổi mã.",
            font=("Segoe UI", 8, "italic"), text_color=TEXT_MUTED
        ).pack(side="left")

        self.btn_mark_all_ok = ctk.CTkButton(
            qc_sub, text="✓ Đánh dấu tất cả OK", font=("Segoe UI", 8, "bold"),
            height=20, width=120, fg_color=ACCENT_TEAL, hover_color="#0D9488",
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

        # ── 4. TABLE VIEW ────────────────────────────────────────────────────
        self.table_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.table_panel.pack(fill="both", expand=True, padx=4, pady=(0, 2))

        # Filter Bar for table
        tbl_filter_bar = ctk.CTkFrame(self.table_panel, fg_color="transparent")
        tbl_filter_bar.pack(fill="x", pady=(0, 4))

        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            tbl_filter_bar, placeholder_text="🔍 Tìm vị trí (Ref Des), mã linh kiện, quy cách...",
            width=280, height=28, corner_radius=6, textvariable=self.search_var
        )
        self.search_entry.pack(side="left", padx=(0, 6))
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate_table())

        self.filter_var = ctk.StringVar(value="🎯 Cần chú ý")
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

        self.tree.column("check", width=110, anchor="center")
        self.tree.column("loc", width=95, anchor="center")
        self.tree.column("status", width=120, anchor="center")
        self.tree.column("part_a", width=160, anchor="w")
        self.tree.column("part_b", width=160, anchor="w")
        self.tree.column("action", width=380, anchor="w")
        self.tree.column("spec", width=220, anchor="w")

        vsb_tbl = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        hsb_tbl = ttk.Scrollbar(table_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb_tbl.set, xscrollcommand=hsb_tbl.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb_tbl.pack(side="right", fill="y")
        hsb_tbl.pack(side="bottom", fill="x")

        self._apply_tree_tags()
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

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

    # ── FILE SELECTION HANDLERS ──────────────────────────────────────────────
    def _get_current_bom_model(self) -> str:
        \"\"\"Returns the base model of the current BOMs (e.g. 'CHA3259AF').\"\"\"
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

    # ── HOLOGRAPHIC CYBERPUNK HUD LOADING CARD ───────────────────────────────
    def _show_loading_hud(self, title: str = "VIPQC AI — SO SÁNH DỮ LIỆU 2 BOM"):
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
            self.after(0, lambda: self._update_loading_hud(0.95, "[PHASE 3/3] Tổng hợp danh mục & hướng dẫn kiểm tra QC..."))

            time.sleep(0.3)
            self.after(0, lambda: self._on_compare_finished(res))
        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._on_compare_error(err_msg))

    def _on_compare_finished(self, result):
        self._hide_loading_hud()
        self.comparison_result = result
        self.btn_run_compare.configure(state="normal", text="⚡ SO SÁNH 2 BOM")
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

        self.app.set_status(f"Hoàn thành so sánh 2 BOM! {s['focus_count']} linh kiện cần chú ý kiểm tra.")

    def _on_compare_error(self, err_msg):
        self._hide_loading_hud()
        self.btn_run_compare.configure(state="normal", text="⚡ SO SÁNH 2 BOM")
        self.is_comparing = False
        self.lbl_validation.configure(text=f"❌ Lỗi khi so sánh: {err_msg}", text_color="#FF5252")
        messagebox.showerror("Lỗi So Sánh", f"Đã xảy ra lỗi khi phân tích BOM:\\n\\n{err_msg}")

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

        self.active_loc = item_id

        # Toggle QC check status: PENDING -> OK -> NG -> PENDING
        current = self.qc_status_map.get(item_id, "PENDING")
        if current == "PENDING":
            self.qc_status_map[item_id] = "OK"
        elif current == "OK":
            self.qc_status_map[item_id] = "NG"
        else:
            self.qc_status_map[item_id] = "PENDING"

        self._populate_table()
        if self.tree.exists(item_id):
            self.tree.selection_set(item_id)

    def _on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id or not self.comparison_result:
            return

        matched = [i for i in self.comparison_result["all_items"] if i["location"] == item_id]
        if not matched:
            return
        item = matched[0]

        msg = (
            f"VỊ TRÍ: {item['location']}\\n"
            f"Trạng thái: {item['status_vn']}\\n\\n"
            f"Mã linh kiện Series A: {item['part_a']}\\n"
            f"Mã linh kiện Series B: {item['part_b']}\\n\\n"
            f"HƯỚNG DẪN KIỂM TRA QC:\\n{item['action_guide']}\\n\\n"
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
                f"Đã tạo file báo cáo FAI thành công tại:\\n{out_path}\\n\\nBạn có muốn mở file ngay không?"
            )
            if resp:
                os.startfile(out_path)
        except Exception as e:
            messagebox.showerror("Lỗi Xuất File", f"Không thể xuất file Excel:\\n\\n{str(e)}")
"""

with open("app.py", "r", encoding="utf-8") as f:
    orig = f.read()

lines = orig.splitlines(keepends=True)
start_idx = None
end_idx = None
for idx, l in enumerate(lines):
    if "class SeriesBOMCompareView" in l:
        start_idx = idx
    if "class BOMExtractorApp" in l:
        end_idx = idx
        break

new_app_content = "".join(lines[:start_idx]) + new_class_code + "\n\n" + "".join(lines[end_idx:])

with open("scratch/test_updated_app.py", "w", encoding="utf-8") as f:
    f.write(new_app_content)

print(f"Wrote scratch/test_updated_app.py: {len(new_app_content)} chars")

try:
    py_compile.compile("scratch/test_updated_app.py", doraise=True)
    print("Syntax verification SUCCESSFUL! Code compiles with NO syntax errors.")
except Exception as e:
    print(f"Syntax error: {e}")
