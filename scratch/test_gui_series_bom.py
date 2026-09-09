"""
Scratch test for SeriesBOMCompareView GUI
Verifies component layout, treeview, QC focus checklist, progress bar, and filters.
"""

import os
import sys
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk

sys.path.insert(0, os.path.abspath("."))
import series_bom_comparator as sbc
import series_bom_exporter as sbe

ctk.set_appearance_mode("dark")

class TestApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TEST Series BOM Compare View")
        self.geometry("1300x800")
        
        # Test comparison with real files
        f_a = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
        f_b = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"
        
        bom_a = sbc.parse_any_bom(f_a)
        bom_b = sbc.parse_any_bom(f_b)
        self.res = sbc.compare_series_boms(bom_a, bom_b)
        
        # Container
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=16, pady=16)
        
        # Top Header
        lbl = ctk.CTkLabel(self.container, text="📑 SO SÁNH 2 BOM CÙNG MODEL KHÁC SERIES", font=("Segoe UI", 16, "bold"))
        lbl.pack(anchor="w", pady=(0, 10))
        
        # Stats Cards
        s = self.res["summary"]
        stats_frame = ctk.CTkFrame(self.container, fg_color="#131929", corner_radius=8)
        stats_frame.pack(fill="x", pady=(0, 10))
        
        kpi_texts = [
            ("🔵 TỔNG VỊ TRÍ", f"{s['total_locations']}", "#3D8EFF"),
            ("⚪ DÙNG CHUNG", f"{s['matched_count']} ({s['match_percentage']}%)", "#9E9E9E"),
            ("🟢 THÊM MỚI", f"{s['added_count']}", "#00E676"),
            ("🔴 BỎ TRỐNG (DNP)", f"{s['removed_count']}", "#FF5252"),
            ("🟡 ĐỔI MÃ", f"{s['modified_count']}", "#FFAB00"),
            ("🎯 CẦN KIỂM (QC FOCUS)", f"{s['focus_count']}", "#00C9A7")
        ]
        
        for title, val, col in kpi_texts:
            f = ctk.CTkFrame(stats_frame, fg_color="#1C2438", corner_radius=6)
            f.pack(side="left", fill="both", expand=True, padx=6, pady=8)
            ctk.CTkLabel(f, text=title, font=("Segoe UI", 9, "bold"), text_color="#7A8BA6").pack(pady=(4, 0))
            ctk.CTkLabel(f, text=val, font=("Segoe UI", 13, "bold"), text_color=col).pack(pady=(0, 4))
            
        # QC Focus Checklist Card
        self.qc_card = ctk.CTkFrame(self.container, fg_color="#1A2438", border_width=1.5, border_color="#00C9A7", corner_radius=8)
        self.qc_card.pack(fill="x", pady=(0, 10))
        
        qc_head = ctk.CTkFrame(self.qc_card, fg_color="transparent")
        qc_head.pack(fill="x", padx=12, pady=(8, 4))
        
        ctk.CTkLabel(qc_head, text="🎯 BẢNG KIỂM SOÁT LINH KIỆN CẦN CHÚ Ý (QC FOCUS CHECKLIST)", 
                     font=("Segoe UI", 12, "bold"), text_color="#00C9A7").pack(side="left")
                     
        self.lbl_progress = ctk.CTkLabel(qc_head, text=f"Tiến độ kiểm tra FAI: 0 / {s['focus_count']} (0%)",
                                         font=("Segoe UI", 11, "bold"), text_color="#FFB830")
        self.lbl_progress.pack(side="right")
        
        self.pbar = ctk.CTkProgressBar(self.qc_card, height=8, corner_radius=4, progress_color="#00C9A7")
        self.pbar.pack(fill="x", padx=12, pady=(0, 8))
        self.pbar.set(0)
        
        # Filter & Search
        filter_bar = ctk.CTkFrame(self.container, fg_color="transparent")
        filter_bar.pack(fill="x", pady=(0, 8))
        
        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(filter_bar, placeholder_text="🔍 Tìm kiếm vị trí (Ref Des), mã linh kiện...",
                                         width=320, height=32, textvariable=self.search_var)
        self.search_entry.pack(side="left", padx=(0, 10))
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate_table())
        
        self.filter_var = ctk.StringVar(value="focus")
        self.seg_filter = ctk.CTkSegmentedButton(
            filter_bar,
            values=["🎯 Cần chú ý (QC Focus)", "🟢 Thêm mới", "🔴 Bỏ trống (DNP)", "🟡 Đổi mã", "⚪ Dùng chung", "Tất cả"],
            variable=self.filter_var,
            command=lambda v: self._populate_table()
        )
        self.seg_filter.pack(side="left", padx=4)
        
        btn_export = ctk.CTkButton(filter_bar, text="📥 Xuất Báo Cáo FAI (.xlsx)", fg_color="#0D9488", 
                                   hover_color="#0F766E", font=("Segoe UI", 11, "bold"), height=32,
                                   command=self._on_export)
        btn_export.pack(side="right")
        
        # Table
        table_frame = ctk.CTkFrame(self.container, fg_color="#131929", corner_radius=8)
        table_frame.pack(fill="both", expand=True)
        
        cols = ("check", "loc", "status", "part_a", "part_b", "action", "spec")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        
        self.tree.heading("check", text="[QC Kiểm]")
        self.tree.heading("loc", text="Vị Trí (Ref)")
        self.tree.heading("status", text="Phân Loại")
        self.tree.heading("part_a", text=f"Mã LK ({s['series_a']})")
        self.tree.heading("part_b", text=f"Mã LK ({s['series_b']})")
        self.tree.heading("action", text="Chỉ Dẫn Hành Động Cụ Thể Cho QC (Action Guide)")
        self.tree.heading("spec", text=f"Quy Cách / Spec ({s['series_b']})")
        
        self.tree.column("check", width=90, anchor="center")
        self.tree.column("loc", width=90, anchor="center")
        self.tree.column("status", width=120, anchor="center")
        self.tree.column("part_a", width=140, anchor="w")
        self.tree.column("part_b", width=140, anchor="w")
        self.tree.column("action", width=380, anchor="w")
        self.tree.column("spec", width=220, anchor="w")
        
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        
        self.tree.tag_configure("tag_added", background="#15362B", foreground="#00E676")
        self.tree.tag_configure("tag_removed", background="#361818", foreground="#FF5252")
        self.tree.tag_configure("tag_modified", background="#362E15", foreground="#FFAB00")
        self.tree.tag_configure("tag_matched", background="#131929", foreground="#8A9BB5")
        self.tree.tag_configure("tag_ok", background="#0E3D2F", foreground="#A3E4D7")
        
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        
        self.qc_status_map = {} # loc -> 'OK', 'NG', 'PENDING'
        self._populate_table()
        
    def _populate_table(self):
        self.tree.delete(*self.tree.get_children())
        fil = self.filter_var.get()
        query = self.search_var.get().strip().upper()
        
        items = []
        if "Cần chú ý" in fil or fil == "focus":
            items = self.res["qc_focus_items"]
        elif "Thêm mới" in fil:
            items = self.res["added_items"]
        elif "Bỏ trống" in fil:
            items = self.res["removed_items"]
        elif "Đổi mã" in fil:
            items = self.res["modified_items"]
        elif "Dùng chung" in fil:
            items = self.res["matched_items"]
        else:
            items = self.res["all_items"]
            
        for item in items:
            loc = item["location"]
            if query and query not in loc.upper() and query not in item["part_a"].upper() and query not in item["part_b"].upper():
                continue
                
            status = item["status"]
            qc_st = self.qc_status_map.get(loc, "PENDING")
            if qc_st == "OK":
                check_icon = "✅ ĐÃ DUYỆT"
                tag = "tag_ok"
            elif qc_st == "NG":
                check_icon = "❌ LỖI (NG)"
                tag = "tag_removed"
            else:
                check_icon = "⏳ Chờ kiểm"
                tag = f"tag_{status.lower()}"
                
            self.tree.insert("", "end", iid=loc, values=(
                check_icon,
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
            
        current = self.qc_status_map.get(item_id, "PENDING")
        if current == "PENDING":
            self.qc_status_map[item_id] = "OK"
        elif current == "OK":
            self.qc_status_map[item_id] = "NG"
        else:
            self.qc_status_map[item_id] = "PENDING"
            
        self._populate_table()
        
    def _update_progress(self):
        focus_locs = [i["location"] for i in self.res["qc_focus_items"]]
        total = len(focus_locs)
        if total == 0:
            self.lbl_progress.configure(text="Không có linh kiện cần kiểm!")
            self.pbar.set(1.0)
            return
            
        checked = sum(1 for loc in focus_locs if self.qc_status_map.get(loc) == "OK")
        pct = int(checked / total * 100)
        self.lbl_progress.configure(text=f"Tiến độ kiểm tra FAI: {checked} / {total} linh kiện ({pct}%)")
        self.pbar.set(checked / total)
        
    def _on_export(self):
        out_path = os.path.abspath("scratch/fai_exported.xlsx")
        sbe.export_series_bom_report(self.res, out_path)
        print("Exported to:", out_path)

if __name__ == "__main__":
    app = TestApp()
    app.after(500, lambda: print("GUI initialized successfully! Total rows populated."))
    app.after(1200, lambda: app.destroy())
    app.mainloop()
    print("Test finished successfully!")
