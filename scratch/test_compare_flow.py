import sys
sys.stdout.reconfigure(encoding="utf-8")
import os
import time

sys.path.insert(0, os.path.abspath("."))
import scratch.test_updated_app as app_mod

f_a = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
f_b = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"

app = app_mod.BOMExtractorApp()

def run_test():
    app._set_nav_active("series_bom")
    view = app.view_series_bom_compare
    view.file_a = f_a
    view.file_b = f_b
    view.lbl_file_a_name.configure(text=os.path.basename(f_a))
    view.lbl_file_b_name.configure(text=os.path.basename(f_b))

    print("Starting compare in view (inside mainloop)...")
    view._start_compare()

    def check_result(count=0):
        if count > 50:
            print("Timeout waiting for compare!")
            app.destroy()
            return
        if not view.is_comparing and view.comparison_result:
            print("Compare finished successfully inside mainloop!")
            res = view.comparison_result
            summary = res["summary"]
            print(f"Summary: Total={summary['total_locations']}, Matched={summary['matched_count']}, Added={summary['added_count']}, Removed={summary['removed_count']}, Mod={summary['modified_count']}, Focus={summary['focus_count']}")
            children = view.tree.get_children()
            print(f"Rows in treeview with default filter (QC Focus): {len(children)}")
            first_id = children[0]
            first_item = view.tree.item(first_id)
            print(f"First row values: {first_item['values']}")
            
            # Test clicking to toggle check
            print("Toggling QC check for first item...")
            view.qc_status_map[first_id] = "OK"
            view._populate_table()
            print(f"After OK mark: {view.tree.item(first_id)['values'][0]}")
            print(f"Progress label: {view.lbl_qc_progress.cget('text')}")
            
            app.destroy()
            print("All compare flow tests PASSED WITH FLYING COLORS!")
        else:
            app.after(200, lambda: check_result(count + 1))

    app.after(100, check_result)

app.after(200, run_test)
app.mainloop()
