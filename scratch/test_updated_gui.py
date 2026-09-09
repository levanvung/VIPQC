import sys
sys.stdout.reconfigure(encoding="utf-8")
import os
import customtkinter as ctk

sys.path.insert(0, os.path.abspath("."))
import scratch.test_updated_app as app_mod

print("Testing app initialization...")
app = app_mod.BOMExtractorApp()
app.update()

print("Current nav:", app.current_nav)
print("Switching to series_bom...")
app._set_nav_active("series_bom")
app.update()

view = app.view_series_bom_compare
print("view_series_bom_compare is visible:", view.winfo_viewable())
print("Table container exists:", hasattr(view, "table_panel"))
print("Treeview columns:", view.tree["columns"])
print("Checking that drawing_panel is NOT present:", not hasattr(view, "drawing_panel"))
print("Checking that card_dwg is NOT present:", not hasattr(view, "card_dwg"))

app.destroy()
print("GUI test completed successfully!")
