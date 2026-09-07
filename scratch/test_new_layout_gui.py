import os, sys
sys.stdout.reconfigure(encoding='utf-8')

# Ensure DLL search path for Windows Python 3.8+
for p in [r"C:\Program Files\Common Files\microsoft shared\ClickToRun", r"C:\Program Files\Microsoft Office\root\Client"]:
    if os.path.exists(p):
        try: os.add_dll_directory(p)
        except Exception: pass

ws_path = r"e:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder"
sys.path.insert(0, ws_path)
os.chdir(ws_path)

from app import BOMExtractorApp

print("Initializing BOMExtractorApp...")
app = BOMExtractorApp()
app.withdraw()  # Keep hidden during automated test

print("Testing Navigation to Tab 3: Model Compare...")
app._set_nav_active("model_comp")
app.update()

mc = app.view_model_compare
assert mc is not None, "ModelCompareView not initialized!"
print("ModelCompareView initialized successfully.")

# Verify Layout: cards_row, card_verdict, workspace_card are packed
assert mc.cards_row.winfo_manager() == "pack", "cards_row should be packed at bottom!"
assert mc.card_verdict.winfo_manager() == "pack", "card_verdict should be packed at bottom!"
assert mc.workspace_card.winfo_manager() == "pack", "workspace_card should be packed at top!"
print("✓ Cards row, verdict banner, and workspace card all properly packed.")

# Test Display Mode Switcher
print("Testing Display Mode Switcher...")
mc._on_display_mode_change("📋 Bảng Dữ Liệu")
app.update()
assert mc.table_card.winfo_manager() == "pack", "Table card should be packed in 'table' mode"
assert mc.inspector_card.winfo_manager() == "", "Inspector card should not be packed in 'table' mode"
print("✓ 'table' mode works.")

mc._on_display_mode_change("📐 Bản Vẽ PCB")
app.update()
assert mc.table_card.winfo_manager() == "", "Table card should not be packed in 'drawing' mode"
assert mc.inspector_card.winfo_manager() == "pack", "Inspector card should be packed in 'drawing' mode"
print("✓ 'drawing' mode works.")

mc._on_display_mode_change("◫ Song Song")
app.update()
assert mc.table_card.winfo_manager() == "pack", "Table card should be packed in 'split' mode"
assert mc.inspector_card.winfo_manager() == "pack", "Inspector card should be packed in 'split' mode"
print("✓ 'split' mode works.")

# Test Loading sample PDFs and running comparison
pdf_a = os.path.join(ws_path, "sample_manuals", "media_1788593034521.pdf")
pdf_b = os.path.join(ws_path, "sample_manuals", "media_1788593034536.pdf")

if os.path.exists(pdf_a) and os.path.exists(pdf_b):
    print("Testing Comparison Run with Sample PDFs...")
    mc.pdf_a_path = pdf_a
    mc.pdf_b_path = pdf_b
    mc._run_comparison()
    app.update_idletasks()
    
    assert mc.comp_result is not None, "comp_result should be populated!"
    diff_count = mc.comp_result["summary"]["diff_count"]
    print(f"✓ Comparison successful: {diff_count} diffs found.")
    
    # Test Opening Fullscreen Window
    print("Testing Fullscreen Drawing Modal...")
    mc._open_fullscreen_drawing()
    app.update_idletasks()
    
    fw = mc.fullscreen_window
    assert fw is not None, "Fullscreen window should be instantiated!"
    assert fw.winfo_exists(), "Fullscreen window should exist!"
    
    # Test Stepper in Fullscreen
    print("Testing diff stepper in Fullscreen...")
    fw._next_diff()
    app.update_idletasks()
    fw._prev_diff()
    app.update_idletasks()
    print("✓ Diff stepper in Fullscreen works.")
    
    # Test Closing Fullscreen
    fw._close()
    app.update_idletasks()
    assert mc.fullscreen_window is None, "Fullscreen window reference should be reset after closing!"
    print("✓ Fullscreen modal closed cleanly.")

# Test Theme & Language Switch
print("Testing Theme switching...")
app._on_theme_change("☀️ Sáng")
app.update()
app._on_theme_change("🌙 Tối")
app.update()
print("✓ Theme switching works.")

print("Testing Language switching...")
app._on_language_change("中文")
app.update()
app._on_language_change("English")
app.update()
app._on_language_change("Tiếng Việt")
app.update()
print("✓ Language switching works.")

app.destroy()
print("\n🎉 ALL TESTS PASSED SUCCESSFULLY 100%!")
