import sys
import os

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath("."))
import series_bom_comparator as sbc

f_a = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994554.pdf"
f_b = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788851994499.pdf"

bom_a = sbc.parse_any_bom(f_a)
bom_b = sbc.parse_any_bom(f_b)

res = sbc.compare_series_boms(bom_a, bom_b)
s = res["summary"]

print("=== COMPARISON RESULTS ===")
print(f"Model A: {s['model_a']} | Series: {s['series_a']} | File: {s['file_a']}")
print(f"Model B: {s['model_b']} | Series: {s['series_b']} | File: {s['file_b']}")
print(f"Validation: {s['validation_status']} -> {s['validation_msg']}")
print(f"Total Locations: {s['total_locations']}")
print(f"Matched (Common): {s['matched_count']} ({s['match_percentage']}%)")
print(f"Added in Series B: {s['added_count']}")
print(f"Removed in Series B: {s['removed_count']}")
print(f"Modified: {s['modified_count']}")
print(f"QC Focus Items (Total to inspect): {s['focus_count']}")

print("\n--- ALL QC FOCUS ITEMS WITH ACTIONS ---")
for idx, item in enumerate(res["qc_focus_items"], 1):
    print(f"{idx:2d}. [{item['location']:<6}] {item['status_vn']:<14} | {item['action_guide']}")
