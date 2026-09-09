import sys
sys.stdout.reconfigure(encoding="utf-8")

with open("app.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

print(f"Total lines in app.py: {len(lines)}")

patterns = ["series_bom", "btn_series_upload_drawing", "drawing_panel", "_build_series_bom_tab", "_render_series_drawing"]

for idx, line in enumerate(lines):
    for pat in patterns:
        if pat.lower() in line.lower():
            safe_line = line.strip()[:120]
            print(f"{idx+1}: {safe_line}")
            break
