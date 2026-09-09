import sys
sys.stdout.reconfigure(encoding="utf-8")
import ast

with open("app.py", "r", encoding="utf-8") as f:
    content = f.read()

lines = content.splitlines(keepends=True)

start_idx = None
end_idx = None
for idx, l in enumerate(lines):
    if "class SeriesBOMCompareView" in l:
        start_idx = idx
    if "class BOMExtractorApp" in l:
        end_idx = idx
        break

print(f"Found SeriesBOMCompareView from line {start_idx+1} to {end_idx+1}")
print(f"Target snippet first line: {lines[start_idx].strip()}")
print(f"Target snippet last line: {lines[end_idx-1].strip()}")
