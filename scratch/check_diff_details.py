import sys, os
sys.path.insert(0, os.path.abspath("scratch"))
from compare_erp_boms import rec_1a, rec_1d

parts_1a = {r['part_no']: r for r in rec_1a}
parts_1d = {r['part_no']: r for r in rec_1d}

print("=== PART LEVEL DIFFERENCES ===")
for p in sorted(list(set(parts_1a.keys()) | set(parts_1d.keys()))):
    a = parts_1a.get(p)
    d = parts_1d.get(p)
    if a and d:
        if a['qty'] != d['qty']:
            print(f"QTY CHANGED for Part: {p}")
            print(f"   Series 1A: Qty = {a['qty']}, Locations = {a['locations']}")
            print(f"   Series 1D: Qty = {d['qty']}, Locations = {d['locations']}")
    elif a and not d:
        print(f"ONLY IN 1A: Part {p}, Qty = {a['qty']}")
    elif d and not a:
        print(f"ONLY IN 1D: Part {p}, Qty = {d['qty']}")
