import os, sys, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if sys.platform == 'win32':
    for p in [r'C:\Program Files\Common Files\microsoft shared\ClickToRun', r'C:\Program Files\Microsoft Office\root\Client']:
        if os.path.exists(p):
            try: os.add_dll_directory(p)
            except: pass
import pymupdf as fitz
from bom_extractor import extract_metadata

def parse_filename_model_series(filename: str) -> tuple[str, str]:
    """
    Extracts (base_model, series) from filename.
    Examples:
      '3275-2A.pdf' -> ('3275-2', 'A')
      '3275-2B.pdf' -> ('3275-2', 'B')
      'MODEL_3275-2A.pdf' -> ('3275-2', 'A')
      'CAR3072_RevB.pdf' -> ('CAR3072', 'B')
      'CHA3072AR-2A.pdf' -> ('CHA3072AR', '2A')
    """
    name = os.path.splitext(os.path.basename(filename))[0]
    clean = re.sub(r'^(?:BOM|WM|Working_Manual|Manual|Model[_\s\-]*[AB]?)[_\-\s]+', '', name, flags=re.IGNORECASE)
    clean = re.sub(r'[_\-\s]+(?:BOM|WM|Drawing|Manual|Final|v\d+)$', '', clean, flags=re.IGNORECASE)
    
    # Check for pattern like <name>-<num><letter> e.g. 3275-2A
    m_dash = re.search(r'^(.*?-\d+)([A-Za-z])$', clean)
    if m_dash:
        return m_dash.group(1).upper(), m_dash.group(2).upper()
        
    # Check for pattern like <name>-[0-9]*[A-Za-z] e.g. CHA3072AR-2A or CAR3072-A
    m_code = re.search(r'^(.*?)[-_]([0-9]+[A-Za-z]|[A-Za-z])$', clean)
    if m_code and len(m_code.group(1)) >= 2:
        return m_code.group(1).upper(), m_code.group(2).upper()

    # Check for pattern like CAR3072_RevA
    m_rev = re.search(r'^(.*?)[-_]?(?:REV|SERIES|VER)[-_]?([0-9]*[A-Za-z])$', clean, re.IGNORECASE)
    if m_rev and len(m_rev.group(1)) >= 2:
        return m_rev.group(1).rstrip('-_').upper(), m_rev.group(2).upper()

    # Generic: ends with letter e.g. MODEL3275A
    m_end = re.search(r'^(.*?\d+)([A-Za-z])$', clean)
    if m_end and len(m_end.group(1)) >= 2:
        return m_end.group(1).upper(), m_end.group(2).upper()

    return clean.upper(), ""

def extract_model_info(pdf_path: str) -> dict:
    fname = os.path.splitext(os.path.basename(pdf_path))[0]
    fn_model, fn_series = parse_filename_model_series(fname)
    
    pwb_code = ""
    process_code = ""
    model_text = ""
    variant_code = ""

    if os.path.exists(pdf_path):
        try:
            doc = fitz.open(pdf_path)
            for pidx in range(min(3, len(doc))):
                txt = doc[pidx].get_text("text")
                meta = extract_metadata(txt)
                if meta.get("pwb_code") and not pwb_code:
                    pwb_code = meta["pwb_code"]
                if meta.get("process_code") and not process_code and meta["process_code"] != "PROGRAM":
                    process_code = meta["process_code"]
                if meta.get("model") and not model_text:
                    model_text = meta["model"].split('\n')[0].strip()
                m_var = re.search(r'MODEL\s*CODE\s*["\']([A-Z0-9]+)["\']', txt, re.IGNORECASE)
                if m_var and not variant_code:
                    variant_code = m_var.group(1).strip()
        except Exception as e:
            print(f"Error extracting metadata from {pdf_path}: {e}")

    # Determine base model and series
    is_generic_fn = bool(re.match(r'^(?:media_\d+|\d{8,}|scan.*|doc.*)$', fname, re.IGNORECASE))
    
    if not is_generic_fn and fn_series:
        display_model = fn_model
        display_series = fn_series
    else:
        display_model = pwb_code or fn_model
        # Process code series: e.g. CHA3072AR-2A -> 2A
        proc_series = ""
        if process_code:
            m_ps = re.search(r'[-_]([0-9]*[A-Z])$', process_code)
            proc_series = m_ps.group(1) if m_ps else process_code
        display_series = variant_code or proc_series or fn_series or ""

    return {
        "filepath": pdf_path,
        "filename": os.path.basename(pdf_path),
        "display_model": display_model,
        "display_series": display_series,
        "pwb_code": pwb_code,
        "process_code": process_code,
        "fn_model": fn_model,
        "fn_series": fn_series,
        "is_generic_fn": is_generic_fn
    }

def validate_model_pair(pdf_path_a: str, pdf_path_b: str) -> tuple[bool, str, dict, dict]:
    if not pdf_path_a or not pdf_path_b:
        return False, "Vui lòng chọn cả 2 file Model A và Model B.", {}, {}
        
    info_a = extract_model_info(pdf_path_a)
    info_b = extract_model_info(pdf_path_b)
    
    # 1. Check duplicate identical file
    if os.path.abspath(pdf_path_a) == os.path.abspath(pdf_path_b):
        return False, (
            "Bạn đang chọn 2 file HOÀN TOÀN GIỐNG NHAU!\n\n"
            "Vui lòng chọn 2 bản vẽ có Series khác nhau để so sánh."
        ), info_a, info_b

    # 2. Check Same Model:
    # Match by PWB code if both available
    same_model = False
    if info_a["pwb_code"] and info_b["pwb_code"]:
        same_model = (info_a["pwb_code"].upper() == info_b["pwb_code"].upper())
    elif not info_a["is_generic_fn"] and not info_b["is_generic_fn"]:
        same_model = (info_a["fn_model"].upper() == info_b["fn_model"].upper())
    else:
        same_model = (info_a["display_model"].upper() == info_b["display_model"].upper())

    if not same_model:
        m_a = info_a["display_model"]
        m_b = info_b["display_model"]
        return False, (
            f"Hai bản vẽ KHÔNG CÙNG MODEL!\n\n"
            f"• File A: Model '{m_a}'\n"
            f"• File B: Model '{m_b}'\n\n"
            f"Yêu cầu: Hai bản vẽ phải CÙNG MODEL nhưng KHÁC SERIES (ví dụ: cùng model 3275-2 nhưng là series 3275-2A và 3275-2B)."
        ), info_a, info_b

    # 3. Check Different Series:
    ser_a = info_a["display_series"].upper()
    ser_b = info_b["display_series"].upper()

    if ser_a and ser_b and ser_a == ser_b:
        return False, (
            f"Hai bản vẽ TRÙNG SERIES ('{ser_a}')!\n\n"
            f"• File A: {info_a['filename']} (Series: {ser_a})\n"
            f"• File B: {info_b['filename']} (Series: {ser_b})\n\n"
            f"Yêu cầu: Hai bản vẽ phải KHÁC SERIES (ví dụ: Series 2A và Series 2B) để phát hiện sự sai khác."
        ), info_a, info_b

    return True, "", info_a, info_b

# Run verification tests on sample manuals and synthetic paths
print("Testing validate_model_pair:")
# Sample 506 (2A) vs 536 (2G)
v, msg, a, b = validate_model_pair("sample_manuals/media_1788593034506.pdf", "sample_manuals/media_1788593034536.pdf")
print("Test 1 (506 2A vs 536 2G):", "PASS" if v else "FAIL", f"(Model: {a['display_model']}, SerA: {a['display_series']}, SerB: {b['display_series']})")

# Sample 506 (2A) vs 521 (2A) -> Should FAIL (Same Series)
v, msg, a, b = validate_model_pair("sample_manuals/media_1788593034506.pdf", "sample_manuals/media_1788593034521.pdf")
print("Test 2 (506 2A vs 521 2A - same series):", "PASS (Correctly rejected)" if not v else "FAIL")

# Sample 488 (TRM3280) vs 506 (CAR3072) -> Should FAIL (Different Models)
v, msg, a, b = validate_model_pair("sample_manuals/media_1788593034488.pdf", "sample_manuals/media_1788593034506.pdf")
print("Test 3 (TRM3280 vs CAR3072 - different models):", "PASS (Correctly rejected)" if not v else "FAIL")

# Synthetic: 3275-2A.pdf vs 3275-2B.pdf
v, msg, a, b = validate_model_pair("3275-2A.pdf", "3275-2B.pdf")
print("Test 4 (3275-2A vs 3275-2B):", "PASS" if v else "FAIL", f"(Model: {a['display_model']}, SerA: {a['display_series']}, SerB: {b['display_series']})")

# Synthetic: 3275-2A.pdf vs 3275-2A.pdf -> Should FAIL (Same)
v, msg, a, b = validate_model_pair("3275-2A.pdf", "3275-2A.pdf")
print("Test 5 (3275-2A vs 3275-2A - identical):", "PASS (Correctly rejected)" if not v else "FAIL")

# Synthetic: 3275-2A.pdf vs 3272-2A.pdf -> Should FAIL (Different models 3275-2 vs 3272-2)
v, msg, a, b = validate_model_pair("3275-2A.pdf", "3272-2A.pdf")
print("Test 6 (3275-2A vs 3272-2A - different models):", "PASS (Correctly rejected)" if not v else "FAIL")
