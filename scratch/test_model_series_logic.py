import re, os

def extract_model_and_series(pdf_path, pwb_code="", process_code="", variant_code=""):
    """
    Extracts (base_model, series_suffix) from a PDF by combining:
    1. Filename patterns (e.g., 3275-2A, 3275-2B, CAR3072_2A, etc.)
    2. Embedded PDF metadata (PWB Code, Process Code, Model/Variant code)
    """
    fname = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # 1. Clean filename (strip common prefixes/suffixes like BOM_, WM_, Working_Manual_, etc.)
    clean_fn = re.sub(r'^(?:BOM|WM|Working_Manual|Manual|Model)[_\-\s]+', '', fname, flags=re.IGNORECASE)
    clean_fn = re.sub(r'[_\-\s]+(?:BOM|WM|Drawing|Manual)$', '', clean_fn, flags=re.IGNORECASE)
    
    # 2. Try parsing filename for Model + Series:
    # Pattern: <BaseModel>[-_]?(?:REV|SERIES|VER)?([0-9]*[A-Za-z]+|[0-9]+)$
    # Examples:
    #   "3275-2A" -> Base: "3275-2", Series: "A" (or "2A" if "3275-2" is base)
    #   "3275-2B" -> Base: "3275-2", Series: "B"
    #   "3272-2A" vs "3275-2B"
    #   "MODEL3072_REV_A" -> Base: "MODEL3072", Series: "A"
    #   "CAR3072-2A" -> Base: "CAR3072-2", Series: "A"
    fn_base = None
    fn_series = None
    
    # Match patterns like 3275-2A or 3275-2B (ending with digit+letter, or dash+letter)
    m_split = re.search(r'^(.*?)(?:[-_]?(?:REV|SERIES|VER|CODE)?[-_]?([0-9]*[A-Z]))$', clean_fn, re.IGNORECASE)
    if m_split and len(m_split.group(1)) >= 2:
        candidate_base = m_split.group(1).rstrip('-_')
        candidate_series = m_split.group(2).upper()
        # If candidate_base ends with a dash and number (like 3275-2), keep 3275-2 as base, series as A
        m_dash_num = re.search(r'^(.*?-\d+)([A-Z])$', clean_fn, re.IGNORECASE)
        if m_dash_num:
            fn_base = m_dash_num.group(1).upper()
            fn_series = m_dash_num.group(2).upper()
        else:
            fn_base = candidate_base.upper()
            fn_series = candidate_series.upper()
    
    # 3. Try parsing Process Code (e.g., CHA3072AR-2A -> base: CHA3072AR, series: 2A)
    proc_base = None
    proc_series = None
    if process_code and process_code != "PROGRAM":
        m_pr = re.search(r'^(.*?)[-_]?([0-9]+[A-Z]|[A-Z])$', process_code, re.IGNORECASE)
        if m_pr:
            proc_base = m_pr.group(1).upper()
            proc_series = m_pr.group(2).upper()
        else:
            proc_base = process_code.upper()
            proc_series = ""

    # 4. Synthesize final model and series
    # Priority:
    # If filename is a generic hash/random (like media_1788593034506):
    is_generic_fn = bool(re.match(r'^media_\d+$', fname, re.IGNORECASE) or re.match(r'^\d{8,}$', fname))
    
    if not is_generic_fn and fn_base and fn_series:
        final_model = fn_base
        final_series = fn_series
    else:
        # Use PWB code / process code
        final_model = pwb_code.upper() if pwb_code else (proc_base or fname)
        final_series = variant_code.upper() if variant_code else (proc_series or (fn_series or ""))

    return {
        "raw_filename": fname,
        "base_model": final_model,
        "series": final_series,
        "pwb_code": pwb_code,
        "process_code": process_code,
        "fn_base": fn_base,
        "fn_series": fn_series
    }

def validate_model_pair(info_a, info_b):
    """
    Validates that:
    1. Both files have the SAME base model.
    2. Both files have DIFFERENT series/revisions.
    Returns (is_valid: bool, error_msg: str)
    """
    # 1. Check same model:
    # Compare either base_model, or pwb_code, or fn_base
    same_model = False
    
    # If both have pwb_code:
    if info_a["pwb_code"] and info_b["pwb_code"]:
        same_model = (info_a["pwb_code"].upper() == info_b["pwb_code"].upper())
    elif info_a["fn_base"] and info_b["fn_base"]:
        same_model = (info_a["fn_base"].upper() == info_b["fn_base"].upper())
    else:
        same_model = (info_a["base_model"].upper() == info_b["base_model"].upper())
    
    if not same_model:
        model_a_name = info_a["fn_base"] or info_a["pwb_code"] or info_a["base_model"]
        model_b_name = info_b["fn_base"] or info_b["pwb_code"] or info_b["base_model"]
        return False, (
            f"Hai bản vẽ KHÔNG CÙNG MODEL!\n\n"
            f"• File A: Model '{model_a_name}'\n"
            f"• File B: Model '{model_b_name}'\n\n"
            f"Hệ thống chỉ so sánh giữa 2 bản vẽ có cùng Model nhưng khác Series."
        )

    # 2. Check different series:
    series_a = info_a["series"] or info_a["fn_series"] or info_a["process_code"]
    series_b = info_b["series"] or info_b["fn_series"] or info_b["process_code"]

    # Check if they are identical
    if info_a["raw_filename"] == info_b["raw_filename"]:
        return False, (
            f"Bạn đã chọn 2 file HOÀN TOÀN TRÙNG NHAU!\n\n"
            f"Vui lòng chọn 2 bản vẽ có Series khác nhau để so sánh."
        )

    if series_a and series_b and series_a.upper() == series_b.upper():
        return False, (
            f"Hai bản vẽ TRÙNG SERIES ('{series_a}')!\n\n"
            f"• File A: {info_a['raw_filename']} (Series {series_a})\n"
            f"• File B: {info_b['raw_filename']} (Series {series_b})\n\n"
            f"Vui lòng chọn 2 bản vẽ khác Series (ví dụ: Series A và Series B) để đối chiếu sai lệch."
        )

    return True, ""


# Test with test cases
test_cases = [
    # Case 1: 3275-2A vs 3275-2B (User's exact example)
    ("3275-2A.pdf", "", "", "", "3275-2B.pdf", "", "", "", True),
    # Case 2: 3275-2A vs 3275-2A (Same series -> invalid)
    ("3275-2A.pdf", "", "", "", "3275-2A.pdf", "", "", "", False),
    # Case 3: 3275-2A vs 4500-1A (Different models -> invalid)
    ("3275-2A.pdf", "", "", "", "4500-1A.pdf", "", "", "", False),
    # Case 4: Real samples 506 (CHA3072AR-2A) vs 536 (CHA3072AR-2G)
    ("media_1788593034506.pdf", "QPWBCAR3072", "CHA3072AR-2A", "", "media_1788593034536.pdf", "QPWBCAR3072", "CHA3072AR-2G", "", True),
    # Case 5: Real samples 506 (CHA3072AR-2A) vs 521 (CHA3072AR-2A) (Same series -> invalid)
    ("media_1788593034506.pdf", "QPWBCAR3072", "CHA3072AR-2A", "", "media_1788593034521.pdf", "QPWBCAR3072", "CHA3072AR-2A", "", False),
    # Case 6: Real sample 488 (TRM3280) vs 506 (CAR3072) (Different models -> invalid)
    ("media_1788593034488.pdf", "PV-QPWBCTRM3280ES", "CHP3280TRM-1A", "", "media_1788593034506.pdf", "QPWBCAR3072", "CHA3072AR-2A", "", False),
]

print("Running test cases:")
for i, (f_a, pwb_a, pr_a, v_a, f_b, pwb_b, pr_b, v_b, expected_valid) in enumerate(test_cases, 1):
    info_a = extract_model_and_series(f_a, pwb_a, pr_a, v_a)
    info_b = extract_model_and_series(f_b, pwb_b, pr_b, v_b)
    valid, msg = validate_model_pair(info_a, info_b)
    status = "OK" if valid == expected_valid else "FAIL"
    print(f"Case {i}: [{status}] Expected {expected_valid}, Got {valid}")
    if not valid:
        first_line = msg.split('\n')[0]
        print(f"   Message: {first_line.encode('ascii', errors='replace').decode()}")
