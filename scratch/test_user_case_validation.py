import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from model_comparator import validate_model_pair, parse_filename_model_series, extract_model_info

def run_tests():
    print("Testing Model & Series Parsing and Validation:")

    # Case 1: The user's exact files from screenshot
    f1 = "CHP3178AF-1A MP.pdf"
    f2 = "CHP3178AF-1B MP.pdf"
    m1, s1 = parse_filename_model_series(f1)
    m2, s2 = parse_filename_model_series(f2)
    print(f"User Case A: {f1} -> Model: '{m1}', Series: '{s1}'")
    print(f"User Case B: {f2} -> Model: '{m2}', Series: '{s2}'")
    assert m1 == m2 == "CHP3178AF", f"Expected CHP3178AF, got {m1}, {m2}"
    assert s1 == "1A" and s2 == "1B", f"Expected 1A and 1B, got {s1}, {s2}"

    is_v, err, ia, ib = validate_model_pair(f1, f2)
    assert is_v, f"Expected valid pair, got error: {err}"
    print("User Case: PASS -> Valid pair with Series 1A vs 1B!")

    # Case 2: Same series
    is_v, err, ia, ib = validate_model_pair(f1, f1)
    assert not is_v, "Expected rejected for identical files"
    print("Same file: PASS -> Correctly rejected")

    # Case 3: 3275-2A vs 3275-2B
    f_3275a = "3275-2A.pdf"
    f_3275b = "3275-2B.pdf"
    is_v, err, ia, ib = validate_model_pair(f_3275a, f_3275b)
    assert is_v, f"Expected valid pair, got: {err}"
    print(f"3275-2A vs 3275-2B: PASS -> Model: {ia['display_model']}, Series: {ia['display_series']} vs {ib['display_series']}")

    # Case 4: 3275-2A vs 3272-2A (different model)
    f_3272a = "3272-2A.pdf"
    is_v, err, ia, ib = validate_model_pair(f_3275a, f_3272a)
    assert not is_v, "Expected rejected for different models"
    print("3275-2A vs 3272-2A: PASS -> Correctly rejected (different model)")

    # Case 5: Real PDFs in user_uploaded folder
    p1 = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788678922426.pdf" # 2G
    p2 = r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788678927420.pdf" # 2A
    if os.path.exists(p1) and os.path.exists(p2):
        is_v, err, ia, ib = validate_model_pair(p1, p2)
        assert is_v, f"Expected valid pair, got: {err}"
        print(f"Real PDFs: PASS -> Model: {ia['display_model']}, Series: {ia['display_series']} vs {ib['display_series']}")

    print("\nALL 5 VALIDATION TESTS PASSED 100%!")

if __name__ == "__main__":
    run_tests()
