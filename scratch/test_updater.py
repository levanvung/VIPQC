import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import updater

def test_updater_logic():
    print("Testing updater logic:")
    assert updater.parse_version_tuple("v2.1.0") == (2, 1, 0)
    assert updater.parse_version_tuple("2.2") == (2, 2)
    assert updater.is_newer_version("2.1.0", "2.2.0") == True
    assert updater.is_newer_version("2.1.0", "2.1.0") == False
    assert updater.is_newer_version("2.1.0", "2.0.9") == False
    assert updater.is_newer_version("2.1.0", "v2.1.1") == True
    print("Version parsing & comparison tests PASSED!")

    # Check local version.json
    v_file = os.path.abspath("version.json")
    assert os.path.exists(v_file), "version.json must exist"
    print("version.json exists!")

    print("All updater tests PASSED successfully!")

if __name__ == "__main__":
    test_updater_logic()
