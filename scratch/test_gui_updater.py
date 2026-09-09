import os
import sys
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app
import updater

def test_gui_updater():
    print("Testing GUI updater integration...")
    app_instance = app.BOMExtractorApp()
    app_instance.withdraw()

    # Check button existence
    assert hasattr(app_instance, "btn_sidebar_update"), "btn_sidebar_update must exist"
    btn_text = app_instance.btn_sidebar_update.cget("text")
    print("Sidebar update button text verified!")
    assert f"v{updater.CURRENT_VERSION}" in btn_text

    # Mock an update response
    mock_update = {
        "has_update": True,
        "current_version": "2.1.0",
        "latest_version": "2.2.0",
        "title": "VIPQC AI v2.2.0 - Tối ưu hóa lớn",
        "changelog": "- Tính năng tự động cập nhật\n- Nâng cấp giao diện",
        "download_url": "https://example.com/VIPQC_AI.exe",
        "release_url": "https://github.com/levanvung/VIPQC"
    }

    # Open update dialog
    dlg = app.UpdateDialog(app_instance, mock_update)
    app_instance.update()
    print("UpdateDialog successfully opened and rendered!")
    dlg.destroy()

    app_instance.destroy()
    print("ALL GUI UPDATER INTEGRATION TESTS PASSED!")

if __name__ == "__main__":
    test_gui_updater()
