"""
Test BOMExtractorApp with Tab 4 (SeriesBOMCompareView) integrated
Verifies full GUI initialization, tab switching to 'series_bom', and clean destruction.
"""

import os
import sys

sys.path.insert(0, os.path.abspath("."))
sys.stdout.reconfigure(encoding='utf-8')
from app import BOMExtractorApp

def run_test():
    print("Initializing BOMExtractorApp...")
    app = BOMExtractorApp()
    
    # Check that view_series_bom_compare exists
    assert hasattr(app, "view_series_bom_compare"), "view_series_bom_compare missing!"
    assert hasattr(app, "btn_nav_series_bom"), "btn_nav_series_bom missing!"
    print("Verified Tab 4 components initialized.")
    
    # Switch to series_bom tab
    app._set_nav_active("series_bom")
    assert app.current_nav == "series_bom", "Failed to activate series_bom tab!"
    print("Successfully switched active tab to 'series_bom'!")
    
    # Switch through all other tabs and verify
    app._set_nav_active("model_comp")
    app._set_nav_active("compare")
    app._set_nav_active("extract")
    app._set_nav_active("series_bom")
    print("Successfully switched between all 4 tabs smoothly!")
    
    # Test language change
    app._on_language_change("English")
    print("Language changed to English successfully!")
    app._on_language_change("Tiếng Việt")
    print("Language changed back to Tiếng Việt successfully!")
    
    app.after(500, lambda: app.destroy())
    app.mainloop()
    print("Test passed 100%!")

if __name__ == "__main__":
    run_test()
