import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app

from unittest.mock import patch, MagicMock

def test_gui_validation():
    print("Initializing BOMExtractorApp for validation GUI test...")
    my_app = app.BOMExtractorApp()
    my_app.withdraw()  # Don't show actual window

    tab3 = my_app.view_model_compare

    f_2g = os.path.abspath(r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788678922426.pdf") # CAR3072 2G
    f_2a = os.path.abspath(r"C:\Users\QC60\.gemini\antigravity-ide\brain\78faf423-be84-4719-bb09-30bb52694765\.user_uploaded\media_1788678927420.pdf") # CAR3072 2A
    f_488 = os.path.abspath(r"media_1788593034488.pdf") if os.path.exists("media_1788593034488.pdf") else None

    # Mock messagebox to avoid blocking modal dialogs
    with patch("tkinter.messagebox.showerror") as mock_err, \
         patch("tkinter.messagebox.showwarning") as mock_warn:

        # Step 1: Select Model A (2G)
        with patch("tkinter.filedialog.askopenfilename", return_value=f_2g):
            tab3._select_pdf_a()

        assert tab3.pdf_a_path == os.path.normpath(f_2g), f"Expected pdf_a_path to be set, got {tab3.pdf_a_path}"
        lbl_a = tab3.lbl_model_a_sub.cget("text")
        print(f"Card A after selecting 2G: '{lbl_a}'")
        assert "QPWBCAR3072" in lbl_a and "2G" in lbl_a

        # Step 2: Try to select Model B with same series (2G again)
        with patch("tkinter.filedialog.askopenfilename", return_value=f_2g):
            tab3._select_pdf_b()

        assert tab3.pdf_b_path is None, "Model B should not be set when selecting identical/same series!"
        assert mock_err.called, "showerror should be called for same series!"
        print("Test Step 2: Same series correctly blocked with error dialog!")

        # Step 3: Try to select Model B with different model (if 488 exists)
        if f_488:
            mock_err.reset_mock()
            with patch("tkinter.filedialog.askopenfilename", return_value=f_488):
                tab3._select_pdf_b()
            assert tab3.pdf_b_path is None, "Model B should not be set when different model!"
            assert mock_err.called, "showerror should be called for different model!"
            print("Test Step 3: Different model correctly blocked with error dialog!")

        # Step 4: Select Model B with valid pair (f_2a - series 2A)
        mock_err.reset_mock()
        with patch("tkinter.filedialog.askopenfilename", return_value=f_2a):
            tab3._select_pdf_b()

        assert tab3.pdf_b_path == os.path.normpath(f_2a), "Model B should be set for valid pair!"
        lbl_b = tab3.lbl_model_b_sub.cget("text")
        print(f"Card B after selecting 2A: '{lbl_b}'")
        assert "QPWBCAR3072" in lbl_b and "2A" in lbl_b
        print("Test Step 4: Valid pair successfully loaded!")

        # Step 5: Test handle_drop_files with 2 files (2G & 2A -> valid)
        tab3._reset()
        assert tab3.pdf_a_path is None and tab3.pdf_b_path is None
        tab3.handle_drop_files([f_2g, f_2a])
        assert tab3.pdf_a_path == os.path.normpath(f_2g)
        assert tab3.pdf_b_path == os.path.normpath(f_2a)
        print("Test Step 5: Valid 2-file drag-and-drop succeeded!")

        # Step 6: Test handle_drop_files with invalid pair (same file twice)
        mock_err.reset_mock()
        tab3._reset()
        tab3.handle_drop_files([f_2g, f_2g])
        assert tab3.pdf_a_path is None and tab3.pdf_b_path is None
        assert mock_err.called, "showerror should be called for 2 identical files dropped!"
        print("Test Step 6: Invalid 2-file drag-and-drop correctly blocked!")

    my_app.destroy()
    print("ALL GUI VALIDATION TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    test_gui_validation()
