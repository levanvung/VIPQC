import os, sys, re, glob
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if sys.platform == 'win32':
    for p in [r'C:\Program Files\Common Files\microsoft shared\ClickToRun', r'C:\Program Files\Microsoft Office\root\Client']:
        if os.path.exists(p):
            try: os.add_dll_directory(p)
            except: pass
import pymupdf
from bom_extractor import extract_metadata

def get_file_model_info(pdf_path):
    fn = os.path.splitext(os.path.basename(pdf_path))[0]
    pwb = ''
    proc = ''
    model_text = ''
    variant_code = ''
    try:
        doc = pymupdf.open(pdf_path)
        for i in range(min(3, len(doc))):
            txt = doc[i].get_text('text')
            m = extract_metadata(txt)
            if m.get('pwb_code') and not pwb:
                pwb = m['pwb_code']
            if m.get('process_code') and not proc and m['process_code'] != 'PROGRAM':
                proc = m['process_code']
            if m.get('model') and not model_text:
                model_text = m['model']
            m_code = re.search(r'MODEL\s*CODE\s*["\']([A-Z0-9]+)["\']', txt, re.IGNORECASE)
            if m_code:
                variant_code = m_code.group(1).strip()
    except Exception as e:
        print(f'Error reading {pdf_path}: {e}')
    
    return {
        'filename': fn,
        'pwb_code': pwb,
        'process_code': proc,
        'model_text': model_text,
        'variant_code': variant_code
    }

for f in sorted(glob.glob('sample_manuals/*.pdf')):
    info = get_file_model_info(f)
    print(os.path.basename(f), '->', info)
