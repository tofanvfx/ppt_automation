import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from lxml import etree

doc = Document('content.docx')
w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

in_section = False
section_name = ''
target_sections = ['sst_previous_page', 'sst_lo_page', 'math_lo_page', 'sst_summary_page', 'math_summary_page']

for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    
    # Check if this is a section header
    if text.startswith('[') and text.endswith(']'):
        name = text[1:-1].strip().lower()
        if name in target_sections:
            in_section = True
            section_name = name
            print(f"\n=== [{section_name}] ===")
        else:
            in_section = False
        continue
    
    if in_section and text:
        # Check paragraph properties for indentation/bullets
        pPr = para._element.find(f'{{{w_ns}}}pPr')
        numPr = para._element.find(f'.//{{{w_ns}}}numPr') if pPr is not None else None
        ilvl_elem = numPr.find(f'{{{w_ns}}}ilvl') if numPr is not None else None
        ilvl = ilvl_elem.get(f'{{{w_ns}}}val') if ilvl_elem is not None else None
        
        ind_elem = pPr.find(f'{{{w_ns}}}ind') if pPr is not None else None
        left_indent = ind_elem.get(f'{{{w_ns}}}left') if ind_elem is not None else None
        
        style_name = para.style.name if para.style else 'None'
        
        print(f"  LINE {i}: style='{style_name}' ilvl={ilvl} indent={left_indent} text={repr(text[:70])}")
