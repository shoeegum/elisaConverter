#!/usr/bin/env python3
"""
Check the Preparations Before Assay section in the source document
to confirm that there are numbered lists that need to be preserved.
"""

from docx import Document

def check_preparations_section(document_path):
    """Check the preparations section for numbered lists in the source document."""
    doc = Document(document_path)
    
    print(f"Checking Preparations Before Assay section in {document_path}:")
    print("-" * 80)
    
    # Look for the preparations section
    found_section = False
    in_section = False
    found_numbered_lists = False
    all_paragraphs = []
    
    for para in doc.paragraphs:
        if "PREPARATIONS BEFORE ASSAY" in para.text.upper():
            found_section = True
            in_section = True
            print(f"Found section: '{para.text}'")
        elif in_section and any(s in para.text.upper() for s in ['KIT COMPONENTS', 'MATERIALS PROVIDED']):
            in_section = False
            print("\nEnd of section.")
        elif in_section:
            para_text = para.text.strip()
            if para_text:
                all_paragraphs.append(para_text)
                display_text = para_text[:80] + '...' if len(para_text) > 80 else para_text
                print(f"Text: {display_text}")
                
                # Check if this is a numbered list item
                if para_text.startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.')):
                    found_numbered_lists = True
                    print(f"  --> Numbered list item found: {para_text}")
    
    print("\nSummary:")
    print(f"Found section: {found_section}")
    print(f"Found numbered lists in section: {found_numbered_lists}")
    print(f"Total paragraphs in section: {len(all_paragraphs)}")
    return found_numbered_lists, all_paragraphs

if __name__ == "__main__":
    import sys
    file_path = sys.argv[1] if len(sys.argv) > 1 else "attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx"
    found, paragraphs = check_preparations_section(file_path)