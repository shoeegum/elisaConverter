#!/usr/bin/env python3
"""
Check the Preparations Before Assay section in the output document
to confirm that numbered lists are properly preserved.
"""

from docx import Document

def check_preparations_section(document_path):
    """Check the preparations section for numbered lists."""
    doc = Document(document_path)
    
    print(f"Checking Preparations Before Assay section in {document_path}:")
    print("-" * 80)
    
    found_numbered_lists = False
    in_section = False
    
    for para in doc.paragraphs:
        if 'PREPARATIONS BEFORE ASSAY' in para.text.upper():
            in_section = True
            print("\nFound section!")
        elif in_section and any(s in para.text.upper() for s in ['KIT COMPONENTS', 'MATERIALS PROVIDED']):
            in_section = False
            print("\nEnd of section.")
        elif in_section:
            para_text = para.text.strip()
            if para_text:
                print(f"Para style: {para.style.name}")
                display_text = para_text[:80] + '...' if len(para_text) > 80 else para_text
                print(f"Text: {display_text}")
                
                # Check if this is a numbered list item
                if para_text.startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.')):
                    found_numbered_lists = True
                    print(f"  --> Numbered list item found: {para_text}")
    
    print("\nSummary:")
    print(f"Found numbered lists in section: {found_numbered_lists}")
    return found_numbered_lists

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        check_preparations_section(sys.argv[1])
    else:
        check_preparations_section("output_populated_template.docx")