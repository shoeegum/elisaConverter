#!/usr/bin/env python3
"""
Check document margins in the enhanced template output
"""

import docx
import sys

def check_document_margins(filename):
    """Check the margins in the document and print details"""
    try:
        doc = docx.Document(filename)
        print(f"\nChecking margins in document: {filename}")
        
        # Check sections (each section can have different page settings)
        for i, section in enumerate(doc.sections):
            print(f"\nSection {i+1} margins:")
            print(f"  Top margin: {section.top_margin.inches:.2f} inches")
            print(f"  Bottom margin: {section.bottom_margin.inches:.2f} inches")
            print(f"  Left margin: {section.left_margin.inches:.2f} inches")
            print(f"  Right margin: {section.right_margin.inches:.2f} inches")
            print(f"  Page height: {section.page_height.inches:.2f} inches")
            print(f"  Page width: {section.page_width.inches:.2f} inches")
            
            # Check if this is "Narrow" margins
            is_narrow = (
                0.45 <= section.top_margin.inches <= 0.55 and
                0.45 <= section.bottom_margin.inches <= 0.55 and
                0.45 <= section.left_margin.inches <= 0.55 and
                0.45 <= section.right_margin.inches <= 0.55
            )
            
            if is_narrow:
                print("  ✓ This section has narrow margins (approximately 0.5 inches)")
            else:
                print("  ✗ This section does NOT have narrow margins")
                print("    Note: Narrow margins are typically around 0.5 inches on all sides")
                
        return True
    
    except Exception as e:
        print(f"Error checking document margins: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        check_document_margins(sys.argv[1])
    else:
        check_document_margins("IMSKLK1KT-20250424.docx")