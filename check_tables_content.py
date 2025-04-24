"""
Check the content of the TECHNICAL DETAILS, OVERVIEW, and REPRODUCIBILITY tables
in the output document to identify missing values.
"""

import docx
import re

def check_tables_content(document_path="output_populated_template.docx"):
    """
    Check the content of the tables in the document to identify missing values.
    
    Args:
        document_path: Path to the document to check
    """
    # Load the document
    doc = docx.Document(document_path)
    
    print(f"The document contains {len(doc.tables)} tables.")
    
    # Find the tables for each section
    technical_details_table = None
    overview_table = None
    reproducibility_table = None
    
    # Find the tables based on their position after section headings
    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph.text.strip().upper() == "TECHNICAL DETAILS":
            # Technical Details table should be the next table after this heading
            for table_idx, table in enumerate(doc.tables):
                if is_table_after_paragraph(doc, table, i):
                    technical_details_table = table
                    print(f"Found TECHNICAL DETAILS table at index {table_idx}")
                    break
        elif paragraph.text.strip().upper() == "OVERVIEW":
            # Overview table should be the next table after this heading
            for table_idx, table in enumerate(doc.tables):
                if is_table_after_paragraph(doc, table, i):
                    overview_table = table
                    print(f"Found OVERVIEW table at index {table_idx}")
                    break
        elif "REPRODUCIBILITY" in paragraph.text.strip().upper():
            # Reproducibility table should be the next table after this heading
            for table_idx, table in enumerate(doc.tables):
                if is_table_after_paragraph(doc, table, i):
                    reproducibility_table = table
                    print(f"Found REPRODUCIBILITY table at index {table_idx}")
                    break
    
    # Check content of Technical Details table
    if technical_details_table:
        print("\n=== TECHNICAL DETAILS TABLE CONTENT ===")
        for row in technical_details_table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            print(f"Row: {cells}")
            # Check for empty or placeholder values
            for cell in cells:
                if not cell or cell.startswith('{{') or cell == 'N/A':
                    print(f"  Warning: Empty or placeholder value: '{cell}'")
    else:
        print("TECHNICAL DETAILS table not found")
    
    # Check content of Overview table
    if overview_table:
        print("\n=== OVERVIEW TABLE CONTENT ===")
        for row in overview_table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            print(f"Row: {cells}")
            # Check for empty or placeholder values
            for cell in cells:
                if not cell or cell.startswith('{{') or cell == 'N/A':
                    print(f"  Warning: Empty or placeholder value: '{cell}'")
    else:
        print("OVERVIEW table not found")
    
    # Check content of Reproducibility table
    if reproducibility_table:
        print("\n=== REPRODUCIBILITY TABLE CONTENT ===")
        for row in reproducibility_table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            print(f"Row: {cells}")
            # Check for empty or placeholder values
            for cell in cells:
                if not cell or cell.startswith('{{') or cell == 'N/A':
                    print(f"  Warning: Empty or placeholder value: '{cell}'")
    else:
        print("REPRODUCIBILITY table not found")

def is_table_after_paragraph(doc, table, paragraph_idx):
    """
    Check if a table appears after a given paragraph.
    
    Args:
        doc: The Document object
        table: The Table object to check
        paragraph_idx: The index of the paragraph to check against
        
    Returns:
        True if the table appears after the paragraph, False otherwise
    """
    # Get the XML element of the paragraph
    para_element = doc.paragraphs[paragraph_idx]._element
    
    # Get the XML element of the table
    table_element = table._element
    
    # Check if the table appears after the paragraph in the document
    current = para_element.getnext()
    while current is not None:
        if current is table_element:
            return True
        current = current.getnext()
    
    return False

if __name__ == "__main__":
    check_tables_content()