#!/usr/bin/env python3
"""
Check the TECHNICAL DETAILS table in the output document.
"""

import logging
from pathlib import Path
from docx import Document

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def check_technical_details(document_path="output_populated_template.docx"):
    """
    Check the content of the TECHNICAL DETAILS table and identify which cells are empty.
    
    Args:
        document_path: Path to the document to check
    """
    doc = Document(document_path)
    
    print("\n=== TECHNICAL DETAILS TABLE CONTENT ===")
    
    # Find the Technical Details table (it should be the first table)
    technical_details_table = None
    for i, table in enumerate(doc.tables):
        table_content = ""
        for row in table.rows:
            for cell in row.cells:
                table_content += cell.text.lower() + " "
                
        if "capture" in table_content and "antibod" in table_content:
            technical_details_table = table
            print(f"Found Technical Details table at index {i}")
            break
    
    if technical_details_table is None:
        print("Technical Details table not found!")
        return
    
    # Check table content
    empty_cells = 0
    total_cells = 0
    
    for i, row in enumerate(technical_details_table.rows):
        # Ensure the row has at least 2 cells
        if len(row.cells) >= 2:
            property_cell = row.cells[0].text.strip()
            value_cell = row.cells[1].text.strip()
            
            print(f"Row {i}: '{property_cell}': '{value_cell}'")
            
            total_cells += 1
            if not value_cell or value_cell == "N/A":
                empty_cells += 1
    
    # Calculate percentage of empty cells
    if total_cells > 0:
        empty_percentage = (empty_cells / total_cells) * 100
        print(f"\nTechnical Details table has {empty_percentage:.1f}% empty cells ({empty_cells}/{total_cells})")
    else:
        print("\nTechnical Details table has no rows to analyze")

if __name__ == "__main__":
    check_technical_details()