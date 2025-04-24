#!/usr/bin/env python3
"""
Check the Kit Components table in the output document.
"""

import logging
from docx import Document
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def check_kit_components(output_path="output_populated_template.docx"):
    """Check the kit components table in the output document."""
    doc = Document(output_path)
    logger.info(f"Checking kit components table in document at {output_path}")
    
    # Table 1 should be the kit components table based on our template
    if len(doc.tables) < 1:
        logger.error("No tables found in document")
        return
        
    # Get the first table
    kit_table = doc.tables[0]
    rows = len(kit_table.rows)
    cols = len(kit_table.columns)
    logger.info(f"Kit Components Table: {rows} rows x {cols} columns")
    
    # Print the table contents
    print("\nKit Components Table Contents:")
    print("-" * 50)
    for row_idx, row in enumerate(kit_table.rows):
        cells = [cell.text.strip() for cell in row.cells]
        print(f"Row {row_idx+1}: {cells}")
        
    # Check for reagent rows with content
    filled_rows = 0
    for row_idx in range(1, rows): # Skip header row
        if row_idx < len(kit_table.rows):
            row = kit_table.rows[row_idx]
            if row.cells[0].text.strip():
                filled_rows += 1
                
    logger.info(f"Found {filled_rows} filled reagent rows")
    print(f"\nFilled reagent rows: {filled_rows}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        check_kit_components(sys.argv[1])
    else:
        check_kit_components()