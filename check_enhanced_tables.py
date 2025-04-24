#!/usr/bin/env python3
"""
Check if the TECHNICAL DETAILS, OVERVIEW, and REPRODUCIBILITY tables are properly populated
in the output document. This verifies that our enhanced template populator is correctly
filling in the tables with extracted data.
"""

import os
from pathlib import Path
from docx import Document
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def check_tables(document_path="output_populated_template.docx"):
    """
    Check if the tables in the document are properly populated.
    
    Args:
        document_path: Path to the document to check
        
    Note:
        Table indices in the document (in order):
        0 - Technical Details Table
        1 - Overview Table 
        2 - Kit Components Table
        3 - Standard Curve Table
        4 - Intra-Assay Table
        5 - Inter-Assay Table
        6 - Lot-to-Lot Table
    """
    if not os.path.exists(document_path):
        logger.error(f"Document not found at {document_path}")
        return
    
    logger.info(f"Checking tables in {document_path}")
    doc = Document(document_path)
    
    # Variables to track our findings
    found_technical_details_table = False
    found_overview_table = False
    found_reproducibility_tables = False
    technical_details_populated = False
    overview_populated = False
    reproducibility_populated = False
    
    # Find sections by heading
    technical_details_section = None
    overview_section = None
    reproducibility_section = None
    
    for i, para in enumerate(doc.paragraphs):
        if "TECHNICAL DETAILS" in para.text.strip().upper():
            technical_details_section = i
            logger.info(f"Found TECHNICAL DETAILS section at paragraph {i}")
        elif "OVERVIEW" in para.text.strip().upper():
            overview_section = i
            logger.info(f"Found OVERVIEW section at paragraph {i}")
        elif "REPRODUCIBILITY" in para.text.strip().upper():
            reproducibility_section = i
            logger.info(f"Found REPRODUCIBILITY section at paragraph {i}")
    
    # Check all tables to identify and validate our target tables
    for i, table in enumerate(doc.tables):
        if not table.rows:
            continue
        
        # Extract table content for analysis
        table_content = ""
        empty_cells = 0
        total_cells = 0
        
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                table_content += cell_text + " "
                total_cells += 1
                if not cell_text or cell_text == "N/A":
                    empty_cells += 1
        
        table_content = table_content.lower()
        
        # Check for technical details table
        if 'sensitivity' in table_content and 'detection range' in table_content:
            found_technical_details_table = True
            logger.info(f"Found technical details table at index {i}")
            
            # Check if values are filled in
            empty_value_cells = 0
            for row in table.rows:
                if len(row.cells) >= 2:
                    value_cell = row.cells[1].text.strip()
                    if not value_cell or value_cell == "N/A":
                        empty_value_cells += 1
            
            # Calculate percentage of empty cells
            if len(table.rows) > 0:
                empty_percentage = (empty_value_cells / len(table.rows)) * 100
                logger.info(f"Technical details table has {empty_percentage:.1f}% empty value cells")
                
                if empty_percentage < 50:  # Less than 50% empty is considered populated
                    technical_details_populated = True
                    logger.info("Technical details table is adequately populated")
                else:
                    logger.warning("Technical details table has too many empty cells")
        
        # Check for overview table
        elif 'product' in table_content and ('species' in table_content or 'reactive' in table_content):
            found_overview_table = True
            logger.info(f"Found overview table at index {i}")
            
            # Check if values are filled in
            empty_value_cells = 0
            for row in table.rows:
                if len(row.cells) >= 2:
                    value_cell = row.cells[1].text.strip()
                    if not value_cell or value_cell == "N/A":
                        empty_value_cells += 1
            
            # Calculate percentage of empty cells
            if len(table.rows) > 0:
                empty_percentage = (empty_value_cells / len(table.rows)) * 100
                logger.info(f"Overview table has {empty_percentage:.1f}% empty value cells")
                
                if empty_percentage < 50:  # Less than 50% empty is considered populated
                    overview_populated = True
                    logger.info("Overview table is adequately populated")
                else:
                    logger.warning("Overview table has too many empty cells")
        
        # Check for reproducibility tables
        elif ('intra-assay' in table_content or 'inter-assay' in table_content or 'cv' in table_content) and reproducibility_section is not None:
            # This is likely a reproducibility table
            if not found_reproducibility_tables:
                found_reproducibility_tables = True
                logger.info(f"Found reproducibility table at index {i}")
            
            # Check if values are filled in for all rows
            empty_cells_in_table = 0
            cell_count = 0
            
            for row_idx, row in enumerate(table.rows):
                if row_idx == 0:  # Skip header row
                    continue
                    
                for cell in row.cells:
                    cell_count += 1
                    if not cell.text.strip():
                        empty_cells_in_table += 1
            
            # Calculate percentage of empty cells
            if cell_count > 0:
                empty_percentage = (empty_cells_in_table / cell_count) * 100
                logger.info(f"Reproducibility table has {empty_percentage:.1f}% empty cells")
                
                if empty_percentage < 20:  # Less than 20% empty is considered populated
                    reproducibility_populated = True
                    logger.info("Reproducibility table is adequately populated")
                else:
                    logger.warning("Reproducibility table has too many empty cells")
    
    # Report overall status
    print("\n=== Table Population Status ===")
    print(f"TECHNICAL DETAILS Table: {'Found' if found_technical_details_table else 'Not found'}, {'Populated' if technical_details_populated else 'Not fully populated'}")
    print(f"OVERVIEW Table: {'Found' if found_overview_table else 'Not found'}, {'Populated' if overview_populated else 'Not fully populated'}")
    print(f"REPRODUCIBILITY Tables: {'Found' if found_reproducibility_tables else 'Not found'}, {'Populated' if reproducibility_populated else 'Not fully populated'}")
    
    if found_technical_details_table and found_overview_table and found_reproducibility_tables and technical_details_populated and overview_populated and reproducibility_populated:
        print("\n✅ All tables are found and properly populated!")
    else:
        print("\n❌ Some tables are missing or not fully populated.")
        
        if not found_technical_details_table:
            print("   - TECHNICAL DETAILS table not found")
        elif not technical_details_populated:
            print("   - TECHNICAL DETAILS table has empty cells")
            
        if not found_overview_table:
            print("   - OVERVIEW table not found")
        elif not overview_populated:
            print("   - OVERVIEW table has empty cells")
            
        if not found_reproducibility_tables:
            print("   - REPRODUCIBILITY tables not found")
        elif not reproducibility_populated:
            print("   - REPRODUCIBILITY tables have empty cells")

if __name__ == "__main__":
    check_tables()