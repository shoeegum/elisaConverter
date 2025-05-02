#!/usr/bin/env python3
"""
Fix Red Dot Document Issues

This script addresses two issues with Red Dot documents:
1. Moves the REAGENTS PROVIDED table to the correct position in the document
2. Replaces all instances of 'Reddot Biotech INC.' with 'Innovative Research, Inc.'
   and 'Reddot Biotech' with 'Innovative Research'
"""

import logging
import shutil
import sys
import re
from pathlib import Path
from docx import Document

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_company_names(doc):
    """
    Replace all instances of Reddot company names with Innovative Research.
    
    Args:
        doc: The Document object to modify
    """
    replacements = [
        ('Reddot Biotech INC.', 'Innovative Research, Inc.'),
        ('Reddot Biotech', 'Innovative Research'),  # Must be after the more specific replacement
    ]
    
    count = 0
    for para in doc.paragraphs:
        for old_text, new_text in replacements:
            if old_text in para.text:
                para.text = para.text.replace(old_text, new_text)
                count += 1
    
    # Replace in tables too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for old_text, new_text in replacements:
                        if old_text in para.text:
                            para.text = para.text.replace(old_text, new_text)
                            count += 1
    
    logger.info(f"Replaced {count} instances of company names")

def fix_document_structure(document_path):
    """
    Fix document structure issues, particularly placing the REAGENTS PROVIDED table
    in the correct position.
    
    Args:
        document_path: Path to the document to modify
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create a backup of the document
        document_path = Path(document_path)
        backup_path = document_path.with_name(f"{document_path.stem}_before_fixes{document_path.suffix}")
        shutil.copy2(document_path, backup_path)
        logger.info(f"Created backup at {backup_path}")
        
        # Load the document
        doc = Document(document_path)
        
        # First replace company names
        fix_company_names(doc)
        
        # Find the REAGENTS PROVIDED section
        reagents_section_idx = None
        table_idx = None
        found_table = False
        
        # Find the section and table
        for i, para in enumerate(doc.paragraphs):
            if para.text.strip() == "REAGENTS PROVIDED":
                reagents_section_idx = i
                logger.info(f"Found REAGENTS PROVIDED section at paragraph {i}")
        
        if reagents_section_idx is None:
            logger.warning("REAGENTS PROVIDED section not found in document")
            return False
        
        # Identify all tables
        logger.info(f"Document contains {len(doc.tables)} tables")
        if len(doc.tables) > 0:
            logger.info(f"Looking for reagents table")
            # Find the table with reagents data
            for i, table in enumerate(doc.tables):
                # Check table headers
                if len(table.rows) > 0:
                    header_cells = [cell.text.strip() for cell in table.rows[0].cells]
                    logger.info(f"Table {i} headers: {header_cells}")
                    # Check if this looks like the reagents table
                    if 'Reagents' in header_cells and 'Quantity' in header_cells:
                        logger.info(f"Found reagents table at index {i}")
                        table_idx = i
                        found_table = True
                        break
        
        if not found_table:
            logger.warning("Reagents table not found in document")
            # Save changes from company name replacements
            doc.save(document_path)
            logger.info(f"Saved document with company name replacements: {document_path}")
            return True
        
        # Move the table to the correct position
        # In Word, we need to take a different approach - we'll delete the table and re-add it
        
        # Get the table data
        table = doc.tables[table_idx]
        table_data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data.append(row_data)
        
        logger.info(f"Extracted data from table: {len(table_data)} rows")
        
        # Delete the table
        # Get the table object's parent element and remove the table
        # We need to access the underlying XML to do this
        element = table._element
        element.getparent().remove(element)
        
        # Add a new table right after the REAGENTS PROVIDED heading
        # Get the paragraph after the heading
        after_heading_idx = reagents_section_idx + 1
        if after_heading_idx >= len(doc.paragraphs):
            logger.warning("No paragraph after REAGENTS PROVIDED section to insert table")
            doc.save(document_path)
            return False
        
        # Insert the table after the heading
        p = doc.paragraphs[after_heading_idx]
        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
        table.style = 'Table Grid'
        
        # Fill the table with the data we collected
        for i, row_data in enumerate(table_data):
            for j, cell_text in enumerate(row_data):
                if j < len(table.columns):  # Safety check
                    cell = table.cell(i, j)
                    cell.text = cell_text
        
        # Save the modified document
        doc.save(document_path)
        logger.info(f"Successfully moved table and replaced company names in: {document_path}")
        return True
            
    except Exception as e:
        logger.error(f"Error fixing document structure: {e}")
        return False

if __name__ == "__main__":
    # Use the provided file path or default to complete_red_dot_output.docx
    if len(sys.argv) > 1:
        document_path = sys.argv[1]
    else:
        document_path = "complete_red_dot_output.docx"
    
    # Run the fix
    fix_document_structure(document_path)