#!/usr/bin/env python3
"""
Test the improved Red Dot template extraction and population.

This script verifies that the enhanced extraction method properly preserves
formatting elements like numbered lists, bulleted lists, and tables.
"""

import logging
from pathlib import Path
from red_dot_template_populator import extract_red_dot_data, populate_red_dot_template

# Configure logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_red_dot_extraction():
    """Test the Red Dot extraction with formatting preservation."""
    # Use the Red Dot sample file
    source_path = Path("attached_assets/RDR-LMNB2-Hu.docx")
    
    # Extract data using the improved method
    data = extract_red_dot_data(source_path)
    
    # Print the extracted sections to verify formatting
    if 'red_dot_sections' in data:
        logger.info("Successfully extracted Red Dot sections:")
        for section, content in data['red_dot_sections'].items():
            if content and isinstance(content, str):
                # Only print the first 5 lines to keep output manageable
                preview = "\n".join(content.split("\n")[:5])
                logger.info(f"\n--- {section} ---\n{preview}...")
            elif content:
                logger.info(f"{section}: {type(content)}")
    else:
        logger.error("Failed to extract Red Dot sections!")
        
    # Also verify table extraction
    if 'tables' in data and data['tables']:
        logger.info(f"Extracted {len(data['tables'])} tables")
        for i, table in enumerate(data['tables']):
            if len(table) > 0:
                # Print table headers to help identify what each table contains
                if len(table[0]) > 0:
                    header_text = " | ".join([str(cell) for cell in table[0]])
                    logger.info(f"Table {i}: {len(table)} rows, Header: {header_text}")
                else:
                    logger.info(f"Table {i}: {len(table)} rows")
                
                # Try to identify what this table contains
                is_reagent_table = False
                is_assay_table = False
                reagent_keywords = ['component', 'reagent', 'kit', 'material', 'content']
                assay_keywords = ['assay', 'step', 'procedure', 'protocol']
                
                # Check table headers for keywords
                if len(table) > 0 and len(table[0]) > 0:
                    header_lower = " ".join([str(cell).lower() for cell in table[0]])
                    if any(keyword in header_lower for keyword in reagent_keywords):
                        is_reagent_table = True
                        logger.info(f"Table {i} appears to be a REAGENTS table")
                    elif any(keyword in header_lower for keyword in assay_keywords):
                        is_assay_table = True
                        logger.info(f"Table {i} appears to be an ASSAY PROCEDURE table")
            else:
                logger.info(f"Table {i}: empty table")
    
    return data

def test_red_dot_population():
    """Test the Red Dot template population with preserved formatting."""
    # Use the Red Dot sample file
    source_path = Path("attached_assets/RDR-LMNB2-Hu.docx")
    template_path = Path("templates_docx/enhanced_red_dot_template.docx")
    output_path = Path("improved_red_dot_output.docx")
    
    # Populate the template 
    success = populate_red_dot_template(
        source_path=source_path,
        template_path=template_path,
        output_path=output_path
    )
    
    if success:
        logger.info(f"Successfully populated Red Dot template at {output_path}")
    else:
        logger.error("Failed to populate Red Dot template!")
    
    return success

if __name__ == "__main__":
    # First test extraction
    data = test_red_dot_extraction()
    
    # Then test population
    test_red_dot_population()