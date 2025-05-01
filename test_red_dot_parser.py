#!/usr/bin/env python3
"""
Test the Red Dot template population functionality.
"""

import logging
from pathlib import Path
from red_dot_template_populator import extract_red_dot_data, populate_red_dot_template

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    # Test Red Dot document extraction
    source_path = Path("attached_assets/RDR-LMNB2-Hu.docx")
    template_path = Path("templates_docx/red_dot_template.docx")
    output_path = Path("output_red_dot_test.docx")
    
    logger.info(f"Testing Red Dot extraction with: {source_path}")
    data = extract_red_dot_data(source_path)
    
    # Log the extracted data structure
    logger.info(f"Extracted kit name: {data.get('kit_name', 'Not found')}")
    logger.info(f"Extracted catalog number: {data.get('catalog_number', 'Not found')}")
    
    # Log Red Dot specific sections if available
    if 'red_dot_sections' in data:
        logger.info("Found Red Dot specific sections:")
        for section, content in data['red_dot_sections'].items():
            if content:
                preview = content[:50] + '...' if len(content) > 50 else content
                logger.info(f"  - {section}: {preview}")
            else:
                logger.info(f"  - {section}: Empty")
    else:
        logger.info("No Red Dot specific sections found")
    
    # Test template population
    logger.info(f"Populating Red Dot template: {template_path}")
    success = populate_red_dot_template(
        source_path=source_path,
        template_path=template_path,
        output_path=output_path
    )
    
    if success:
        logger.info(f"Successfully populated template: {output_path}")
    else:
        logger.error("Failed to populate template")

if __name__ == "__main__":
    main()