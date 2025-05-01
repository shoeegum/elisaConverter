#!/usr/bin/env python3
"""
Test Red Dot Template Population

This script tests the red_dot_template_populator by generating a populated Red Dot template
from the source ELISA kit datasheet.
"""

import logging
from pathlib import Path

from red_dot_template_populator import populate_red_dot_template
from format_document import apply_document_formatting
from add_assay_principle import add_assay_principle

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """
    Test the Red Dot template population process.
    """
    try:
        # Define paths
        source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
        template_path = Path("templates_docx/red_dot_template.docx")
        output_path = Path("output_red_dot_template.docx")
        
        # Set optional overrides
        kit_name = "Mouse KLK1 (Kallikrein 1) ELISA Kit"
        catalog_number = "RDR-KLK1-Ms"
        lot_number = "20250501"
        
        # Populate the template
        logger.info(f"Populating Red Dot template with data from {source_path}")
        success = populate_red_dot_template(
            source_path=source_path,
            template_path=template_path,
            output_path=output_path,
            kit_name=kit_name,
            catalog_number=catalog_number,
            lot_number=lot_number
        )
        
        if success:
            # Apply consistent formatting to the document
            logger.info("Applying document formatting")
            apply_document_formatting(output_path)
            
            logger.info(f"Successfully generated Red Dot template at {output_path}")
            return 0
        else:
            logger.error("Failed to populate Red Dot template")
            return 1
            
    except Exception as e:
        logger.exception(f"Error testing Red Dot template: {e}")
        return 1

if __name__ == "__main__":
    exit(main())