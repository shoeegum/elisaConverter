#!/usr/bin/env python3
"""
Run Red Dot Template Populator

This script tests the Red Dot template populator by parsing a source ELISA datasheet
and generating a Red Dot formatted document.
"""

import logging
from pathlib import Path

from red_dot_template_populator import populate_red_dot_template
from format_document import apply_document_formatting
from modify_footer import modify_footer_text

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """
    Main function to test Red Dot template population.
    """
    try:
        # Source file is the sample ELISA datasheet
        source_path = Path('attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx')
        # Template is the Red Dot template we created
        template_path = Path('templates_docx/red_dot_template.docx')
        # Output path
        output_path = Path('output_red_dot_template.docx')
        
        # Check if files exist
        if not source_path.exists():
            logger.error(f"Source file {source_path} does not exist")
            return False
            
        if not template_path.exists():
            logger.error(f"Template file {template_path} does not exist")
            return False
            
        # Define kit details
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
        
        if not success:
            logger.error("Failed to populate Red Dot template")
            return False
            
        # Apply consistent formatting
        logger.info("Applying document formatting")
        apply_document_formatting(output_path)
        
        # Modify footer text
        logger.info("Modifying footer text")
        modify_footer_text(output_path)
        
        logger.info(f"Successfully generated Red Dot template at {output_path}")
        return True
        
    except Exception as e:
        logger.exception(f"Error running Red Dot template test: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if success:
        print("Red Dot template population completed successfully!")
    else:
        print("Error: Failed to complete Red Dot template population.")