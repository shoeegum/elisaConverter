#!/usr/bin/env python3
"""
Run Boster Template Processing

This script provides a helper function to process a Boster ELISA kit datasheet 
and generate an Innovative Research format output.
"""

import logging
from pathlib import Path

from boster_template_populator import populate_boster_template
from fix_red_dot_document_comprehensive import fix_red_dot_document

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def run_boster_processing(
    source_path,
    output_path,
    kit_name=None,
    catalog_number=None,
    lot_number=None
):
    """
    Run the full Boster document processing workflow
    
    Args:
        source_path: Path to the Boster document
        output_path: Path to save the output
        kit_name: Optional kit name override
        catalog_number: Optional catalog number override
        lot_number: Optional lot number override
        
    Returns:
        bool: True if processing was successful, False otherwise
    """
    try:
        # Find or create template
        template_path = Path("templates_docx/boster_template.docx")
        if not template_path.exists():
            from create_boster_template import create_boster_template
            template_path = create_boster_template()
            logger.info(f"Created Boster template at: {template_path}")
        
        # Execute the Boster template population
        logger.info(f"Processing Boster document: {source_path}")
        success = populate_boster_template(
            source_path=source_path,
            template_path=template_path,
            output_path=output_path,
            kit_name=kit_name,
            catalog_number=catalog_number,
            lot_number=lot_number
        )
        
        if not success:
            logger.error("Failed to populate Boster template")
            return False
            
        # Apply comprehensive formatting fixes 
        logger.info("Applying comprehensive formatting fixes")
        success = fix_red_dot_document(output_path)
        
        if not success:
            logger.warning("Could not apply all formatting fixes, document may need manual adjustment")
        else:
            logger.info(f"Successfully processed Boster document to: {output_path}")
            
        return True
        
    except Exception as e:
        logger.exception(f"Error processing Boster document: {e}")
        return False

if __name__ == "__main__":
    # When run directly, use the default Boster test document
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    output_path = Path("boster_output.docx")
    
    # Example custom values
    kit_name = "Mouse KLK1/Kallikrein 1 ELISA Kit"
    catalog_number = "IMSKLK1KT"
    lot_number = "20250506"
    
    run_boster_processing(
        source_path=source_path,
        output_path=output_path,
        kit_name=kit_name,
        catalog_number=catalog_number,
        lot_number=lot_number
    )