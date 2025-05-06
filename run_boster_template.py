#!/usr/bin/env python3
"""
Run Boster Template Processing

This script creates the Boster template (if it doesn't exist) and processes 
a Boster ELISA kit datasheet into the Innovative Research format.
"""

import logging
import sys
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def run_boster_processing(source_path=None, output_path=None, 
                         kit_name=None, catalog_number=None, lot_number=None):
    """
    Process a Boster document into the Innovative Research format.
    
    Args:
        source_path: Path to the source Boster document
        output_path: Path to save the output document
        kit_name: Optional kit name override
        catalog_number: Optional catalog number override
        lot_number: Optional lot number override
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Default source document if none provided
        if source_path is None:
            source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
        else:
            source_path = Path(source_path)
        
        # Default output path if none provided
        if output_path is None:
            output_path = Path("boster_output.docx")
        else:
            output_path = Path(output_path)
        
        # Create template directory if it doesn't exist
        template_dir = Path("templates_docx")
        template_dir.mkdir(exist_ok=True)
        
        # Check if Boster template exists, create it if not
        template_path = template_dir / "boster_template.docx"
        if not template_path.exists():
            logger.info("Boster template not found, creating it...")
            from create_boster_template import create_boster_template
            template_path = create_boster_template()
        else:
            logger.info(f"Using existing Boster template at {template_path}")
        
        # Process the Boster document
        logger.info(f"Processing Boster document: {source_path}")
        from boster_template_populator import populate_boster_template
        result = populate_boster_template(
            source_path=source_path,
            template_path=template_path,
            output_path=output_path,
            kit_name=kit_name,
            catalog_number=catalog_number,
            lot_number=lot_number
        )
        
        if result:
            logger.info(f"Successfully processed Boster document to: {output_path}")
            
            # Apply the company name replacements
            from fix_red_dot_document_comprehensive import fix_red_dot_document
            if fix_red_dot_document(output_path):
                logger.info(f"Applied comprehensive formatting fixes to: {output_path}")
            else:
                logger.warning(f"Could not apply all formatting fixes to: {output_path}")
            
            return True
        else:
            logger.error("Failed to process Boster document")
            return False
            
    except Exception as e:
        logger.exception(f"Error processing Boster document: {e}")
        return False

if __name__ == "__main__":
    # Get the source path from command line or use default
    if len(sys.argv) > 1:
        source_path = sys.argv[1]
    else:
        source_path = "attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx"
    
    # Get the output path or use default
    if len(sys.argv) > 2:
        output_path = sys.argv[2]
    else:
        output_path = "boster_output.docx"
    
    # Run the Boster processing
    if run_boster_processing(source_path, output_path):
        logger.info("Boster document processing completed successfully")
    else:
        logger.error("Boster document processing failed")