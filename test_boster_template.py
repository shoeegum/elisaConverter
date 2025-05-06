#!/usr/bin/env python3
"""
Test Boster Template Processing

This script tests the processing of a Boster ELISA kit datasheet into Innovative Research format.
"""

import logging
import sys
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_boster_template():
    """Test the Boster template processing"""
    from run_boster_template import run_boster_processing
    
    # Default source document - Boster Mouse KLK1 datasheet
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    output_path = Path("boster_to_innovative_output.docx")
    
    # Sample values for testing
    kit_name = "Mouse KLK1/Kallikrein 1 ELISA Kit"
    catalog_number = "IMSKLK1KT"
    lot_number = "20250506"
    
    # Run the Boster processing with our test values
    logger.info(f"Processing Boster document: {source_path}")
    success = run_boster_processing(
        source_path=source_path,
        output_path=output_path,
        kit_name=kit_name,
        catalog_number=catalog_number,
        lot_number=lot_number
    )
    
    if success:
        logger.info(f"Successfully processed Boster document to Innovative Research format at: {output_path}")
        # Try to open the document
        try:
            import subprocess
            subprocess.run(['xdg-open', str(output_path)], check=False)
            logger.info("Attempting to open the output document...")
        except Exception as e:
            logger.warning(f"Could not open document automatically: {e}")
    else:
        logger.error("Failed to process Boster document")
        
    return success

if __name__ == "__main__":
    test_boster_template()