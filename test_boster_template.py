#!/usr/bin/env python3
"""
Test Boster Template Processing

This script tests the processing of a Boster ELISA kit datasheet into Innovative Research format.
"""

import logging
import os
import subprocess
from pathlib import Path

from run_boster_template import run_boster_processing

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_boster_template():
    """Test the Boster template processing"""
    # Boster test document
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    output_path = Path("boster_to_innovative_output.docx")
    
    # Example custom values
    kit_name = "Mouse KLK1/Kallikrein 1 ELISA Kit"
    catalog_number = "IMSKLK1KT"
    lot_number = "20250506"  # Today's date for example
    
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
        logger.info("Attempting to open the output document...")
        try:
            # On Windows and macOS
            if os.name == 'nt':  # Windows
                os.startfile(output_path)
            elif os.name == 'posix':  # macOS and Linux
                if os.uname().sysname == 'Darwin':  # macOS
                    subprocess.run(['open', output_path])
                else:  # Linux
                    subprocess.run(['xdg-open', output_path])
        except Exception as e:
            logger.warning(f"Could not automatically open document: {e}")
            logger.info(f"Please open the document manually at: {output_path}")
    else:
        logger.error("Failed to process Boster document")
        
if __name__ == "__main__":
    test_boster_template()