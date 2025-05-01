#!/usr/bin/env python3
"""
Test Red Dot Template in Batch Processing

This script tests batch processing using the Red Dot template.
"""

import logging
import os
from pathlib import Path

from batch_processor import BatchProcessor

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """
    Test batch processing with Red Dot template.
    """
    try:
        # Create output directory for the test
        output_dir = Path("batch_test_output")
        output_dir.mkdir(exist_ok=True)
        
        # Set template path to Red Dot template
        template_path = Path("templates_docx/red_dot_template.docx")
        
        # Create batch processor with Red Dot template
        processor = BatchProcessor(template_path, output_dir)
        
        # Define file to process
        files = [Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")]
        
        # Define custom catalog number and lot number
        catalog_numbers = ["RDR-KLK1-Ms"]
        lot_numbers = ["20250501"]
        
        # Process the file
        logger.info(f"Processing files with Red Dot template: {template_path}")
        results = processor.process_batch(
            file_paths=files,
            catalog_numbers=catalog_numbers,
            lot_numbers=lot_numbers
        )
        
        # Print results
        logger.info(f"Batch processing completed with {results['successful']} successful and {results['failed']} failed")
        
        # Print details of each processed file
        for file_result in results['files']:
            if file_result['success']:
                logger.info(f"Successfully processed {os.path.basename(file_result['file'])} to {os.path.basename(file_result['output'])}")
            else:
                logger.error(f"Failed to process {os.path.basename(file_result['file'])}: {file_result['error']}")
        
        return results['successful'] > 0
        
    except Exception as e:
        logger.exception(f"Error in batch processing test: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if success:
        print("Red Dot template batch processing completed successfully!")
    else:
        print("Error: Failed to complete Red Dot template batch processing.")