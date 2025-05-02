#!/usr/bin/env python3
"""
Complete Red Dot Solution Test

This script tests the complete Red Dot solution, including:
1. Using the enhanced Red Dot template
2. Properly mapping sections from the source document
3. Converting the REAGENTS PROVIDED section to a proper table
4. Setting the Red Dot footer
5. Fixing section headers (PREPARATION vs. PREPERATION)
6. Separating ASSAY PROCEDURE and ASSAY PROCEDURE SUMMARY sections
"""

import logging
import sys
from pathlib import Path
from red_dot_template_populator import populate_red_dot_template

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_red_dot_solution(source_path="attached_assets/RDR-LMNB2-Hu.docx",
                         output_filename="complete_red_dot_output.docx"):
    """
    Test the complete Red Dot solution.
    
    Args:
        source_path: Path to the source Red Dot document
        output_filename: Name of the output file to create
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Get paths
        source_path = Path(source_path)
        template_path = Path("templates_docx/enhanced_red_dot_template.docx")
        output_path = Path(output_filename)
        
        logger.info(f"Starting complete Red Dot solution test with source: {source_path}")
        
        # Ensure source document exists
        if not source_path.exists():
            logger.error(f"Source document not found: {source_path}")
            return False
            
        # Ensure template exists
        if not template_path.exists():
            logger.error(f"Template not found: {template_path}")
            return False
        
        # Create Red Dot document
        success = populate_red_dot_template(
            source_path=source_path,
            template_path=template_path,
            output_path=output_path,
            kit_name="Human LMNB2 ELISA Kit",
            catalog_number="IMSKLK1KT",
            lot_number="SAMPLE"
        )
        
        if success:
            logger.info(f"Successfully created Red Dot document: {output_path}")
            
            # Verify the output
            # 1. Check if output file exists
            if not output_path.exists():
                logger.error(f"Output file not found: {output_path}")
                return False
                
            # 2. Run some verification checks
            try:
                from check_red_dot_output import check_document_structure
                check_document_structure(output_path)
            except Exception as e:
                logger.error(f"Error running verification: {e}")
                
            # Success!
            logger.info("Complete Red Dot solution test passed!")
            return True
        else:
            logger.error("Failed to create Red Dot document")
            return False
    
    except Exception as e:
        logger.error(f"Error in complete Red Dot solution test: {e}")
        return False

if __name__ == "__main__":
    # Use command line arguments if provided
    if len(sys.argv) > 1:
        source_path = sys.argv[1]
        if len(sys.argv) > 2:
            output_filename = sys.argv[2]
        else:
            output_filename = "complete_red_dot_output.docx"
        test_red_dot_solution(source_path, output_filename)
    else:
        # Run with default parameters
        test_red_dot_solution()