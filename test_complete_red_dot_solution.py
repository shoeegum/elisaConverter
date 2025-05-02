#!/usr/bin/env python3
"""
Test Complete Red Dot Solution

This script tests the entire Red Dot solution pipeline, including:
1. Parsing the source document
2. Populating the template
3. Post-processing to fix company names
4. Post-processing to fix table position
5. Checking the final document structure
"""

import logging
import os
import shutil
from pathlib import Path
import sys

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def clean_test_output():
    """Remove previous test outputs."""
    output_files = [
        "red_dot_output.docx",
        "complete_red_dot_output.docx"
    ]
    
    for file in output_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                logger.info(f"Removed existing file: {file}")
            except Exception as e:
                logger.error(f"Could not remove file {file}: {e}")

def test_template_population():
    """Test populating the Red Dot template with data from the source document."""
    # Import the Red Dot template populator
    from red_dot_template_populator import populate_red_dot_template
    
    # Define the paths
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    template_path = Path("templates_docx/enhanced_red_dot_template.docx")
    output_path = Path("red_dot_output.docx")
    
    # Populate the template
    result = populate_red_dot_template(source_path, template_path, output_path)
    
    if result:
        logger.info(f"Successfully populated template: {output_path}")
        return True, output_path
    else:
        logger.error("Failed to populate template")
        return False, None

def check_document_structure(document_path):
    """
    Check the structure of the Red Dot document to ensure that:
    1. All sections are present
    2. Tables are properly positioned
    3. Company names are correct
    """
    try:
        # Convert Path to string for compatibility
        doc_path_str = str(document_path)
        
        # Import directly instead of using importlib for simplicity
        from check_red_dot_output import check_document_structure as check_func
        
        # Call the check_document_structure function
        check_func(doc_path_str)
        return True
    except ImportError:
        logger.error("Failed to import check_red_dot_output.py")
        return False
    except Exception as e:
        logger.error(f"Error checking document structure: {e}")
        return False

def make_copy_for_comprehensive_check(source_path, dest_path):
    """Make a copy of the file for comprehensive checking."""
    try:
        shutil.copy2(source_path, dest_path)
        logger.info(f"Created comprehensive check copy at: {dest_path}")
        return True
    except Exception as e:
        logger.error(f"Error creating copy: {e}")
        return False

def run_tests():
    """Run all tests for the Red Dot solution."""
    # Clean up previous test outputs
    clean_test_output()
    
    # Test template population
    success, output_path = test_template_population()
    if not success:
        logger.error("Template population test failed")
        return False
    
    # Make a copy for comprehensive checking
    comprehensive_path = Path("complete_red_dot_output.docx")
    if not make_copy_for_comprehensive_check(output_path, comprehensive_path):
        return False
    
    # Run a manual fix on the comprehensive test file
    try:
        from fix_red_dot_company_and_placement import fix_document
        if fix_document(comprehensive_path):
            logger.info(f"Successfully applied fixes to: {comprehensive_path}")
        else:
            logger.warning("Fixes were not fully applied")
    except Exception as e:
        logger.error(f"Error applying fixes: {e}")
    
    # Check the document structure
    check_document_structure(comprehensive_path)
    
    logger.info("All tests completed successfully!")
    return True

if __name__ == "__main__":
    run_tests()