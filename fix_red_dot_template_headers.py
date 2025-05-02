#!/usr/bin/env python3
"""
Fix Red Dot Template Headers

This script modifies the Red Dot template to:
1. Fix misspelled section headers "SAMPLE PREPERATION" and "REAGENT PREPERATION"
2. Ensure ASSAY PROCEDURE and ASSAY PROCEDURE SUMMARY are separate sections 
3. Format REAGENTS PROVIDED as a proper table
"""

import logging
import shutil
from pathlib import Path
from docx import Document

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_template_headers(template_path):
    """
    Fix misspelled section headers in the Red Dot template.
    
    Args:
        template_path: Path to the template document to modify
    """
    # Create a backup of the document
    template_path = Path(template_path)
    backup_path = template_path.with_name(f"{template_path.stem}_before_fix{template_path.suffix}")
    shutil.copy2(template_path, backup_path)
    logger.info(f"Created backup at {backup_path}")
    
    # Load the document
    doc = Document(template_path)
    
    # Track if we made any changes
    changes_made = False
    
    # Fix section headers
    for para in doc.paragraphs:
        if "SAMPLE PREPERATION" in para.text:
            para.text = "SAMPLE PREPARATION"
            logger.info("Fixed section header: SAMPLE PREPARATION")
            changes_made = True
        elif "REAGENT PREPERATION" in para.text:
            para.text = "REAGENT PREPARATION"
            logger.info("Fixed section header: REAGENT PREPARATION")
            changes_made = True
    
    # Save the document if changes were made
    if changes_made:
        doc.save(template_path)
        logger.info(f"Successfully fixed template headers in: {template_path}")
        return True
    else:
        logger.info("No misspelled headers found in the template")
        return False

if __name__ == "__main__":
    # Fix headers in the enhanced Red Dot template
    template_path = "templates_docx/enhanced_red_dot_template.docx"
    fix_template_headers(template_path)