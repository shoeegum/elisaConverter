#!/usr/bin/env python3
"""
Create Innovative Research Template from Sample

This script creates an Innovative Research template from the provided sample document.
"""

import logging
from pathlib import Path
import shutil

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_red_dot_template():
    """
    Create an Innovative Research template from the sample document.
    """
    try:
        # Source is the sample document from attached assets
        source_path = Path('attached_assets/RDR-LMNB2-Hu.docx')
        # Destination is in the templates_docx folder
        dest_path = Path('templates_docx/red_dot_template.docx')
        
        if not source_path.exists():
            logger.error(f"Source file {source_path} does not exist")
            return False
            
        # Create templates_docx folder if it doesn't exist
        dest_path.parent.mkdir(exist_ok=True)
        
        # Copy the file
        shutil.copy2(source_path, dest_path)
        logger.info(f"Successfully created Innovative Research template at {dest_path}")
        
        return True
    except Exception as e:
        logger.exception(f"Error creating Innovative Research template: {e}")
        return False

if __name__ == "__main__":
    create_red_dot_template()