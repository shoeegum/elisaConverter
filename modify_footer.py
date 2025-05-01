#!/usr/bin/env python3
"""
Modify Footer Text

This script changes the footer text from 'All rights reserved' to 'Made by Sophie Gall'
and removes any copyright symbols (©) from the document footer.
"""

import logging
from pathlib import Path
import shutil
from docx import Document

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def modify_footer_text(document_path):
    """
    Modifies the footer text in the document.
    
    Changes:
    - For standard documents:
      - Replaces "All rights reserved" with "Made by Sophie Gall"
      - Removes the copyright symbol (©)
      - Adds "Made by Sophie Gall" if not present
    - For Red Dot documents:
      - Sets footer to "www.innov-research.com\nPh: 248.896.0145 | Fx: 248.896.0149"
    
    Args:
        document_path: Path to the document to modify
    """
    try:
        # Make a backup of the document
        document_path = Path(document_path)
        backup_path = document_path.with_name(f"{document_path.stem}_before_footer_change{document_path.suffix}")
        shutil.copy2(document_path, backup_path)
        logger.info(f"Created backup at {backup_path}")
        
        # Load the document
        doc = Document(document_path)
        
        # Track if we made any changes to the footer
        made_changes = False
        
        # Get a list of all sections
        sections = list(doc.sections)
        logger.info(f"Found {len(sections)} sections in the document")
        
        # Check if this is a Red Dot document based on filename
        file_name = document_path.name.upper()
        is_red_dot = "RDR" in file_name or file_name.endswith('RDR.DOCX') or "RED_DOT" in file_name

        # Also check document content for Red Dot indicators if not already identified
        if not is_red_dot:
            # Check a few paragraphs to see if it mentions Red Dot
            for i, para in enumerate(doc.paragraphs[:20]):
                if "reddotbiotech.com" in para.text.lower() or "red dot" in para.text.lower():
                    is_red_dot = True
                    logger.info(f"Detected Red Dot document from content in paragraph {i}")
                    break
        
        # Process each section's footer
        for i, section in enumerate(sections):
            # Skip if linked to previous (except the first section)
            if i > 0 and section.footer.is_linked_to_previous:
                continue
            
            logger.info(f"Processing section {i+1} footer")
            
            # Clear the existing footer
            for paragraph in list(section.footer.paragraphs):
                paragraph._element.getparent().remove(paragraph._element)
                
            # Create a new paragraph for the footer
            new_para = section.footer.add_paragraph()
            
            # Set the appropriate footer text based on document type
            if is_red_dot:
                # Use Innovative Research footer for Red Dot documents
                new_para.text = "www.innov-research.com\nPh: 248.896.0145 | Fx: 248.896.0149"
                logger.info(f"Set Red Dot footer text for document: {document_path.name}")
            else:
                # Use standard "Made by Sophie Gall" footer for non-Red Dot documents
                new_para.text = "Made by Sophie Gall"
                logger.info(f"Set standard footer text: 'Made by Sophie Gall'")
                
            made_changes = True
        
        # Save the document
        doc.save(document_path)
        logger.info(f"Successfully modified footer in: {document_path}")
        return True
        
    except Exception as e:
        logger.error(f"Error modifying footer: {e}")
        return False

if __name__ == "__main__":
    # Apply to the current output document
    modify_footer_text("output_populated_template.docx")