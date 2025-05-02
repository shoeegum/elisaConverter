#!/usr/bin/env python3
"""
Test the improved Red Dot template population with enhanced extraction.
"""

import logging
from pathlib import Path
from red_dot_template_populator import populate_red_dot_template

# Configure logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_improved_red_dot():
    """Test the improved Red Dot template population."""
    source_path = Path("attached_assets/RDR-LMNB2-Hu.docx")
    template_path = Path("templates_docx/enhanced_red_dot_template.docx")
    output_path = Path("improved_red_dot_output.docx")
    
    success = populate_red_dot_template(
        source_path,
        template_path,
        output_path,
        kit_name="Human LMNB2 ELISA Kit",
        catalog_number="IMSKLK1KT",
        lot_number="SAMPLE"
    )
    
    if success:
        logger.info(f"Successfully populated Red Dot template at {output_path}")
    else:
        logger.error("Failed to populate Red Dot template")

if __name__ == "__main__":
    test_improved_red_dot()