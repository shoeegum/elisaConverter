#!/usr/bin/env python3
"""
Test Innovative Research document detection logic.
"""
import sys
from pathlib import Path
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

def is_red_dot_document(source_path: Path) -> bool:
    """
    Determine if a document is an Innovative Research document based on filename patterns.
    
    Args:
        source_path: Path to the document to check
        
    Returns:
        True if the document is an Innovative Research document, False otherwise
    """
    # Check filename indicators
    name_upper = source_path.name.upper()
    is_red_dot = "RDR" in name_upper or name_upper.endswith('RDR.DOCX')
    
    logger.info(f"Checking {source_path.name}")
    logger.info(f"  Filename upper: {name_upper}")
    logger.info(f"  Contains 'RDR': {'RDR' in name_upper}")
    logger.info(f"  Ends with 'RDR.DOCX': {name_upper.endswith('RDR.DOCX')}")
    logger.info(f"  Is Innovative Research: {is_red_dot}")
    
    return is_red_dot

def main():
    """
    Test Red Dot document detection on a specified file or the default file.
    """
    # Use the provided file or default to the attached assets file
    if len(sys.argv) > 1:
        source_path = Path(sys.argv[1])
    else:
        source_path = Path('attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx')
    
    if not source_path.exists():
        logger.error(f"File not found: {source_path}")
        return 1
    
    # Test if the document is a Red Dot document
    is_red_dot = is_red_dot_document(source_path)
    
    # List available files in attached_assets
    print("\nFiles in attached_assets:")
    for file_path in Path('attached_assets').glob('*.docx'):
        is_red = is_red_dot_document(file_path)
        print(f"  {file_path.name}: {'RED DOT' if is_red else 'Standard'}")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())