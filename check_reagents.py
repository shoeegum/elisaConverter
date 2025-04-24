#!/usr/bin/env python3
"""
Check the reagents in the extracted data
"""

import logging
from pathlib import Path
from elisa_parser import ELISADatasheetParser

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def check_reagents(source_path='attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx'):
    """Check the reagents extracted from the source document."""
    source_path = Path(source_path)
    
    if not source_path.exists():
        logger.error(f"Source file does not exist: {source_path}")
        return
    
    # Parse the ELISA datasheet
    logger.info(f"Parsing ELISA datasheet: {source_path}")
    parser = ELISADatasheetParser(source_path)
    data = parser.extract_data()
    
    # Check reagents
    if 'reagents' in data:
        reagents_data = data['reagents']
        logger.info(f"Reagents data type: {type(reagents_data)}")
        
        if isinstance(reagents_data, dict):
            header_row = reagents_data.get('header_row', [])
            reagents = reagents_data.get('reagents', [])
            
            logger.info(f"Header row: {header_row}")
            logger.info(f"Found {len(reagents)} reagents")
            
            # Print all reagents
            print("\nReagents (from dict):")
            print("-" * 50)
            for i, reagent in enumerate(reagents):
                print(f"Reagent {i+1}:")
                for key, value in reagent.items():
                    print(f"  {key}: {value}")
                print("")
        elif isinstance(reagents_data, list):
            logger.info(f"Found {len(reagents_data)} reagents (list format)")
            
            # Print all reagents
            print("\nReagents (from list):")
            print("-" * 50)
            for i, reagent in enumerate(reagents_data):
                print(f"Reagent {i+1}:")
                if isinstance(reagent, dict):
                    for key, value in reagent.items():
                        print(f"  {key}: {value}")
                else:
                    print(f"  {reagent}")
                print("")
        else:
            logger.warning("Reagents data is not in expected format")
    else:
        logger.warning("No reagents found in extracted data")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        check_reagents(sys.argv[1])
    else:
        check_reagents()