#!/usr/bin/env python3
"""
Test document generation with emphasis on special sections
"""

import logging
import sys
from pathlib import Path

from elisa_parser import ELISADatasheetParser
from template_populator import TemplatePopulator

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """Test document generation and print extracted sections"""
    
    # Define paths
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    template_path = Path("attached_assets/boster_template_ready.docx")
    output_path = Path("outputs/test_output.docx")
    
    if not source_path.exists():
        logger.error(f"Source file does not exist: {source_path}")
        return 1
    
    if not template_path.exists():
        logger.error(f"Template file does not exist: {template_path}")
        return 1
    
    # Extract data from source document
    logger.info(f"Parsing ELISA datasheet: {source_path}")
    parser = ELISADatasheetParser(source_path)
    data = parser.extract_data()
    
    # Print extracted sections of interest
    print("\n===== EXTRACTED DATA SECTIONS =====")
    print(f"Catalog Number: {data['catalog_number']}")
    print(f"\nIntended Use:\n{data['intended_use']}")
    print(f"\nBackground:\n{data['background'][:150]}...")
    print(f"\nAssay Principle:\n{data['assay_principle'][:150]}...")
    
    print(f"\nOverview:\n{data['overview'][:150]}...")
    print(f"\nTechnical Details:\n{data['technical_details'][:150]}...")
    print(f"\nPreparations Before Assay:\n{data['preparations_before_assay'][:150]}...")
    
    # Process required materials
    if 'required_materials' in data:
        materials = data['required_materials']
        print(f"\nRequired Materials:\n{materials[:200]}...")

    # Check if reagents were extracted
    if 'reagents' in data and data['reagents']:
        print(f"\nReagents (first 3):")
        for i, reagent in enumerate(data['reagents'][:3]):
            print(f"  - {reagent.get('name', 'Unknown')}: {reagent.get('quantity', 'N/A')}")
    
    # Populate template with data
    logger.info(f"Populating template: {template_path}")
    populator = TemplatePopulator(template_path)
    populator.populate(
        data, 
        output_path, 
        kit_name="Mouse KLK1 ELISA Kit",
        catalog_number="IMSKLK1KT",  
        lot_number="20250424"
    )
    
    logger.info(f"Successfully generated populated template at: {output_path}")
    print(f"\nGenerated output file: {output_path}")
    return 0

if __name__ == "__main__":
    sys.exit(main())