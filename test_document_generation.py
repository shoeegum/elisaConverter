#!/usr/bin/env python3
"""
Test document generation with emphasis on special sections
"""

import logging
import sys
from pathlib import Path

from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """Test document generation and print extracted sections"""
    
    # Define paths
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    template_path = Path("templates_docx/enhanced_template.docx")
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
    # Technical details is now a dictionary with 'text' and 'technical_table'
    if isinstance(data['technical_details'], dict):
        print(f"\nTechnical Details (Text):\n{data['technical_details'].get('text', '')[:150]}...")
        print("\nTechnical Details Table:")
        for item in data['technical_details'].get('technical_table', []):
            print(f"  - {item.get('property', 'Unknown')}: {item.get('value', 'N/A')}")
    else:
        print(f"\nTechnical Details:\n{str(data['technical_details'])[:150]}...")
    print(f"\nPreparations Before Assay:\n{data['preparations_before_assay'][:150]}...")
    
    # Process required materials
    if 'required_materials' in data:
        materials = data['required_materials']
        print(f"\nRequired Materials:\n{materials[:200]}...")

    # Check if reagents were extracted
    if 'reagents' in data and data['reagents']:
        print(f"\nReagents (first 3):")
        for i, reagent in enumerate(data['reagents'][:3] if len(data['reagents']) >= 3 else data['reagents']):
            if 'name' in reagent and 'quantity' in reagent:
                print(f"  - {reagent.get('name', 'Unknown')}: {reagent.get('quantity', 'N/A')}")
    
    # Populate template with data
    logger.info(f"Populating template: {template_path}")
    # Override background with scientifically accurate information
    data['background'] = """
    Kallikreins are a group of serine proteases with diverse physiological functions. 
    Kallikrein 1 (KLK1) is a tissue kallikrein that is primarily expressed in the kidney, pancreas, and salivary glands.
    It plays important roles in blood pressure regulation, inflammation, and tissue remodeling through the kallikrein-kinin system.
    KLK1 specifically cleaves kininogen to produce the vasoactive peptide bradykinin, which acts through bradykinin receptors to mediate various biological effects.
    Studies have implicated KLK1 in cardiovascular homeostasis, renal function, and inflammation-related processes.
    """
    
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