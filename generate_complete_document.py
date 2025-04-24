#!/usr/bin/env python3
"""
Generate a complete document with all the required sections.
- ASSAY PRINCIPLE
- SAMPLE PREPARATION AND STORAGE
- SAMPLE COLLECTION NOTES
- SAMPLE DILUTION GUIDELINE
- DATA ANALYSIS
"""

import logging
from pathlib import Path
from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def generate_complete_document(
    source_path='attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx',
    template_path='templates_docx/enhanced_template.docx',
    output_path='output_complete.docx'
):
    """Generate a complete document with all required sections."""
    logger.info(f"Parsing ELISA datasheet: {source_path}")
    
    # Parse the source document
    parser = ELISADatasheetParser(source_path)
    data = parser.extract_data()
    
    # Make sure all required sections are present and show extracted data
    logger.info("Assay principle: " + (data.get('assay_principle', 'Not found')[:50] + "..." if data.get('assay_principle') else "Not found"))
    logger.info("Sample preparation: " + (data.get('sample_preparation_and_storage', 'Not found')[:50] + "..." if data.get('sample_preparation_and_storage') else "Not found"))
    logger.info("Sample collection: " + (data.get('sample_collection_notes', 'Not found')[:50] + "..." if data.get('sample_collection_notes') else "Not found"))
    logger.info("Sample dilution: " + (data.get('sample_dilution_guideline', 'Not found')[:50] + "..." if data.get('sample_dilution_guideline') else "Not found"))
    logger.info("Data analysis: " + (data.get('data_analysis', 'Not found')[:50] + "..." if data.get('data_analysis') else "Not found"))
    
    # Populate the template with the extracted data
    logger.info(f"Populating template: {template_path}")
    populator = TemplatePopulator(template_path)
    
    # Generate a document with Mouse KLK1/Kallikrein 1 ELISA Kit as the kit name
    populator.populate(
        data, 
        output_path, 
        kit_name="Mouse KLK1/Kallikrein 1 ELISA Kit",
        catalog_number="IMSKLK1KT",
        lot_number="20250424"
    )
    
    logger.info(f"Generated complete document: {output_path}")
    
    return output_path

if __name__ == "__main__":
    output_path = generate_complete_document()
    logger.info(f"Complete document generated at: {output_path}")
    
    # Verify that all sections are in the output
    print("\nVerify that these sections are in the output document:")
    print("- ASSAY PRINCIPLE")
    print("- SAMPLE PREPARATION AND STORAGE")
    print("- SAMPLE COLLECTION NOTES")
    print("- SAMPLE DILUTION GUIDELINE") 
    print("- DATA ANALYSIS")