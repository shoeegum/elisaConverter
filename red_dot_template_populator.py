#!/usr/bin/env python3
"""
Red Dot Template Populator

This module populates the Red Dot template with data extracted from source documents.
It maps extracted ELISA kit data to the Red Dot template format.
"""

import logging
import re
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional
from docxtpl import DocxTemplate

from elisa_parser import extract_elisa_data

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Mapping of source document sections to Red Dot template sections
SECTION_MAPPING = {
    "INTENDED USE": "INTENDED USE",
    "BACKGROUND": None,  # No direct mapping, we'll handle this separately
    "ASSAY PRINCIPLE": "TEST PRINCIPLE",
    "KIT COMPONENTS": "REAGENTS AND MATERIALS PROVIDED",
    "MATERIALS REQUIRED BUT NOT SUPPLIED": "MATERIALS REQUIRED BUT NOT SUPPLIED",
    "SAMPLE COLLECTION AND STORAGE": "SAMPLE COLLECTION AND STORAGE",
    "PREPARATION BEFORE ASSAY": "REAGENT PREPARATION",
    "SAMPLE PREPARATION": "SAMPLE PREPARATION",
    "ASSAY PROCEDURE": "ASSAY PROCEDURE",
    "DATA ANALYSIS": "CALCULATION OF RESULTS",
    "TYPICAL DATA": "TYPICAL DATA",
    "DETECTION RANGE": "DETECTION RANGE",
    "SENSITIVITY": "SENSITIVITY",
    "SPECIFICITY": "SPECIFICITY",
    "PRECISION": "PRECISION",
    "RECOVERY": "STABILITY",  # Map recovery to stability since no exact match
    "LINEARITY": None,  # No direct mapping
    "CALIBRATION": None,  # No direct mapping
    "ASSAY PROCEDURE SUMMARY": "ASSAY PROCEDURE SUMMARY",
    "GENERAL NOTES": "IMPORTANT NOTE",
    "PRECAUTION": "PRECAUTION",
    "DISCLAIMER": "DISCLAIMER"
}

def populate_red_dot_template(
    source_path: Path, 
    template_path: Path, 
    output_path: Path,
    kit_name: str = None,
    catalog_number: str = None,
    lot_number: str = None
) -> bool:
    """
    Populate the Red Dot template with data from the source ELISA kit datasheet.
    
    Args:
        source_path: Path to the source ELISA kit datasheet
        template_path: Path to the Red Dot template
        output_path: Path where the populated template will be saved
        kit_name: Override the kit name extracted from the source
        catalog_number: Override the catalog number extracted from the source
        lot_number: Override the lot number extracted from the source
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Extract data from source document
        logger.info(f"Extracting data from {source_path}")
        data = extract_elisa_data(source_path)
        
        # Override with provided values if any
        if kit_name:
            data['kit_name'] = kit_name
        if catalog_number:
            data['catalog_number'] = catalog_number
        if lot_number:
            data['lot_number'] = lot_number
            
        # Create context for template population
        context = {}
        
        # Basic document information
        context['kit_name'] = data.get('kit_name', '')
        context['catalog_number'] = data.get('catalog_number', '')
        context['lot_number'] = data.get('lot_number', '')
        
        # Map sections according to SECTION_MAPPING
        for src_section, tgt_section in SECTION_MAPPING.items():
            if not tgt_section:
                continue  # Skip if no target mapping
                
            # Convert target section to context variable name
            var_name = tgt_section.lower().replace(' ', '_')
            
            # Get source content
            source_content = data.get('sections', {}).get(src_section, '')
            
            # Assign content to context
            context[var_name] = source_content
            
        # Special handling for sections that need custom processing
        
        # If TEST PRINCIPLE is empty, try to use ASSAY PRINCIPLE
        if not context.get('test_principle'):
            default_principle = """This assay employs the quantitative sandwich enzyme immunoassay technique. 
A monoclonal antibody specific for the target protein has been pre-coated onto a microplate. 
Standards and samples are pipetted into the wells and any target protein present is bound by the immobilized antibody. 
After washing away any unbound substances, an enzyme-linked polyclonal antibody specific for the target protein is added to the wells. 
Following a wash to remove any unbound antibody-enzyme reagent, a substrate solution is added to the wells and color develops in proportion to the amount of target protein bound in the initial step. 
The color development is stopped and the intensity of the color is measured."""
            context['test_principle'] = data.get('sections', {}).get('ASSAY PRINCIPLE', default_principle)
            
        # Format the reagents table
        reagents = data.get('reagents', [])
        if reagents:
            # Convert reagents to a formatted string representation for the table
            reagents_text = ""
            for reagent in reagents:
                reagents_text += f"{reagent.get('name', '')}\t{reagent.get('quantity', '')}\t{reagent.get('volume', '')}\t{reagent.get('storage', '')}\n"
            context['reagents_table'] = reagents_text
        else:
            context['reagents_table'] = "No reagents found in source document."
            
        # Handle materials required but not supplied
        materials = data.get('materials_required', [])
        if materials:
            materials_text = "\\n".join([f"• {material}" for material in materials])
            context['materials_required_but_not_supplied'] = materials_text
        else:
            context['materials_required_but_not_supplied'] = "Standard laboratory materials are required."
        
        # Fill in missing sections with generic content
        for section in [s.lower().replace(' ', '_') for s in SECTION_MAPPING.values() if SECTION_MAPPING[s]]:
            if section not in context or not context[section]:
                context[section] = f"Information not available in source document."
                
        # Add storage information if missing
        if not context.get('storage_of_the_kits'):
            context['storage_of_the_kits'] = """Store at 2-8°C for unopened kit.
All reagents should be stored according to individual storage requirements noted on the product label."""
                
        # Add disclaimer if missing
        if not context.get('disclaimer'):
            context['disclaimer'] = """THE PRODUCTS ARE FOR RESEARCH USE ONLY AND NOT FOR DIAGNOSTIC OR THERAPEUTIC USE.
The information provided here is based on our best knowledge. However, no warranty, expressed or implied, is made due to the fact that many factors which may influence the performance of this product are beyond our control."""
        
        # Load template and populate
        logger.info(f"Populating template: {template_path}")
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        # Save populated template
        doc.save(output_path)
        logger.info(f"Successfully populated template: {output_path}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error populating Red Dot template: {e}")
        return False
        
if __name__ == "__main__":
    # Example usage
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    template_path = Path("templates_docx/red_dot_template.docx")
    output_path = Path("output_red_dot_template.docx")
    
    populate_red_dot_template(source_path, template_path, output_path)