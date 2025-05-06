#!/usr/bin/env python3
"""
Boster Template Populator

This module populates the Innovative Research template with data extracted from Boster ELISA kit datasheets.
It maps extracted ELISA kit data from Boster format to the Innovative Research template format.
"""

import logging
import re
from pathlib import Path
import shutil

import docxtpl
from docxtpl import DocxTemplate
import docx
from docx import Document

from elisa_parser import ELISADatasheetParser

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def extract_boster_data(source_path):
    """
    Extract data from a Boster document with special handling for Boster-specific sections.
    
    Args:
        source_path: Path to the Boster document
        
    Returns:
        Dictionary containing extracted data
    """
    try:
        parser = ELISADatasheetParser(source_path)
        data = parser.extract_data()
        
        # Parse the Boster document structure
        doc = Document(source_path)
        
        # Extract additional data not handled by the standard parser
        # This will map Boster-specific sections to our standard format
        
        # Initialize standard sections to be populated
        section_mappings = {
            'INTENDED USE': None,  # Section text or part of title
            'TEST PRINCIPLE': None,  # ASSAY PRINCIPLE 
            'CHARACTERISTICS': None,  
            'KIT COMPONENTS': None,
            'STORAGE INFORMATION': None,
            'DILUTION OF STANDARD': None,
            'SAMPLE PREPARATION': None,
            'REAGENT PREPARATION': None,  
            'ASSAY PROTOCOL': None,
            'DATA ANALYSIS': None
        }
        
        # Locate Boster-specific sections
        current_section = None
        section_content = {}
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            
            # Check if this is a section header
            if text.upper() in section_mappings:
                current_section = text.upper()
                section_content[current_section] = []
            elif current_section:
                # Add content to current section
                if text:
                    section_content[current_section].append(text)
        
        # Extract specific content for targeted sections
        if 'INTENDED USE' in section_content:
            data['intended_use'] = '\n'.join(section_content['INTENDED USE'])
        
        if 'TEST PRINCIPLE' in section_content:
            data['test_principle'] = '\n'.join(section_content['TEST PRINCIPLE'])
            
        # Map Boster sections to the Innovative Research template
        # The Boster document uses different section names than our template
        # Map the extracted content to the correct template placeholders
        if data.get('test_principle'):
            data['assay_principle'] = data['test_principle']
        
        # Handle Boster-specific tables
        # Extract special tables like Typical Data (standard curve), etc.
        
        # Cleanup content - remove text that shouldn't appear in the Innovative Research version
        # Remove phrases like "according to the picture shown below"
        for key in data:
            if isinstance(data[key], str):
                data[key] = data[key].replace("according to the picture shown below", "")
        
        # Add default disclaimer for Boster documents
        data['disclaimer'] = (
            "FOR RESEARCH USE ONLY. NOT FOR USE IN DIAGNOSTIC PROCEDURES.\n\n"
            "This kit is manufactured by Boster Biological Technology. "
            "Innovative Research is the exclusive distributor of this product. "
            "The product is warranted to perform as described in the accompanying "
            "protocol. If this product does not perform as described in our published "
            "materials, please contact Innovative Research for replacement."
        )
        
        # Map Boster section names to Innovative Research template section names
        # This helps ensure our template is populated with the right content
        section_name_mappings = {
            'background': 'background',
            'assay_principle': 'test_principle',
            'standard_curve': 'typical_data',
            'sample_collection': 'sample_preparation',
            'reagent_preparation': 'reagent_preparation',
            'assay_procedure': 'assay_protocol'
        }
        
        # Log the mappings that are being applied
        for standard_name, boster_name in section_name_mappings.items():
            if boster_name in data and standard_name != boster_name:
                logger.info(f"Mapped standard {standard_name} to {boster_name}")
                data[standard_name] = data[boster_name]
        
        return data
        
    except Exception as e:
        logger.exception(f"Error extracting data from Boster document: {e}")
        return {}

def populate_boster_template(source_path, template_path, output_path,
                           kit_name=None, catalog_number=None, lot_number=None):
    """
    Extract data from a Boster document and populate an Innovative Research template.
    
    Args:
        source_path: Path to the Boster document
        template_path: Path to the Innovative Research template
        output_path: Path to save the populated template
        kit_name: Optional kit name override
        catalog_number: Optional catalog number override
        lot_number: Optional lot number override
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Convert paths to Path objects if they're strings
        source_path = Path(source_path)
        template_path = Path(template_path)
        output_path = Path(output_path)
        
        # Extract data from Boster document
        logger.info(f"Extracting data from Boster document: {source_path}")
        data = extract_boster_data(source_path)
        
        if not data:
            logger.error("Failed to extract data from Boster document")
            return False
            
        # Override extracted values with user-provided values if available
        if kit_name:
            logger.info(f"Using custom kit name: {kit_name}")
            data['kit_name'] = kit_name
            data['document_title'] = kit_name
            
        if catalog_number:
            logger.info(f"Using custom catalog number: {catalog_number}")
            data['catalog_number'] = catalog_number
            
        if lot_number:
            logger.info(f"Using custom lot number: {lot_number}")
            data['lot_number'] = lot_number
            
        # Populate the Innovative Research template with extracted data
        logger.info(f"Populating template: {template_path}")
        try:
            # Use DocxTemplate to render the template with our data
            doc = DocxTemplate(template_path)
            doc.render(data)
            doc.save(output_path)
            logger.info(f"Saved populated template to: {output_path}")
            
            # Apply comprehensive fixes that include footer and formatting updates
            from fix_red_dot_document_comprehensive import fix_red_dot_document
            success = fix_red_dot_document(output_path)
            
            if success:
                logger.info(f"Applied comprehensive formatting fixes to: {output_path}")
            else:
                logger.warning("Could not apply all formatting fixes, document may need manual adjustment")
                
            return True
            
        except Exception as e:
            logger.exception(f"Error rendering template: {e}")
            return False
            
    except Exception as e:
        logger.exception(f"Error processing Boster document: {e}")
        return False

if __name__ == "__main__":
    # When run directly, use the default Boster test document
    source_path = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    # Use the enhanced template since it has the right structure
    template_path = Path("templates_docx/enhanced_template.docx")
    output_path = Path("boster_output.docx")
    
    # Example custom values
    kit_name = "Mouse KLK1/Kallikrein 1 ELISA Kit"
    catalog_number = "IMSKLK1KT"
    lot_number = "20250506"
    
    populate_boster_template(
        source_path=source_path,
        template_path=template_path,
        output_path=output_path,
        kit_name=kit_name,
        catalog_number=catalog_number,
        lot_number=lot_number
    )