#!/usr/bin/env python3
"""
Updated Template Populator for ELISA Kit Datasheets

This module extends the EnhancedTemplatePopulator to add:
1. Improved Sample Preparation and Storage section with a proper table
2. Shortened Sample Dilution Guideline section
"""

import logging
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
import re

from docx import Document
import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor
from docxtpl import DocxTemplate

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def update_template_populator(
    input_document: Path,
    template_path: Path,
    output_path: Path,
    kit_name: Optional[str] = None,
    catalog_number: Optional[str] = None,
    lot_number: Optional[str] = None
) -> None:
    """
    Process ELISA datasheet by extracting data, populating template, and fixing sample sections.
    
    Args:
        input_document: Path to the input ELISA datasheet
        template_path: Path to the enhanced template
        output_path: Path where the output will be saved
        kit_name: Optional kit name provided by user
        catalog_number: Optional catalog number provided by user
        lot_number: Optional lot number provided by user
    """
    # Import here to avoid circular imports
    from elisa_parser import ElisaParser
    from template_populator_enhanced import EnhancedTemplatePopulator
    
    # Create parser and template populator instances
    parser = ElisaParser()
    populator = EnhancedTemplatePopulator(template_path)
    
    try:
        # Parse the ELISA datasheet
        extracted_data = parser.extract(input_document)
        
        # Populate the template with extracted data
        populator.populate(
            extracted_data, 
            output_path, 
            kit_name, 
            catalog_number, 
            lot_number
        )
        
        # Fix the Sample Preparation and Sample Dilution sections
        fix_sample_sections(output_path)
        
        logger.info(f"Successfully processed document with updated sample sections: {output_path}")
        
    except Exception as e:
        logger.error(f"Error processing document: {e}")
        raise

def fix_sample_sections(document_path: Path) -> None:
    """
    Fix the Sample Preparation and Sample Dilution sections in the document.
    
    Args:
        document_path: Path to the document to fix
    """
    try:
        # Load the document
        doc = Document(document_path)
        
        # Find the Sample Preparation and Sample Dilution sections
        sample_prep_idx = None
        sample_dilution_idx = None
        assay_procedure_idx = None
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if "SAMPLE PREPARATION AND STORAGE" in text:
                sample_prep_idx = i
                logger.info(f"Found SAMPLE PREPARATION AND STORAGE at paragraph {i}")
            elif "SAMPLE DILUTION GUIDELINE" in text:
                sample_dilution_idx = i
                logger.info(f"Found SAMPLE DILUTION GUIDELINE at paragraph {i}")
            elif "ASSAY PROCEDURE" in text or "ASSAY PROTOCOL" in text:
                assay_procedure_idx = i
                logger.info(f"Found ASSAY PROCEDURE at paragraph {i}")
        
        # Create a new document with the fixed sections
        new_doc = Document()
        
        # First, copy all paragraphs up to and including SAMPLE PREPARATION AND STORAGE
        if sample_prep_idx is not None:
            logger.info("Restructuring SAMPLE PREPARATION AND STORAGE section")
            for i in range(sample_prep_idx + 1):
                para = doc.paragraphs[i]
                new_para = new_doc.add_paragraph(para.text)
                new_para.style = para.style
            
            # Add sample preparation content
            new_doc.add_paragraph("These sample collection instructions and storage conditions are intended as a general guideline. Sample stability has not been evaluated.")
            new_doc.add_paragraph("")
            
            # Add SAMPLE COLLECTION NOTES
            sample_notes_para = new_doc.add_paragraph("SAMPLE COLLECTION NOTES")
            sample_notes_para.style = 'Heading 3'
            
            # Add collection notes content
            new_doc.add_paragraph("Innovative Research recommends that samples are used immediately upon preparation.")
            new_doc.add_paragraph("Avoid repeated freeze-thaw cycles for all samples.")
            new_doc.add_paragraph("Samples should be brought to room temperature (18-25°C) before performing the assay.")
            new_doc.add_paragraph("")
            
            # Add a table for sample types
            table = new_doc.add_table(rows=5, cols=2)
            table.style = 'Table Grid'
            
            # Set the table header
            table.cell(0, 0).text = "Sample Type"
            table.cell(0, 1).text = "Collection and Handling"
            
            # Set the table content
            table.cell(1, 0).text = "Cell Culture Supernatant"
            table.cell(1, 1).text = "Centrifuge at 1000 × g for 10 minutes to remove insoluble particulates. Collect supernatant."
            
            table.cell(2, 0).text = "Serum"
            table.cell(2, 1).text = "Use a serum separator tube (SST). Allow samples to clot for 30 minutes before centrifugation for 15 minutes at approximately 1000 × g. Remove serum and assay immediately or store samples at -20°C."
            
            table.cell(3, 0).text = "Plasma"
            table.cell(3, 1).text = "Collect plasma using EDTA or heparin as an anticoagulant. Centrifuge samples for 15 minutes at 1000 × g within 30 minutes of collection. Store samples at -20°C."
            
            table.cell(4, 0).text = "Cell Lysates"
            table.cell(4, 1).text = "Collect cells and rinse with ice-cold PBS. Homogenize at 1×10^7/ml in PBS with a protease inhibitor cocktail. Freeze/thaw 3 times. Centrifuge at 10,000×g for 10 min at 4°C. Aliquot the supernatant for testing and store at -80°C."
        
        # Add Sample Dilution Guideline section
        if sample_dilution_idx is not None:
            logger.info("Restructuring SAMPLE DILUTION GUIDELINE section")
            
            dilution_para = new_doc.add_paragraph("SAMPLE DILUTION GUIDELINE")
            dilution_para.style = 'Heading 2'
            
            # Add dilution guideline content
            new_doc.add_paragraph("To inspect the validity of experimental operation and the appropriateness of sample dilution proportion, it is recommended to test all plates with the provided samples. Dilute the sample so the expected concentration falls near the middle of the standard curve range.")
        
        # Add all content from the ASSAY PROCEDURE section to the end
        if assay_procedure_idx and assay_procedure_idx < len(doc.paragraphs):
            for i in range(assay_procedure_idx, len(doc.paragraphs)):
                para = doc.paragraphs[i]
                new_para = new_doc.add_paragraph(para.text)
                new_para.style = para.style
        
        # Save the document with the fixed sections
        new_doc.save(document_path)
        logger.info(f"Fixed sample sections and saved to {document_path}")
        
    except Exception as e:
        logger.error(f"Error fixing sample sections: {e}")
        # Don't raise, continue as best we can

if __name__ == "__main__":
    # Example usage
    input_doc = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    template = Path("templates_docx/enhanced_template.docx")
    output = Path("output_updated_template.docx")
    
    update_template_populator(input_doc, template, output)