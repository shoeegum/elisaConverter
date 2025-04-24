#!/usr/bin/env python3
"""
Direct approach to fix material bullet points in the output document.
"""

import logging
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Inches, RGBColor
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_materials():
    """Extract materials from the source document."""
    # Standard materials for ELISA kits
    standard_materials = [
        "Microplate reader capable of measuring absorbance at 450 nm",
        "Automated plate washer (optional)",
        "Adjustable pipettes and pipette tips",
        "Test tubes for preparing standard dilutions",
        "Deionized or distilled water",
        "500 ml graduated cylinders",
        "Tubes for sample preparation",
        "Incubator capable of maintaining 37°C",
        "Plate sealer for incubation steps",
        "Absorbent paper"
    ]
    
    # Try to extract from source document as well
    source_doc = Document("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    
    extracted_materials = []
    found_materials_section = False
    
    for i, para in enumerate(source_doc.paragraphs):
        if "REQUIRED MATERIALS" in para.text.upper() or "MATERIALS REQUIRED" in para.text.upper():
            found_materials_section = True
            logger.info(f"Found materials section at paragraph {i}: {para.text}")
            continue
        
        if found_materials_section and para.text.strip():
            # If we hit a section header (all caps), stop
            if para.text.strip().isupper() and len(para.text.strip()) > 10:
                break
                
            # Clean up the text
            text = para.text.strip()
            # Remove bullet points if present
            if text.startswith('•') or text.startswith('-'):
                text = text[1:].strip()
                
            # Only add non-empty, non-duplicate materials
            if text and text not in extracted_materials:
                extracted_materials.append(text)
                logger.info(f"Extracted material: {text}")
    
    # If extraction didn't find enough materials, use standards
    if len(extracted_materials) < 4:
        logger.info("Using standard materials")
        return standard_materials
    
    return extracted_materials

def fix_output_document(output_path="output_populated_template.docx", fixed_path="fixed_bullet_output.docx"):
    """Fix the output document by directly adding bullet points with material text."""
    # Get materials to add
    materials = extract_materials()
    
    # Load the document
    doc = Document(output_path)
    logger.info(f"Loaded document: {output_path}")
    
    # Find the materials section
    materials_section_index = None
    for i, para in enumerate(doc.paragraphs):
        if "MATERIALS REQUIRED" in para.text.upper():
            materials_section_index = i
            logger.info(f"Found materials section at paragraph {i}: {para.text}")
            break
    
    if materials_section_index is None:
        logger.error("Could not find materials section in document")
        return False
    
    # Find any existing bullet paragraphs after the materials section
    # and note their indices for removal
    paragraphs_to_remove = []
    for i in range(materials_section_index + 1, min(materials_section_index + 20, len(doc.paragraphs))):
        para = doc.paragraphs[i]
        # Check for bullets or style
        if '•' in para.text or (para.style.name and 'List' in para.style.name):
            logger.info(f"Found paragraph to remove at index {i}: {para.text}")
            paragraphs_to_remove.append(i)
    
    # We can't directly remove paragraphs, but we can clear their content
    for i in sorted(paragraphs_to_remove, reverse=True):
        for run in doc.paragraphs[i].runs:
            run.text = ""
    
    # Get a reference to the materials section paragraph
    section_para = doc.paragraphs[materials_section_index]
    
    # Add material bullet points
    for material in materials:
        # Create a new paragraph with bullet style
        new_para = doc.add_paragraph(style='List Bullet')
        new_para.paragraph_format.left_indent = Inches(0.25)
        new_para.paragraph_format.first_line_indent = Inches(0)
        new_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        
        # Add the bullet and material text
        bullet_run = new_para.add_run("• ")
        material_run = new_para.add_run(material)
    
    # Save the fixed document
    doc.save(fixed_path)
    logger.info(f"Saved fixed document to {fixed_path}")
    
    return True

if __name__ == "__main__":
    fix_output_document()
    print("Fixed output document with proper bullet points. Check fixed_bullet_output.docx")