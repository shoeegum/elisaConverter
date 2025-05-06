#!/usr/bin/env python3
"""
Create Boster-Specific Template for ELISA Kit Datasheets

This script creates a specialized template with proper styles, formatting, and placeholders
for Boster ELISA kit datasheets, converting them to the Innovative Research format
with only the specified sections.
"""

import logging
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_boster_template():
    """
    Create a specialized template for Boster ELISA kit datasheets with proper styling.
    """
    # Create a new document
    doc = Document()
    
    # Set default font for the document
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # Ensure paragraphs have proper spacing
    paragraph_format = style.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(6)
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    
    # Set document margins (1 inch on all sides)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Create Heading 2 style for section titles
    heading_2_style = doc.styles['Heading 2']
    heading_2_font = heading_2_style.font
    heading_2_font.name = 'Calibri'
    heading_2_font.size = Pt(12)
    heading_2_font.bold = True
    heading_2_font.color.rgb = RGBColor(0, 70, 180)  # blue color
    
    # First page - Title and basic info
    # Add title placeholder
    title = doc.add_paragraph()
    title.style = doc.styles['Title']
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("{{ kit_name }}")
    title_run.font.size = Pt(36)
    title_run.font.bold = True
    title_run.font.name = 'Calibri'
    
    # Add catalog/lot numbers
    catalog = doc.add_paragraph()
    catalog.alignment = WD_ALIGN_PARAGRAPH.CENTER
    catalog.add_run("Catalog No: ").bold = True
    catalog.add_run("{{ catalog_number }}").bold = False
    
    lot = doc.add_paragraph()
    lot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lot.add_run("Lot No: ").bold = True
    lot.add_run("{{ lot_number }}").bold = False
    
    # Add the INTENDED USE section on the first page
    intended_use_title = doc.add_heading("INTENDED USE", level=2)
    intended_use_para = doc.add_paragraph("{{ intended_use }}")
    
    # Add a page break to start the main content on a new page
    doc.add_page_break()
    
    # Set up the second section with different header/footer
    new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    
    # Add TEST PRINCIPLE section
    doc.add_heading("TEST PRINCIPLE", level=2)
    doc.add_paragraph("{{ test_principle }}")
    
    # Add TECHNICAL DETAILS section
    doc.add_heading("TECHNICAL DETAILS", level=2)
    
    # Add a table for technical details
    tech_table = doc.add_table(rows=4, cols=2)
    tech_table.style = 'Table Grid'
    
    # Set headers for technical details table
    tech_table.cell(0, 0).text = "Species"
    tech_table.cell(0, 1).text = "{{ species }}"
    tech_table.cell(1, 0).text = "Sensitivity"
    tech_table.cell(1, 1).text = "{{ sensitivity }}"
    tech_table.cell(2, 0).text = "Detection Range"
    tech_table.cell(2, 1).text = "{{ detection_range }}"
    tech_table.cell(3, 0).text = "Sample Type"
    tech_table.cell(3, 1).text = "{{ sample_types }}"
    
    # Add OVERVIEW section
    doc.add_heading("OVERVIEW", level=2)
    doc.add_paragraph("{{ overview }}")
    
    # Add BACKGROUND section
    doc.add_heading("BACKGROUND", level=2)
    doc.add_paragraph("{{ background }}")
    
    # Add KIT COMPONENTS section
    doc.add_heading("KIT COMPONENTS/MATERIALS PROVIDED", level=2)
    doc.add_paragraph("{{ reagents_provided }}")
    
    # Add REQUIRED MATERIALS section
    doc.add_heading("REQUIRED MATERIALS THAT ARE NOT SUPPLIED", level=2)
    doc.add_paragraph("{{ other_supplies_required }}")
    
    # Add STANDARD CURVE EXAMPLE section
    doc.add_heading("{{ kit_name }} ELISA STANDARD CURVE EXAMPLE", level=2)
    doc.add_paragraph("{{ typical_data }}")
    
    # Add INTRA/INTER-ASSAY VARIABILITY section
    doc.add_heading("INTRA/INTER-ASSAY VARIABILITY", level=2)
    doc.add_paragraph("{{ reproducibility }}")
    
    # Add REPRODUCIBILITY section
    doc.add_heading("REPRODUCIBILITY", level=2)
    doc.add_paragraph("{{ reproducibility }}")
    
    # Add PREPARATIONS BEFORE THE EXPERIMENT section
    doc.add_heading("PREPARATIONS BEFORE THE EXPERIMENT", level=2)
    doc.add_paragraph("{{ preparations_before_assay }}")
    
    # Add DILUTION OF STANDARD section
    product_placeholder = "{{ kit_name }}"
    doc.add_heading(f"DILUTION OF {product_placeholder} STANDARD", level=2)
    doc.add_paragraph("{{ dilution_of_standard }}")
    
    # Add SAMPLE PREPARATION AND STORAGE section
    doc.add_heading("SAMPLE PREPARATION AND STORAGE", level=2)
    doc.add_paragraph("{{ sample_preparation_and_storage }}")
    
    # Add SAMPLE COLLECTION NOTES section
    doc.add_heading("SAMPLE COLLECTION NOTES", level=2)
    doc.add_paragraph("{{ sample_collection_notes }}")
    
    # Add SAMPLE DILUTION GUIDELINE section
    doc.add_heading("SAMPLE DILUTION GUIDELINE", level=2)
    doc.add_paragraph("{{ sample_dilution_guideline }}")
    
    # Add ASSAY PROTOCOL section
    doc.add_heading("ASSAY PROTOCOL", level=2)
    doc.add_paragraph("{{ assay_procedure }}")
    
    # Add DATA ANALYSIS section
    doc.add_heading("DATA ANALYSIS", level=2)
    doc.add_paragraph("{{ calculation_of_results }}")
    
    # Add BACKGROUND ON PRODUCT section
    doc.add_heading(f"BACKGROUND ON {product_placeholder}", level=2)
    doc.add_paragraph("{{ background }}")
    
    # Add footer with Innovative Research
    for section in doc.sections:
        footer = section.footer
        footer_para = footer.paragraphs[0]
        if footer_para.text:
            footer_para.text = ""  # Clear any existing text
        
        # Add Innovative Research, Inc. on the right
        footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        footer_run = footer_para.add_run("Innovative Research, Inc.")
        footer_run.font.name = 'Calibri'
        footer_run.font.size = Pt(26)
        
        # Add contact info on the left in a new paragraph
        contact_para = footer.add_paragraph()
        contact_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        contact_run = contact_para.add_run("www.innov-research.com Ph: 248.896.0145 | Fx: 248.896.0149")
        contact_run.font.name = 'Calibri'
        contact_run.font.size = Pt(10)
    
    # Save the template
    template_path = Path("templates_docx/boster_template.docx")
    doc.save(template_path)
    logger.info(f"Boster-specific template created at {template_path}")
    
    return template_path

if __name__ == "__main__":
    create_boster_template()