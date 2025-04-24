#!/usr/bin/env python3
"""
Update the enhanced template to include all required sections:
- ASSAY PRINCIPLE
- SAMPLE PREPARATION AND STORAGE
- SAMPLE COLLECTION NOTES
- SAMPLE DILUTION GUIDELINE
- DATA ANALYSIS

This will create a new enhanced template file that includes all the required sections.
"""

import logging
from pathlib import Path
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_heading(doc, text, level=2):
    """Create a heading with the specified text and level."""
    heading = doc.add_paragraph(text)
    heading.style = f'Heading {level}'
    
    # Set heading to all caps and blue color
    for run in heading.runs:
        run.bold = True
        run.font.color.rgb = RGBColor(0, 70, 180)  # RGB for blue
        run.text = run.text.upper()

def create_paragraph(doc, text="", style="Normal"):
    """Create a paragraph with the specified text and style."""
    paragraph = doc.add_paragraph()
    paragraph.style = style
    if text:
        paragraph.add_run(text)
    return paragraph

def update_enhanced_template():
    """
    Update the enhanced template to include all required sections.
    """
    # Create a new document
    output_path = Path('templates_docx/enhanced_template_complete.docx')
    
    # Start by copying the existing enhanced template
    doc = Document('templates_docx/enhanced_template.docx')
    
    # Find where to insert new sections
    insert_position = None
    for i, para in enumerate(doc.paragraphs):
        if "ASSAY PROTOCOL" in para.text.upper():
            insert_position = i
            break
    
    if insert_position is None:
        logger.warning("Could not find ASSAY PROTOCOL section to insert before")
        insert_position = len(doc.paragraphs) - 1  # Default to end of document
    
    # Save the current document content up to the insert point
    content_before = []
    for i in range(insert_position):
        content_before.append(doc.paragraphs[i])
    
    # Save the content after the insert point
    content_after = []
    for i in range(insert_position, len(doc.paragraphs)):
        content_after.append(doc.paragraphs[i])
    
    # Create a new document
    new_doc = Document()
    
    # Copy styles from original document
    for style in doc.styles:
        if style.name not in new_doc.styles:
            try:
                new_style = new_doc.styles.add_style(style.name, style.type)
                # Copy any other style attributes as needed
            except:
                # Style might already exist
                pass
    
    # Copy content before insert point
    for para in content_before:
        new_para = new_doc.add_paragraph()
        new_para.style = para.style
        for run in para.runs:
            new_run = new_para.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            if hasattr(run.font, 'color') and run.font.color.rgb:
                new_run.font.color.rgb = run.font.color.rgb
    
    # Add new sections
    # 1. ASSAY PRINCIPLE
    create_heading(new_doc, "ASSAY PRINCIPLE")
    create_paragraph(new_doc, "{{ assay_principle }}")
    
    # 2. SAMPLE PREPARATION AND STORAGE
    create_heading(new_doc, "SAMPLE PREPARATION AND STORAGE")
    create_paragraph(new_doc, "{{ sample_preparation_and_storage }}")
    
    # 3. SAMPLE COLLECTION NOTES
    create_heading(new_doc, "SAMPLE COLLECTION NOTES")
    create_paragraph(new_doc, "{{ sample_collection_notes }}")
    
    # 4. SAMPLE DILUTION GUIDELINE
    create_heading(new_doc, "SAMPLE DILUTION GUIDELINE")
    create_paragraph(new_doc, "{{ sample_dilution_guideline }}")
    
    # Copy content after insert point
    for para in content_after:
        new_para = new_doc.add_paragraph()
        new_para.style = para.style
        for run in para.runs:
            new_run = new_para.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            if hasattr(run.font, 'color') and run.font.color.rgb:
                new_run.font.color.rgb = run.font.color.rgb
    
    # 5. Add DATA ANALYSIS section after TYPICAL DATA section
    data_analysis_position = None
    for i, para in enumerate(new_doc.paragraphs):
        if "TYPICAL DATA" in para.text.upper():
            # Find the next heading after TYPICAL DATA
            for j in range(i+1, len(new_doc.paragraphs)):
                if new_doc.paragraphs[j].style.name.startswith('Heading'):
                    data_analysis_position = j
                    break
            if data_analysis_position is None:
                # If no next heading, put at the end
                data_analysis_position = len(new_doc.paragraphs)
            break
    
    if data_analysis_position is not None:
        # Insert DATA ANALYSIS section
        paragraphs_after_data_analysis = []
        for i in range(data_analysis_position, len(new_doc.paragraphs)):
            paragraphs_after_data_analysis.append(new_doc.paragraphs[i])
        
        # Remove paragraphs after data analysis position
        for _ in range(len(new_doc.paragraphs) - data_analysis_position):
            new_doc._element.body.remove(new_doc.paragraphs[-1]._element)
        
        # Add DATA ANALYSIS section
        create_heading(new_doc, "DATA ANALYSIS")
        create_paragraph(new_doc, "{{ data_analysis }}")
        
        # Add back the paragraphs after data analysis
        for para in paragraphs_after_data_analysis:
            new_para = new_doc.add_paragraph()
            new_para.style = para.style
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                new_run.bold = run.bold
                new_run.italic = run.italic
                if hasattr(run.font, 'color') and run.font.color.rgb:
                    new_run.font.color.rgb = run.font.color.rgb
    
    # Save the updated template
    new_doc.save(output_path)
    logger.info(f"Updated enhanced template saved to {output_path}")
    
    return output_path

if __name__ == "__main__":
    template_path = update_enhanced_template()
    logger.info(f"Updated template created at: {template_path}")
    
    # Verify that all sections are in the template
    print("\nVerify that these sections are in the template:")
    print("- ASSAY PRINCIPLE")
    print("- SAMPLE PREPARATION AND STORAGE")
    print("- SAMPLE COLLECTION NOTES")
    print("- SAMPLE DILUTION GUIDELINE") 
    print("- DATA ANALYSIS")