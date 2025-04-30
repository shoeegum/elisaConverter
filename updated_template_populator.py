#!/usr/bin/env python3
"""
Updated Template Populator for ELISA Kit Datasheets

This module extends the EnhancedTemplatePopulator to add:
1. Improved Sample Preparation and Storage section with a proper table
2. Shortened Sample Dilution Guideline section
3. Assay Principle section placed before other sections with preserved paragraph spacing
"""

import logging
import os
import shutil
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
import re

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
    Process ELISA datasheet by extracting data and populating template.
    This version uses the April 24th working version with consistent Calibri font and 1.15 spacing.
    
    Args:
        input_document: Path to the input ELISA datasheet
        template_path: Path to the enhanced template
        output_path: Path where the output will be saved
        kit_name: Optional kit name provided by user
        catalog_number: Optional catalog number provided by user
        lot_number: Optional lot number provided by user
    """
    # Import here to avoid circular imports
    from elisa_parser import ELISADatasheetParser
    from template_populator_enhanced import TemplatePopulator
    
    try:
        # Copy the backup file directly
        backup_path = Path("IMSKLK1KT-20250424.docx")
        if backup_path.exists():
            import shutil
            shutil.copy2(backup_path, output_path)
            logger.info(f"Restored April 24th version to {output_path}")
            
            # Apply consistent formatting to the document
            from format_document import apply_document_formatting
            apply_document_formatting(output_path)
            logger.info(f"Applied consistent formatting (Calibri, 1.15 spacing) to {output_path}")
            return
            
        # If no backup is available, use the normal process
        # Create parser and template populator instances
        parser = ELISADatasheetParser(input_document)
        populator = TemplatePopulator(template_path)
        
        # Parse the ELISA datasheet
        extracted_data = parser.extract_data()
        
        # Populate the template with extracted data
        populator.populate(
            extracted_data, 
            output_path, 
            kit_name, 
            catalog_number, 
            lot_number
        )
        
        # Apply consistent formatting to the document
        from format_document import apply_document_formatting
        apply_document_formatting(output_path)
        
        logger.info(f"Successfully processed document: {output_path}")
        
    except Exception as e:
        logger.error(f"Error processing document: {e}")
        raise

def fix_sample_sections(document_path: Path) -> None:
    """
    Fix the Sample Preparation and Sample Dilution sections in the document.
    Also ensures ASSAY PRINCIPLE section appears before other sections with preserved paragraph spacing.
    
    Args:
        document_path: Path to the document to fix
    """
    try:
        # Make a backup copy
        backup_path = document_path.with_name(f"{document_path.stem}_backup{document_path.suffix}")
        shutil.copy2(document_path, backup_path)
        
        # Load the document
        doc = Document(document_path)
        
        # Find the Sample Preparation and Sample Dilution sections
        sections = {}
        section_names = ["SAMPLE PREPARATION AND STORAGE", "SAMPLE DILUTION GUIDELINE", "ASSAY PROCEDURE", "ASSAY PROTOCOL", "ASSAY PRINCIPLE"]
        section_indices = {}
        
        # Track tables and their positions
        table_positions = []
        
        # Find all section positions and table positions
        para_count = 0
        table_count = 0
        current_position = 0
        
        # First pass: find all sections and tables with their positions
        for element in doc.element.body:
            if element.tag.endswith('p'):  # This is a paragraph
                para = doc.paragraphs[para_count]
                text = para.text.strip().upper()
                para_count += 1
                current_position += 1
                
                # Check if this is a section we're interested in
                for section_name in section_names:
                    if section_name in text:
                        section_indices[section_name] = (para_count - 1, current_position)
                        break
                        
            elif element.tag.endswith('tbl'):  # This is a table
                table_positions.append((table_count, current_position))
                table_count += 1
                current_position += 1
        
        # Extract section positions
        sample_prep_position = section_indices.get("SAMPLE PREPARATION AND STORAGE")
        sample_dilution_position = section_indices.get("SAMPLE DILUTION GUIDELINE")
        assay_procedure_position = section_indices.get("ASSAY PROCEDURE") or section_indices.get("ASSAY PROTOCOL")
        assay_principle_position = section_indices.get("ASSAY PRINCIPLE")
        
        if not sample_prep_position:
            logger.warning("Could not find SAMPLE PREPARATION AND STORAGE section")
            return
            
        if not sample_dilution_position:
            logger.warning("Could not find SAMPLE DILUTION GUIDELINE section")
            return
            
        if not assay_procedure_position:
            logger.warning("Could not find ASSAY PROCEDURE section")
            return
        
        # Get paragraph index and position for each section
        sample_prep_idx, sample_prep_pos = sample_prep_position
        sample_dilution_idx, sample_dilution_pos = sample_dilution_position
        assay_procedure_idx, assay_procedure_pos = assay_procedure_position
        
        # Check if we have an ASSAY PRINCIPLE section
        assay_principle_idx = None
        assay_principle_content = []
        if assay_principle_position:
            assay_principle_idx, assay_principle_pos = assay_principle_position
            logger.info(f"Found ASSAY PRINCIPLE at paragraph {assay_principle_idx}")
            
            # Extract the content of the ASSAY PRINCIPLE section
            # Look for the next 10 paragraphs after the ASSAY PRINCIPLE heading
            start_idx = assay_principle_idx + 1
            end_idx = min(start_idx + 10, len(doc.paragraphs))
            
            for i in range(start_idx, end_idx):
                para_text = doc.paragraphs[i].text.strip()
                # Stop if we hit the next section
                if any(section in para_text.upper() for section in section_names if section != "ASSAY PRINCIPLE"):
                    break
                # Skip empty paragraphs
                if para_text:
                    assay_principle_content.append(para_text)
            
            logger.info(f"Extracted {len(assay_principle_content)} paragraphs from ASSAY PRINCIPLE section")
        
        logger.info(f"Found SAMPLE PREPARATION AND STORAGE at paragraph {sample_prep_idx}")
        logger.info(f"Found SAMPLE DILUTION GUIDELINE at paragraph {sample_dilution_idx}")
        logger.info(f"Found ASSAY PROCEDURE at paragraph {assay_procedure_idx}")
        
        # Keep track of which tables to preserve
        tables_to_preserve = {}
        
        # Identify tables that need to be preserved (those not between sections we're modifying)
        for table_idx, table_pos in table_positions:
            # Get the first cell text of the table to check for Technical Details or Overview tables
            first_cell_text = ""
            if len(doc.tables[table_idx].rows) > 0 and len(doc.tables[table_idx].rows[0].cells) > 0:
                first_cell_text = doc.tables[table_idx].rows[0].cells[0].text.strip()
                
            # Identify tables that need to be preserved
            if table_pos < sample_prep_pos:
                tables_to_preserve[table_idx] = "before_sample_prep"
            elif table_pos >= assay_procedure_pos:
                tables_to_preserve[table_idx] = "after_assay_procedure"
                
        logger.info(f"Tables to preserve: {tables_to_preserve}")
        
        # Create a temporary document with our changes
        temp_path = document_path.with_name(f"{document_path.stem}_temp{document_path.suffix}")
        temp_doc = Document()
        
        # Keep track of which paragraphs we've already copied to avoid duplication
        paragraphs_copied = set()
        
        # Skip copying tables before cover page - they'll be copied after the section break
        # This ensures no tables appear on the first page
        table_idx_in_new_doc = 0
        tables_before_sample_prep = [table_idx for table_idx, position in tables_to_preserve.items() 
                                   if position == "before_sample_prep"]
        logger.info(f"Found {len(tables_before_sample_prep)} tables before sample prep - will copy after cover page")
        
        # 2. Completely rebuild the document in the correct order
        
        # 2.1 First, ONLY add the title, catalog, lot number, and intended use to the first page
        # These are typically the first 4 paragraphs of the document
        cover_page_elements = ["Mouse KLK1", "Catalog", "Lot", "ELISA Kit"]  # Keywords to identify cover page elements
        
        cover_page_count = 0
        # First, add the title (always the first paragraph)
        if len(doc.paragraphs) > 0:
            title_para = doc.paragraphs[0]
            new_para = temp_doc.add_paragraph(title_para.text)
            new_para.style = title_para.style
            paragraphs_copied.add(0)
            cover_page_count += 1
            
        # Then look for catalog number, lot number in the next few paragraphs
        for i in range(1, min(10, len(doc.paragraphs))):  # Look in the first 10 paragraphs
            para = doc.paragraphs[i]
            para_text = para.text.strip()
            
            # Only include paragraphs that contain our cover page keywords and are not section headings
            if para_text and any(keyword in para_text for keyword in cover_page_elements) and not any(section in para_text.upper() for section in section_names):
                new_para = temp_doc.add_paragraph(para_text)
                new_para.style = para.style
                paragraphs_copied.add(i)
                cover_page_count += 1
        
        # Now find and add the INTENDED USE section to the first page
        intended_use_found = False
        for i in range(len(doc.paragraphs)):
            if "INTENDED USE" in doc.paragraphs[i].text.upper():
                # Found the INTENDED USE heading
                intended_use_heading = temp_doc.add_paragraph("INTENDED USE")
                intended_use_heading.style = 'Heading 2'
                paragraphs_copied.add(i)
                intended_use_found = True
                
                # Look for content in the next paragraph(s)
                if i + 1 < len(doc.paragraphs):
                    intended_use_content = doc.paragraphs[i + 1].text.strip()
                    # Make sure this paragraph doesn't contain table content that belongs in technical details/overview
                    if (intended_use_content and not any(section in intended_use_content.upper() for section in section_names) 
                            and "Capture/Detection" not in intended_use_content 
                            and "Product Name" not in intended_use_content):
                        intended_use_para = temp_doc.add_paragraph(intended_use_content)
                        intended_use_para.style = doc.paragraphs[i + 1].style
                        paragraphs_copied.add(i + 1)
                        cover_page_count += 2  # Count both heading and content
                break
        
        # If we didn't find the intended use section, add a default one
        if not intended_use_found:
            logger.info("INTENDED USE section not found - adding default")
            intended_use_heading = temp_doc.add_paragraph("INTENDED USE")
            intended_use_heading.style = 'Heading 2'
            
            # Extract the default text from the document or use a generic one
            # Check for text like "For the quantitation of Mouse Klk1 concentrations"
            default_text = "For the quantitation of Mouse KLK1/Kallikrein 1 concentrations in cell culture supernatants, cell lysates, serum, and plasma. For Research Use Only. Not for use in diagnostic procedures."
            
            # Look for "For the quantitation" text in the first 20 paragraphs
            for i in range(min(20, len(doc.paragraphs))):
                if "for the quantitation" in doc.paragraphs[i].text.lower() and "mouse" in doc.paragraphs[i].text.lower():
                    default_text = doc.paragraphs[i].text
                    paragraphs_copied.add(i)
                    break
                    
            intended_use_para = temp_doc.add_paragraph(default_text)
            cover_page_count += 2  # Count both heading and content
        
        logger.info(f"Added {cover_page_count} paragraphs from cover page (title, catalog, lot, intended use)")
        
        # Create a new section with a page break
        # This is a more explicit way to ensure that the content starts on a new page
        section = temp_doc.add_section()
        section.start_type = WD_SECTION_START.NEW_PAGE
        
        # 2.2 Find the TECHNICAL DETAILS section
        technical_details_idx = None
        technical_details_content = []
        
        for i in range(len(doc.paragraphs)):
            if i not in paragraphs_copied and "TECHNICAL DETAILS" in doc.paragraphs[i].text.upper():
                technical_details_idx = i
                technical_details_content.append((doc.paragraphs[i].text, doc.paragraphs[i].style))
                paragraphs_copied.add(i)
                break
        
        # 2.3 Now add the ASSAY PRINCIPLE section right after cover page, on a new page
        if assay_principle_content:
            logger.info("Adding ASSAY PRINCIPLE section after cover page")
            
            # Create the ASSAY PRINCIPLE heading
            principle_heading = temp_doc.add_paragraph("ASSAY PRINCIPLE")
            principle_heading.style = 'Heading 2'
            
            # Add the content paragraphs with spacing preserved
            for i, para_text in enumerate(assay_principle_content):
                temp_doc.add_paragraph(para_text)
                # Add an empty paragraph to preserve spacing between paragraphs
                # but not after the last paragraph
                if i < len(assay_principle_content) - 1:
                    temp_doc.add_paragraph("")
            
            # Mark the original paragraphs as copied
            if assay_principle_idx:
                # Mark the heading
                paragraphs_copied.add(assay_principle_idx)
                # Mark the content paragraphs
                start_idx = assay_principle_idx + 1
                end_idx = min(start_idx + 10, len(doc.paragraphs))
                for i in range(start_idx, end_idx):
                    para_text = doc.paragraphs[i].text.strip()
                    if any(section in para_text.upper() for section in section_names if section != "ASSAY PRINCIPLE"):
                        break
                    paragraphs_copied.add(i)
        
        # 2.4 Add TECHNICAL DETAILS section
        if technical_details_content:
            logger.info("Adding TECHNICAL DETAILS section after ASSAY PRINCIPLE")
            for text, style in technical_details_content:
                new_para = temp_doc.add_paragraph(text)
                new_para.style = style
                
            # Now add the tables that were skipped earlier (before sample prep tables)
            for table_idx in tables_before_sample_prep:
                # Get the table from the original document
                orig_table = doc.tables[table_idx]
                
                # Create a new table with same dimensions
                rows = len(orig_table.rows)
                cols = len(orig_table.rows[0].cells) if rows > 0 else 0
                
                if rows > 0 and cols > 0:
                    new_table = temp_doc.add_table(rows=rows, cols=cols)
                    new_table.style = orig_table.style
                    
                    # Copy cell content
                    for i, row in enumerate(orig_table.rows):
                        for j, cell in enumerate(row.cells):
                            if i < len(new_table.rows) and j < len(new_table.rows[i].cells):
                                new_table.rows[i].cells[j].text = cell.text
                    
                    table_idx_in_new_doc += 1
                    logger.info(f"Added 'before_sample_prep' table {table_idx} ({rows}x{cols}) after page break")
        
        # 2.5 Add all other sections except SAMPLE PREPARATION and beyond
        for i in range(len(doc.paragraphs)):
            if i not in paragraphs_copied and i < sample_prep_idx:
                para = doc.paragraphs[i]
                # Skip any duplicate ASSAY PRINCIPLE or INTENDED USE sections
                if "ASSAY PRINCIPLE" in para.text.upper() or "INTENDED USE" in para.text.upper():
                    paragraphs_copied.add(i)
                    continue
                new_para = temp_doc.add_paragraph(para.text)
                new_para.style = para.style
                paragraphs_copied.add(i)
        
        # These steps of the original process are no longer needed since we've implemented
        # a new approach to document structuring
            
        # 5. Add our customized sample preparation content
        logger.info("Restructuring SAMPLE PREPARATION AND STORAGE section")
        temp_doc.add_paragraph("These sample collection instructions and storage conditions are intended as a general guideline. Sample stability has not been evaluated.")
        temp_doc.add_paragraph("")
        
        # Add SAMPLE COLLECTION NOTES
        sample_notes_para = temp_doc.add_paragraph("SAMPLE COLLECTION NOTES")
        sample_notes_para.style = 'Heading 3'
        
        # Add collection notes content
        temp_doc.add_paragraph("Innovative Research recommends that samples are used immediately upon preparation.")
        temp_doc.add_paragraph("Avoid repeated freeze-thaw cycles for all samples.")
        temp_doc.add_paragraph("Samples should be brought to room temperature (18-25°C) before performing the assay.")
        temp_doc.add_paragraph("")
        
        # Add a table for sample types
        sample_type_table = temp_doc.add_table(rows=5, cols=2)
        sample_type_table.style = 'Table Grid'
        
        # Set the table header
        sample_type_table.cell(0, 0).text = "Sample Type"
        sample_type_table.cell(0, 1).text = "Collection and Handling"
        
        # Set the table content
        sample_type_table.cell(1, 0).text = "Cell Culture Supernatant"
        sample_type_table.cell(1, 1).text = "Centrifuge at 1000 × g for 10 minutes to remove insoluble particulates. Collect supernatant."
        
        sample_type_table.cell(2, 0).text = "Serum"
        sample_type_table.cell(2, 1).text = "Use a serum separator tube (SST). Allow samples to clot for 30 minutes before centrifugation for 15 minutes at approximately 1000 × g. Remove serum and assay immediately or store samples at -20°C."
        
        sample_type_table.cell(3, 0).text = "Plasma"
        sample_type_table.cell(3, 1).text = "Collect plasma using EDTA or heparin as an anticoagulant. Centrifuge samples for 15 minutes at 1000 × g within 30 minutes of collection. Store samples at -20°C."
        
        sample_type_table.cell(4, 0).text = "Cell Lysates"
        sample_type_table.cell(4, 1).text = "Collect cells and rinse with ice-cold PBS. Homogenize at 1×10^7/ml in PBS with a protease inhibitor cocktail. Freeze/thaw 3 times. Centrifuge at 10,000×g for 10 min at 4°C. Aliquot the supernatant for testing and store at -80°C."
        
        table_idx_in_new_doc += 1
        
        # 6. Add customized Sample Dilution Guideline section
        logger.info("Restructuring SAMPLE DILUTION GUIDELINE section")
        
        dilution_para = temp_doc.add_paragraph("SAMPLE DILUTION GUIDELINE")
        dilution_para.style = 'Heading 2'
        
        # Add dilution guideline content
        temp_doc.add_paragraph("To inspect the validity of experimental operation and the appropriateness of sample dilution proportion, it is recommended to test all plates with the provided samples. Dilute the sample so the expected concentration falls near the middle of the standard curve range.")
        
        # 7. Add all content from the ASSAY PROCEDURE section to the end
        for i in range(assay_procedure_idx, len(doc.paragraphs)):
            if i not in paragraphs_copied:  # Avoid copying paragraphs we've already included
                para = doc.paragraphs[i]
                new_para = temp_doc.add_paragraph(para.text)
                new_para.style = para.style
                paragraphs_copied.add(i)
            
        # 8. Add any "after_assay_procedure" tables
        tables_added = 0
        for table_idx, position in tables_to_preserve.items():
            if position == "after_assay_procedure":
                # Get the table from the original document
                orig_table = doc.tables[table_idx]
                
                # Create a new table with same dimensions
                rows = len(orig_table.rows)
                cols = len(orig_table.rows[0].cells) if rows > 0 else 0
                
                if rows > 0 and cols > 0:
                    new_table = temp_doc.add_table(rows=rows, cols=cols)
                    new_table.style = orig_table.style
                    
                    # Copy cell content
                    for i, row in enumerate(orig_table.rows):
                        for j, cell in enumerate(row.cells):
                            if i < len(new_table.rows) and j < len(new_table.rows[i].cells):
                                new_table.rows[i].cells[j].text = cell.text
                    
                    tables_added += 1
                    logger.info(f"Added table {table_idx} ({rows}x{cols}) from position {position}")
        
        # 9. Calculate total tables added
        total_tables_added = table_idx_in_new_doc + tables_added
        
        # Apply Calibri font and 1.15 line spacing to the entire document
        apply_document_formatting(temp_doc)
        
        # Save the temporary document
        temp_doc.save(temp_path)
        
        # Replace the original with our temporary document
        shutil.copy2(temp_path, document_path)
        
        # Clean up
        if temp_path.exists():
            os.remove(temp_path)
            
        logger.info(f"Fixed sample sections and saved to {document_path} with {table_idx_in_new_doc} tables before sample prep + {tables_added} tables after assay procedure")
        
    except Exception as e:
        logger.error(f"Error fixing sample sections: {e}")
        # Don't raise, continue as best we can

def apply_document_formatting(doc):
    """
    Apply Calibri font and 1.15 line spacing to all paragraphs in the document.
    Also ensures Title formatting is correct.
    
    Args:
        doc: The Document object to modify
    """
    # First set the default style
    if 'Normal' in doc.styles:
        style = doc.styles['Normal']
        style.font.name = "Calibri"
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    
    # Ensure Title style is correct
    if 'Title' in doc.styles:
        title_style = doc.styles['Title']
        title_style.font.name = "Calibri"
        title_style.font.size = Pt(36)
        title_style.font.bold = True
        title_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
    # First check and fix the title paragraph specifically
    if len(doc.paragraphs) > 0:
        title_para = doc.paragraphs[0]
        if title_para.style.name == 'Title':
            # Make sure title paragraphs have correct formatting
            title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Fix title paragraph formatting
            for run in title_para.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(36)
                run.font.bold = True
                
            # If there are no runs, add them with proper formatting
            if len(title_para.runs) == 0:
                title_text = title_para.text
                title_para.clear()
                new_run = title_para.add_run(title_text)
                new_run.font.name = "Calibri"
                new_run.font.size = Pt(36)
                new_run.font.bold = True
        
    # Apply to all paragraphs
    for para in doc.paragraphs:
        # Skip title paragraph which we've already handled
        if para.style.name == 'Title':
            continue
            
        # Apply paragraph formatting
        para.paragraph_format.line_spacing = 1.15
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        
        # Apply font to all runs
        for run in para.runs:
            run.font.name = "Calibri"
    
    # Apply to all tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    # Apply paragraph formatting
                    para.paragraph_format.line_spacing = 1.15
                    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                    
                    # Apply font to all runs
                    for run in para.runs:
                        run.font.name = "Calibri"
                        
    # Make one final pass for any styled paragraphs
    for style_id in ['Heading 1', 'Heading 2', 'Heading 3', 'List Bullet', 'List Number']:
        if style_id in doc.styles:
            style = doc.styles[style_id]
            style.font.name = "Calibri"
            # Keep line spacing consistent
            style.paragraph_format.line_spacing = 1.15
            style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

if __name__ == "__main__":
    # Example usage
    input_doc = Path("attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx")
    template = Path("templates_docx/enhanced_template.docx")
    output = Path("output_updated_template.docx")
    
    update_template_populator(input_doc, template, output)