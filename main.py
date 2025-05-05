#!/usr/bin/env python3
"""
ELISA Kit Datasheet Processor and Web Application
------------------------------------------------
Provides both command-line and web interfaces for processing ELISA kit datasheets.

Command-line usage:
    python main.py --source SOURCE_FILE --template TEMPLATE_FILE --output OUTPUT_FILE

Web application:
    The Flask web application is defined in app.py and imported here for Gunicorn.
"""

import os
import argparse
import logging
import re
import sys
from pathlib import Path

from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator
from updated_template_populator import update_template_populator
from red_dot_template_populator import populate_red_dot_template

# Import Flask app for Gunicorn
from app import app

def setup_logging():
    """Configure logging for the application"""
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description='Extract data from ELISA kit datasheets and populate DOCX templates'
    )
    
    parser.add_argument(
        '--source', 
        required=True,
        help='Path to the source ELISA kit datasheet DOCX file'
    )
    
    parser.add_argument(
        '--template', 
        required=True,
        help='Path to the template DOCX file to be populated'
    )
    
    parser.add_argument(
        '--output', 
        required=True,
        help='Path where the populated DOCX file will be saved'
    )
    
    parser.add_argument(
        '--kit-name', 
        help='Name of the ELISA kit (e.g., "Mouse KLK1 ELISA Kit")'
    )
    
    parser.add_argument(
        '--catalog-number', 
        help='Catalog number of the ELISA kit (e.g., "EK1586")'
    )
    
    parser.add_argument(
        '--lot-number', 
        help='Lot number of the ELISA kit'
    )
    
    parser.add_argument(
        '--debug', 
        action='store_true',
        help='Enable debug mode with additional logging'
    )
    
    return parser.parse_args()

def main():
    """Main entry point for the command-line application"""
    # Set up logging
    setup_logging()
    logger = logging.getLogger(__name__)
    
    # Parse command line arguments
    args = parse_arguments()
    
    # Validate file paths
    source_path = Path(args.source)
    
    # Check if template is the default boster template and replace with enhanced template
    template_path = Path(args.template)
    if template_path.name == "boster_template_ready.docx" and Path("templates_docx/enhanced_template.docx").exists():
        logger.info("Using enhanced template instead of basic boster template")
        template_path = Path("templates_docx/enhanced_template.docx")
        
    output_path = Path(args.output)
    
    if not source_path.exists():
        logger.error(f"Source file does not exist: {source_path}")
        return 1
    
    if not template_path.exists():
        logger.error(f"Template file does not exist: {template_path}")
        return 1
    
    # Create output directory if it doesn't exist
    output_dir = output_path.parent
    if not output_dir.exists():
        logger.info(f"Creating output directory: {output_dir}")
        output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Get custom parameters if provided
        kit_name = args.kit_name if hasattr(args, 'kit_name') and args.kit_name else None
        catalog_number = args.catalog_number if hasattr(args, 'catalog_number') and args.catalog_number else None
        lot_number = args.lot_number if hasattr(args, 'lot_number') and args.lot_number else None
        
        # If both catalog and lot number are provided, rename the output file
        if catalog_number and lot_number and str(output_path) == str(args.output):
            # Create new output path with catalog-lot format
            output_dir = output_path.parent
            new_filename = f"{catalog_number}-{lot_number}.docx"
            output_path = output_dir / new_filename
            logger.info(f"Using catalog/lot number format for output: {output_path}")
            
        # Log the parameters we're using
        if kit_name:
            logger.info(f"Using custom kit name: {kit_name}")
        if catalog_number:
            logger.info(f"Using custom catalog number: {catalog_number}")
        if lot_number:
            logger.info(f"Using custom lot number: {lot_number}")
        
        # Check if we should use the Red Dot template populator
        is_red_dot_template = template_path.name.lower() == 'red_dot_template.docx'
        is_red_dot_document = "RDR" in source_path.name.upper() or source_path.name.upper().endswith('RDR.DOCX')
        
        if is_red_dot_template or is_red_dot_document:
            # Use Red Dot template populator
            logger.info("Using Red Dot template populator")
            success = populate_red_dot_template(
                source_path=source_path, 
                template_path=template_path, 
                output_path=output_path,
                kit_name=kit_name if kit_name else "",
                catalog_number=catalog_number if catalog_number else "",
                lot_number=lot_number if lot_number else ""
            )
            
            if not success:
                logger.error("Error populating Red Dot template")
                return 1
                
            # Apply comprehensive fixes for Red Dot documents
            logger.info("Applying comprehensive fixes for Red Dot document")
            
            # Apply fix for ASSAY PROCEDURE vs ASSAY PROCEDURE SUMMARY
            from fix_assay_procedure import fix_assay_sections_in_document
            if fix_assay_sections_in_document(output_path):
                logger.info(f"Successfully fixed ASSAY PROCEDURE sections in: {output_path}")
            else:
                logger.warning("Could not properly fix ASSAY PROCEDURE sections")
                
            # Apply all other comprehensive fixes (footer, formatting, table placement)
            from fix_red_dot_document_comprehensive import fix_red_dot_document
            if fix_red_dot_document(output_path):
                logger.info(f"Applied comprehensive formatting fixes to: {output_path}")
            else:
                logger.warning("Could not apply all formatting fixes, document may need manual adjustment")
        else:
            # Use standard template populator for non-Red Dot documents
            logger.info(f"Parsing ELISA datasheet: {source_path}")
            parser = ELISADatasheetParser(source_path)
            data = parser.extract_data()
            
            # Populate the template with extracted data
            logger.info(f"Populating template: {template_path}")
            populator = TemplatePopulator(template_path)
            
            # Pass optional user-provided values
            populator.populate(
                data, 
                output_path,
                kit_name=kit_name,
                catalog_number=catalog_number,
                lot_number=lot_number
            )
            
            # Fix the sample preparation and dilution sections
            logger.info("Fixing sample preparation and dilution sections")
            update_template_populator(source_path, output_path, output_path)
            
            # Add ASSAY PRINCIPLE section
            logger.info("Adding ASSAY PRINCIPLE section")
            from add_assay_principle import add_assay_principle
            add_assay_principle(output_path)
            
            # Fix OVERVIEW table
            logger.info("Fixing OVERVIEW table with correct data")
            from fix_overview_table import fix_overview_table
            fix_overview_table(output_path)
            
            # Fix document structure to ensure tables are properly positioned
            logger.info("Fixing document structure and table positions")
            from fix_document_structure import ensure_sections_with_tables
            ensure_sections_with_tables(output_path)
            
            # Modify footer text
            logger.info("Modifying footer text")
            from modify_footer import modify_footer_text
            modify_footer_text(output_path)
        
        # Create a date-based version of the output for preservation
        from datetime import datetime
        if catalog_number:
            today = datetime.now().strftime("%Y%m%d")
            catalog_based_filename = f"{catalog_number}-{today}.docx"
            output_dir = output_path.parent
            dated_path = output_dir / catalog_based_filename
            
            # Skip if the output path is already using the same filename
            if str(output_path) != str(dated_path):
                import shutil
                shutil.copy2(output_path, dated_path)
                logger.info(f"Copy saved with date-based filename: {catalog_based_filename}")
            else:
                logger.info(f"Output file already using date-based filename: {catalog_based_filename}")
        
        logger.info(f"Successfully generated populated template at: {output_path}")
        return 0
        
    except Exception as e:
        logger.exception(f"Error processing files: {e}")
        return 1

if __name__ == "__main__":
    exit(main())
