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
from pathlib import Path

from elisa_parser import ELISADatasheetParser
from template_populator import TemplatePopulator

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
    template_path = Path(args.template)
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
        # Parse the ELISA datasheet
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
            kit_name=args.kit_name if hasattr(args, 'kit_name') else None,
            catalog_number=args.catalog_number if hasattr(args, 'catalog_number') else None,
            lot_number=args.lot_number if hasattr(args, 'lot_number') else None
        )
        
        logger.info(f"Successfully generated populated template at: {output_path}")
        return 0
        
    except Exception as e:
        logger.exception(f"Error processing files: {e}")
        return 1

if __name__ == "__main__":
    exit(main())
