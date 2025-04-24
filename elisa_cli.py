#!/usr/bin/env python3
"""
ELISA Kit Datasheet Parser CLI
-----------------------------
A simple interactive CLI application for extracting data from ELISA kit datasheets.
This is a lightweight alternative to the GUI version for environments where
graphical interfaces are not available.
"""

import os
import sys
import logging
import shutil
import time
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

# Import our existing ELISA parser components
from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator
from docx_templates import get_available_templates, initialize_templates

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Define paths
UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
TEMPLATE_FOLDER = Path('templates_docx')
ASSETS_FOLDER = Path('attached_assets')
BATCH_FOLDER = Path('batch_outputs')
DEFAULT_TEMPLATE = TEMPLATE_FOLDER / 'enhanced_template.docx'

# Create folders if they don't exist
for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_FOLDER, BATCH_FOLDER]:
    folder.mkdir(exist_ok=True)

# Initialize templates
initialize_templates(TEMPLATE_FOLDER, ASSETS_FOLDER)

# Make sure the enhanced template is the default
if not DEFAULT_TEMPLATE.exists():
    logger.warning(f"Default enhanced template not found at {DEFAULT_TEMPLATE}")
    logger.info("Looking for any available template to use as default...")
    templates = list(TEMPLATE_FOLDER.glob('*.docx'))
    if templates:
        DEFAULT_TEMPLATE = templates[0]
        logger.info(f"Using {DEFAULT_TEMPLATE.name} as the default template")
    else:
        logger.warning("No templates found. The application may not work correctly.")


class ELISAParserCLI:
    """Command-line interface for the ELISA Kit Datasheet Parser."""
    
    def __init__(self):
        """Initialize the CLI interface."""
        self.source_path = None
        self.template_path = DEFAULT_TEMPLATE
        self.output_path = None
        self.kit_name = None
        self.catalog_number = None
        self.lot_number = None
        self.batch_source_paths = []
        self.use_metadata = True
    
    def clear_screen(self):
        """Clear the terminal screen."""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def print_header(self):
        """Print the application header."""
        self.clear_screen()
        print("=" * 70)
        print("             ELISA KIT DATASHEET PARSER - INTERACTIVE CLI")
        print("=" * 70)
        print()
    
    def print_menu(self):
        """Print the main menu."""
        self.print_header()
        print("Main Menu:")
        print("1. Process Single ELISA Datasheet")
        print("2. Batch Process Multiple ELISA Datasheets")
        print("3. Manage Templates")
        print("4. View Help")
        print("0. Exit")
        print()
    
    def process_single_menu(self):
        """Display the single file processing menu."""
        while True:
            self.print_header()
            print("Process Single ELISA Datasheet")
            print("-" * 50)
            
            # Display current settings
            print("Current Settings:")
            print(f"Source File:     {self.source_path or 'Not selected'}")
            print(f"Template:        {self.template_path.name if self.template_path else 'Default'}")
            print(f"Output Path:     {self.output_path or 'Auto-generated'}")
            print(f"Kit Name:        {self.kit_name or 'Extract from file'}")
            print(f"Catalog Number:  {self.catalog_number or 'Extract from file'}")
            print(f"Lot Number:      {self.lot_number or 'Extract from file'}")
            print(f"Use Metadata:    {'Yes' if self.use_metadata else 'No'}")
            print()
            
            print("Options:")
            print("1. Select Source File")
            print("2. Select Template")
            print("3. Set Output Path")
            print("4. Set Kit Name")
            print("5. Set Catalog Number")
            print("6. Set Lot Number")
            print("7. Toggle Metadata Usage")
            print("8. Process File")
            print("0. Back to Main Menu")
            print()
            
            choice = input("Enter your choice (0-8): ").strip()
            
            if choice == '0':
                break
            elif choice == '1':
                self.select_source_file()
            elif choice == '2':
                self.select_template()
            elif choice == '3':
                self.set_output_path()
            elif choice == '4':
                self.kit_name = input("Enter Kit Name (or leave empty to extract from file): ").strip() or None
            elif choice == '5':
                self.catalog_number = input("Enter Catalog Number (or leave empty to extract from file): ").strip() or None
            elif choice == '6':
                self.lot_number = input("Enter Lot Number (or leave empty to extract from file): ").strip() or None
            elif choice == '7':
                self.use_metadata = not self.use_metadata
                print(f"Metadata usage set to: {'Yes' if self.use_metadata else 'No'}")
                input("Press Enter to continue...")
            elif choice == '8':
                if not self.source_path:
                    print("Error: No source file selected.")
                    input("Press Enter to continue...")
                    continue
                
                self.process_single_file()
                input("Press Enter to continue...")
    
    def batch_process_menu(self):
        """Display the batch processing menu."""
        while True:
            self.print_header()
            print("Batch Process Multiple ELISA Datasheets")
            print("-" * 50)
            
            # Display current settings
            print("Current Settings:")
            print(f"Number of Files: {len(self.batch_source_paths)}")
            print(f"Template:        {self.template_path.name if self.template_path else 'Default'}")
            print(f"Use Metadata:    {'Yes' if self.use_metadata else 'No'}")
            print()
            
            print("Options:")
            print("1. Select Source Files")
            print("2. Select Template")
            print("3. Toggle Metadata Usage")
            print("4. Process Files")
            print("0. Back to Main Menu")
            print()
            
            choice = input("Enter your choice (0-4): ").strip()
            
            if choice == '0':
                break
            elif choice == '1':
                self.select_batch_files()
            elif choice == '2':
                self.select_template()
            elif choice == '3':
                self.use_metadata = not self.use_metadata
                print(f"Metadata usage set to: {'Yes' if self.use_metadata else 'No'}")
                input("Press Enter to continue...")
            elif choice == '4':
                if not self.batch_source_paths:
                    print("Error: No source files selected.")
                    input("Press Enter to continue...")
                    continue
                
                self.process_batch_files()
                input("Press Enter to continue...")
    
    def templates_menu(self):
        """Display the templates management menu."""
        while True:
            self.print_header()
            print("Template Management")
            print("-" * 50)
            
            # List available templates
            templates = get_available_templates(TEMPLATE_FOLDER)
            print("Available Templates:")
            for i, template in enumerate(templates, 1):
                print(f"{i}. {template['name']} - {template['description']}")
            print()
            
            print("Options:")
            print("1. Make a Template Default")
            print("0. Back to Main Menu")
            print()
            
            choice = input("Enter your choice (0-1): ").strip()
            
            if choice == '0':
                break
            elif choice == '1':
                if not templates:
                    print("No templates available.")
                    input("Press Enter to continue...")
                    continue
                
                try:
                    template_idx = int(input(f"Enter template number (1-{len(templates)}): ").strip()) - 1
                    if 0 <= template_idx < len(templates):
                        self.template_path = TEMPLATE_FOLDER / templates[template_idx]['name']
                        print(f"Default template set to: {templates[template_idx]['name']}")
                    else:
                        print("Invalid template number.")
                except ValueError:
                    print("Invalid input. Please enter a number.")
                
                input("Press Enter to continue...")
    
    def help_menu(self):
        """Display the help information."""
        self.print_header()
        print("ELISA Kit Datasheet Parser - Help")
        print("-" * 50)
        print()
        print("This application extracts data from ELISA kit datasheets and populates DOCX templates.")
        print()
        print("Single File Processing:")
        print("  - Select a source ELISA datasheet in DOCX format")
        print("  - Choose a template or use the default enhanced template")
        print("  - Optionally provide a kit name, catalog number, and lot number")
        print("  - The application will extract data and generate a formatted output document")
        print()
        print("Batch Processing:")
        print("  - Select multiple source ELISA datasheets")
        print("  - Choose a template to apply to all files")
        print("  - The application will process all files and generate output documents")
        print()
        print("Metadata Usage:")
        print("  - When enabled, the application will use extracted metadata (catalog/lot numbers)")
        print("    for naming output files")
        print("  - When disabled, the application will use generic filenames")
        print()
        input("Press Enter to return to main menu...")
    
    def select_source_file(self):
        """Ask the user to select a source file."""
        self.print_header()
        print("Select Source File")
        print("-" * 50)
        print("Enter the path to the ELISA datasheet DOCX file:")
        
        while True:
            source_path = input("> ").strip()
            
            if not source_path:
                return
            
            source_path = Path(source_path)
            
            if not source_path.exists():
                print(f"Error: File {source_path} does not exist.")
                continue
            
            if source_path.suffix.lower() != '.docx':
                print("Warning: File is not a DOCX file. It may not work correctly.")
            
            self.source_path = source_path
            
            # Try to extract catalog/lot numbers from filename
            if not self.catalog_number or not self.lot_number:
                filename = source_path.stem
                if '-' in filename:
                    parts = filename.split('-')
                    if len(parts) >= 2:
                        # Assume format like "EK1586-6058725" (catalog-lot)
                        if not self.catalog_number:
                            self.catalog_number = parts[0]
                        if not self.lot_number and len(parts) > 1:
                            self.lot_number = parts[1]
            
            print(f"Source file set to: {self.source_path}")
            input("Press Enter to continue...")
            break
    
    def select_batch_files(self):
        """Ask the user to select multiple source files."""
        self.print_header()
        print("Select Batch Files")
        print("-" * 50)
        print("Enter the paths to the ELISA datasheet DOCX files (one per line).")
        print("Enter a blank line when finished.")
        
        self.batch_source_paths = []
        
        print("Enter file paths (or leave empty to finish):")
        while True:
            source_path = input("> ").strip()
            
            if not source_path:
                break
            
            source_path = Path(source_path)
            
            if not source_path.exists():
                print(f"Error: File {source_path} does not exist.")
                continue
            
            if source_path.suffix.lower() != '.docx':
                print("Warning: File is not a DOCX file. It may not work correctly.")
            
            self.batch_source_paths.append(source_path)
            print(f"Added: {source_path}")
        
        print(f"Added {len(self.batch_source_paths)} files for batch processing.")
        input("Press Enter to continue...")
    
    def select_template(self):
        """Ask the user to select a template."""
        self.print_header()
        print("Select Template")
        print("-" * 50)
        
        # List available templates
        templates = get_available_templates(TEMPLATE_FOLDER)
        print("Available Templates:")
        for i, template in enumerate(templates, 1):
            print(f"{i}. {template['name']} - {template['description']}")
        print()
        
        try:
            template_idx = int(input(f"Enter template number (1-{len(templates)}, or 0 to cancel): ").strip())
            if template_idx == 0:
                return
            
            template_idx -= 1  # Adjust for 0-based indexing
            if 0 <= template_idx < len(templates):
                self.template_path = TEMPLATE_FOLDER / templates[template_idx]['name']
                print(f"Template set to: {templates[template_idx]['name']}")
            else:
                print("Invalid template number.")
        except ValueError:
            print("Invalid input. Please enter a number.")
        
        input("Press Enter to continue...")
    
    def set_output_path(self):
        """Ask the user to set an output path."""
        self.print_header()
        print("Set Output Path")
        print("-" * 50)
        print("Enter the path for the output file (or leave empty for auto-generated):")
        
        output_path = input("> ").strip()
        
        if not output_path:
            self.output_path = None
            print("Output path will be auto-generated.")
        else:
            output_path = Path(output_path)
            
            # Ensure the file has a .docx extension
            if output_path.suffix.lower() != '.docx':
                output_path = output_path.with_suffix('.docx')
            
            self.output_path = output_path
            print(f"Output path set to: {self.output_path}")
        
        input("Press Enter to continue...")
    
    def process_single_file(self):
        """Process a single ELISA datasheet."""
        try:
            # Determine output path
            if self.output_path:
                output_path = self.output_path
            else:
                # Generate output filename based on catalog and lot numbers if provided
                catalog_number = self.catalog_number
                lot_number = self.lot_number
                
                if catalog_number and lot_number and self.use_metadata:
                    output_filename = f"{catalog_number}-{lot_number}.docx"
                else:
                    # Fall back to default naming
                    unique_id = os.urandom(4).hex()
                    output_filename = f"output_{unique_id}.docx"
                
                output_path = OUTPUT_FOLDER / output_filename
            
            print(f"Processing file: {self.source_path}")
            print(f"Using template: {self.template_path}")
            print(f"Output will be saved to: {output_path}")
            print()
            
            # Parse the datasheet
            print("Parsing ELISA datasheet...")
            parser = ELISADatasheetParser(self.source_path)
            data = parser.extract_data()
            
            print("Populating template...")
            
            # Populate the template
            populator = TemplatePopulator(self.template_path)
            populator.populate(
                data, 
                output_path,
                kit_name=self.kit_name,
                catalog_number=self.catalog_number,
                lot_number=self.lot_number
            )
            
            print(f"Processing complete. Output saved to: {output_path}")
            return True
        
        except Exception as e:
            logger.exception(f"Error processing file: {e}")
            print(f"Error: {str(e)}")
            return False
    
    def process_batch_files(self):
        """Process multiple ELISA datasheets."""
        if not self.batch_source_paths:
            print("No files selected for batch processing.")
            return
        
        # Create a batch output directory
        unique_id = os.urandom(4).hex()
        batch_output_dir = BATCH_FOLDER / f"batch_{unique_id}"
        batch_output_dir.mkdir(exist_ok=True)
        
        total_files = len(self.batch_source_paths)
        successful_files = 0
        failed_files = 0
        output_paths = []
        
        print(f"Processing {total_files} files...")
        print(f"Output directory: {batch_output_dir}")
        print()
        
        for i, source_path in enumerate(self.batch_source_paths):
            try:
                print(f"[{i+1}/{total_files}] Processing {source_path.name}...")
                
                # Parse the datasheet
                parser = ELISADatasheetParser(source_path)
                data = parser.extract_data()
                
                # Extract catalog number for filename if available and requested
                catalog_number = data.get('catalog_number', '')
                output_filename = f"output_{i}_{source_path.stem}.docx"
                
                if self.use_metadata and catalog_number:
                    output_filename = f"{catalog_number}.docx"
                
                output_path = batch_output_dir / output_filename
                
                # Populate the template
                populator = TemplatePopulator(self.template_path)
                populator.populate(data, output_path)
                
                output_paths.append(output_path)
                successful_files += 1
                print(f"  Success: Output saved to {output_path.name}")
                
            except Exception as e:
                logger.exception(f"Error processing file {source_path}: {e}")
                print(f"  Error: {str(e)}")
                failed_files += 1
        
        print()
        print(f"Batch processing complete.")
        print(f"Successfully processed: {successful_files}/{total_files} files")
        print(f"Failed: {failed_files}/{total_files} files")
        print(f"Output directory: {batch_output_dir}")
        
        return successful_files > 0
    
    def run(self):
        """Run the CLI application main loop."""
        while True:
            self.print_menu()
            choice = input("Enter your choice (0-4): ").strip()
            
            if choice == '0':
                print("Exiting application...")
                break
            elif choice == '1':
                self.process_single_menu()
            elif choice == '2':
                self.batch_process_menu()
            elif choice == '3':
                self.templates_menu()
            elif choice == '4':
                self.help_menu()
            else:
                print("Invalid choice. Please try again.")
                time.sleep(1)


def main():
    """Main entry point for the CLI application."""
    cli = ELISAParserCLI()
    cli.run()


if __name__ == "__main__":
    main()