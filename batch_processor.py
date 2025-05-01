"""
Batch Processing Module
----------------------
Handles batch processing of multiple ELISA datasheets at once.
"""

import os
import logging
import uuid
from pathlib import Path
from typing import List, Dict, Any, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator
from updated_template_populator import update_template_populator

# Configure logging
logger = logging.getLogger(__name__)

class BatchProcessor:
    """
    Processes multiple ELISA datasheets in batch, applying the same template to all.
    
    Provides both synchronous and asynchronous processing options with progress tracking.
    """
    
    def __init__(self, 
                 template_path: Path, 
                 output_dir: Path,
                 max_workers: int = 4):
        """
        Initialize the batch processor.
        
        Args:
            template_path: Path to the template to use for all files
            output_dir: Directory where output files will be saved
            max_workers: Maximum number of concurrent worker threads
        """
        self.template_path = template_path
        self.output_dir = output_dir
        self.max_workers = max_workers
        self.results = {}
        self.progress = {}
        
        # Ensure output directory exists
        self.output_dir.mkdir(exist_ok=True)
        
        logger.info(f"Batch processor initialized with template: {template_path.name}")
        
    def process_file(self, 
                    file_path: Path, 
                    output_filename: str = None,
                    kit_name: str = None,
                    catalog_number: str = None, 
                    lot_number: str = None) -> Tuple[bool, str, Path]:
        """
        Process a single file in the batch.
        
        Args:
            file_path: Path to the ELISA datasheet to process
            output_filename: Optional custom filename for the output, if None one will be generated
            kit_name: Optional kit name to override extracted value
            catalog_number: Optional catalog number to override extracted value
            lot_number: Optional lot number to override extracted value
            
        Returns:
            Tuple containing (success_status, error_message_if_any, output_path)
        """
        batch_id = str(uuid.uuid4())[:8]
        self.progress[batch_id] = {
            'file': file_path.name,
            'status': 'processing',
            'progress': 0,
            'message': f'Processing {file_path.name}'
        }
        
        try:
            logger.info(f"Processing file: {file_path}")
            self.progress[batch_id]['progress'] = 10
            
            # Generate output filename if not provided
            if not output_filename:
                if catalog_number and lot_number:
                    output_filename = f"{catalog_number}-{lot_number}.docx"
                else:
                    output_filename = f"output_{batch_id}_{file_path.stem}.docx"
            
            output_path = self.output_dir / output_filename
            
            # Parse the ELISA datasheet
            parser = ELISADatasheetParser(file_path)
            self.progress[batch_id]['progress'] = 30
            self.progress[batch_id]['message'] = f'Extracting data from {file_path.name}'
            
            data = parser.extract_data()
            self.progress[batch_id]['progress'] = 60
            
            # Extract catalog number and lot number from data if not provided
            extracted_catalog = data.get('catalog_number', '')
            if not catalog_number and extracted_catalog:
                catalog_number = extracted_catalog
                
            # Apply the template
            self.progress[batch_id]['progress'] = 70
            self.progress[batch_id]['message'] = f'Populating template for {file_path.name}'
            
            # Check if we're using the Red Dot template or document
            is_red_dot_template = self.template_path.name.lower() == 'red_dot_template.docx'
            is_red_dot_document = "RDR" in file_path.name.upper() or file_path.name.upper().endswith('RDR.DOCX')
            
            if is_red_dot_template or is_red_dot_document:
                logger.info("Using Red Dot template populator for batch processing")
                # Import the Red Dot template populator
                from red_dot_template_populator import populate_red_dot_template
                
                # If document is Red Dot but template isn't, use the Red Dot template
                if is_red_dot_document and not is_red_dot_template:
                    red_dot_template_path = Path("templates_docx/red_dot_template.docx")
                    if red_dot_template_path.exists():
                        logger.info(f"Switching to Red Dot template for document {file_path.name}")
                        template_to_use = red_dot_template_path
                    else:
                        template_to_use = self.template_path
                else:
                    template_to_use = self.template_path
                
                # Populate the template with the Red Dot populator
                success = populate_red_dot_template(
                    source_path=file_path,
                    template_path=template_to_use, 
                    output_path=output_path,
                    kit_name=kit_name if kit_name else "",
                    catalog_number=catalog_number if catalog_number else "",
                    lot_number=lot_number if lot_number else ""
                )
                
                if not success:
                    return False, "Error populating Red Dot template", output_path
            else:
                # Use the standard template populator for other templates
                populator = TemplatePopulator(self.template_path)
                populator.populate(
                    data, 
                    output_path,
                    kit_name=kit_name,
                    catalog_number=catalog_number,
                    lot_number=lot_number
                )
            
            # Apply additional processing only for standard templates (not Red Dot)
            is_red_dot_template = self.template_path.name.lower() == 'red_dot_template.docx'
            is_red_dot_document = "RDR" in file_path.name.upper() or file_path.name.upper().endswith('RDR.DOCX')
            
            if not is_red_dot_template and not is_red_dot_document:
                self.progress[batch_id]['progress'] = 85
                self.progress[batch_id]['message'] = f'Applying enhancements for {file_path.name}'
                update_template_populator(file_path, output_path, output_path)
                
                # Add ASSAY PRINCIPLE section
                from add_assay_principle import add_assay_principle
                add_assay_principle(output_path)
                
                # Fix OVERVIEW table
                from fix_overview_table import fix_overview_table
                fix_overview_table(output_path)
                
                # Fix document structure to ensure tables are properly positioned
                from fix_document_structure import ensure_sections_with_tables
                ensure_sections_with_tables(output_path)
            else:
                # For Red Dot templates/documents, just update progress and modify footer
                self.progress[batch_id]['progress'] = 85
                self.progress[batch_id]['message'] = f'Red Dot document already fully populated for {file_path.name}'
                # Modify footer text
                from modify_footer import modify_footer_text
                modify_footer_text(output_path)
            
            self.progress[batch_id]['progress'] = 100
            self.progress[batch_id]['status'] = 'completed'
            self.progress[batch_id]['message'] = f'Successfully processed {file_path.name}'
            self.progress[batch_id]['output'] = str(output_path)
            
            logger.info(f"Successfully processed {file_path} to {output_path}")
            return True, '', output_path
            
        except Exception as e:
            error_msg = f"Error processing {file_path}: {str(e)}"
            logger.exception(error_msg)
            
            self.progress[batch_id]['status'] = 'failed'
            self.progress[batch_id]['message'] = error_msg
            
            return False, error_msg, None
    
    def process_batch(self, 
                     file_paths: List[Path], 
                     output_filenames: List[str] = None,
                     kit_names: List[str] = None,
                     catalog_numbers: List[str] = None,
                     lot_numbers: List[str] = None) -> Dict[str, Any]:
        """
        Process a batch of files sequentially.
        
        Args:
            file_paths: List of paths to ELISA datasheets to process
            output_filenames: Optional list of custom filenames for outputs
            kit_names: Optional list of kit names to override extracted values
            catalog_numbers: Optional list of catalog numbers
            lot_numbers: Optional list of lot numbers
            
        Returns:
            Dictionary with results of the batch processing
        """
        results = {
            'total': len(file_paths),
            'successful': 0,
            'failed': 0,
            'files': []
        }
        
        # Initialize default parameter lists if not provided
        if not output_filenames:
            output_filenames = [None] * len(file_paths)
        if not kit_names:
            kit_names = [None] * len(file_paths)
        if not catalog_numbers:
            catalog_numbers = [None] * len(file_paths)
        if not lot_numbers:
            lot_numbers = [None] * len(file_paths)
            
        # Process each file
        for i, file_path in enumerate(file_paths):
            success, error, output_path = self.process_file(
                file_path,
                output_filenames[i] if i < len(output_filenames) else None,
                kit_names[i] if i < len(kit_names) else None,
                catalog_numbers[i] if i < len(catalog_numbers) else None,
                lot_numbers[i] if i < len(lot_numbers) else None
            )
            
            file_result = {
                'file': str(file_path),
                'success': success
            }
            
            if success:
                results['successful'] += 1
                file_result['output'] = str(output_path)
            else:
                results['failed'] += 1
                file_result['error'] = error
                
            results['files'].append(file_result)
            
        return results
    
    def process_batch_parallel(self, 
                              file_paths: List[Path], 
                              output_filenames: List[str] = None,
                              kit_names: List[str] = None,
                              catalog_numbers: List[str] = None,
                              lot_numbers: List[str] = None) -> Dict[str, Any]:
        """
        Process a batch of files in parallel using a thread pool.
        
        Args:
            file_paths: List of paths to ELISA datasheets to process
            output_filenames: Optional list of custom filenames for outputs
            kit_names: Optional list of kit names to override extracted values
            catalog_numbers: Optional list of catalog numbers
            lot_numbers: Optional list of lot numbers
            
        Returns:
            Dictionary with results of the batch processing
        """
        results = {
            'total': len(file_paths),
            'successful': 0,
            'failed': 0,
            'files': []
        }
        
        # Initialize default parameter lists if not provided
        if not output_filenames:
            output_filenames = [None] * len(file_paths)
        if not kit_names:
            kit_names = [None] * len(file_paths)
        if not catalog_numbers:
            catalog_numbers = [None] * len(file_paths)
        if not lot_numbers:
            lot_numbers = [None] * len(file_paths)
        
        # Process files in parallel
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks to the thread pool
            future_to_file = {
                executor.submit(
                    self.process_file,
                    file_path,
                    output_filenames[i] if i < len(output_filenames) else None,
                    kit_names[i] if i < len(kit_names) else None,
                    catalog_numbers[i] if i < len(catalog_numbers) else None,
                    lot_numbers[i] if i < len(lot_numbers) else None
                ): file_path
                for i, file_path in enumerate(file_paths)
            }
            
            # Process results as they complete
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    success, error, output_path = future.result()
                    file_result = {
                        'file': str(file_path),
                        'success': success
                    }
                    
                    if success:
                        results['successful'] += 1
                        file_result['output'] = str(output_path)
                    else:
                        results['failed'] += 1
                        file_result['error'] = error
                        
                    results['files'].append(file_result)
                
                except Exception as e:
                    logger.exception(f"Error processing {file_path}: {e}")
                    results['failed'] += 1
                    results['files'].append({
                        'file': str(file_path),
                        'success': False,
                        'error': str(e)
                    })
        
        return results
    
    def get_progress(self) -> Dict[str, Any]:
        """
        Get the current progress of batch processing.
        
        Returns:
            Dictionary with progress information
        """
        return self.progress