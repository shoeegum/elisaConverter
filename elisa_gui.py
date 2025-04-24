#!/usr/bin/env python3
"""
ELISA Kit Datasheet Parser GUI
-----------------------------
A desktop GUI application for extracting data from ELISA kit datasheets.
"""

import os
import sys
import logging
import platform
from pathlib import Path
from threading import Thread
from typing import Dict, List, Any, Optional

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QPushButton, QLabel, QLineEdit, QFileDialog, QComboBox,
                            QTabWidget, QProgressBar, QMessageBox, QCheckBox, QTableWidget,
                            QTableWidgetItem, QHeaderView, QGroupBox, QFormLayout, QTextEdit,
                            QSplitter, QFrame, QSizePolicy, QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QIcon, QFont, QPixmap

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


class ProcessingWorker(QThread):
    """Worker thread for processing ELISA datasheets in the background."""
    
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, str)
    
    def __init__(self, 
                source_path: Path, 
                template_path: Path, 
                output_path: Path,
                kit_name: Optional[str] = None,
                catalog_number: Optional[str] = None,
                lot_number: Optional[str] = None):
        """
        Initialize the processing worker.
        
        Args:
            source_path: Path to the source document
            template_path: Path to the template document
            output_path: Path where the output will be saved
            kit_name: Optional kit name to override extracted value
            catalog_number: Optional catalog number to override extracted value
            lot_number: Optional lot number to override extracted value
        """
        super().__init__()
        self.source_path = source_path
        self.template_path = template_path
        self.output_path = output_path
        self.kit_name = kit_name
        self.catalog_number = catalog_number
        self.lot_number = lot_number
    
    def run(self):
        """Process the ELISA datasheet."""
        try:
            # Parse the datasheet
            self.status_signal.emit("Parsing ELISA datasheet...")
            self.progress_signal.emit(25)
            
            parser = ELISADatasheetParser(self.source_path)
            data = parser.extract_data()
            
            self.progress_signal.emit(50)
            self.status_signal.emit("Populating template...")
            
            # Populate the template
            populator = TemplatePopulator(self.template_path)
            populator.populate(
                data, 
                self.output_path,
                kit_name=self.kit_name,
                catalog_number=self.catalog_number,
                lot_number=self.lot_number
            )
            
            self.progress_signal.emit(100)
            self.status_signal.emit("Processing complete.")
            self.finished_signal.emit(True, str(self.output_path), "")
            
        except Exception as e:
            logger.exception(f"Error processing file: {e}")
            self.status_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(False, "", str(e))


class BatchProcessingWorker(QThread):
    """Worker thread for batch processing multiple ELISA datasheets."""
    
    overall_progress_signal = pyqtSignal(int)
    file_progress_signal = pyqtSignal(int, int, str)  # file_index, progress, status
    status_signal = pyqtSignal(str)
    file_finished_signal = pyqtSignal(int, bool, str)  # file_index, success, output_path
    all_finished_signal = pyqtSignal(bool, List[str])  # success, list of output paths
    
    def __init__(self, 
                source_paths: List[Path], 
                template_path: Path, 
                output_folder: Path,
                use_metadata: bool = True):
        """
        Initialize the batch processing worker.
        
        Args:
            source_paths: List of paths to the source documents
            template_path: Path to the template document
            output_folder: Folder where the outputs will be saved
            use_metadata: Whether to use metadata from the file for naming
        """
        super().__init__()
        self.source_paths = source_paths
        self.template_path = template_path
        self.output_folder = output_folder
        self.use_metadata = use_metadata
        self.output_paths = []
    
    def run(self):
        """Process multiple ELISA datasheets."""
        total_files = len(self.source_paths)
        processed_files = 0
        successful_files = 0
        
        for i, source_path in enumerate(self.source_paths):
            try:
                self.file_progress_signal.emit(i, 10, f"Parsing {source_path.name}...")
                
                # Parse the datasheet
                parser = ELISADatasheetParser(source_path)
                data = parser.extract_data()
                
                self.file_progress_signal.emit(i, 50, f"Populating template for {source_path.name}...")
                
                # Extract catalog number for filename if available and requested
                catalog_number = data.get('catalog_number', '')
                output_filename = f"output_{i}_{source_path.stem}.docx"
                
                if self.use_metadata and catalog_number:
                    output_filename = f"{catalog_number}.docx"
                
                output_path = self.output_folder / output_filename
                
                # Populate the template
                populator = TemplatePopulator(self.template_path)
                populator.populate(data, output_path)
                
                # Update progress
                processed_files += 1
                successful_files += 1
                self.output_paths.append(str(output_path))
                overall_progress = int((processed_files / total_files) * 100)
                
                self.file_progress_signal.emit(i, 100, "Complete")
                self.overall_progress_signal.emit(overall_progress)
                self.file_finished_signal.emit(i, True, str(output_path))
                
            except Exception as e:
                logger.exception(f"Error processing file {source_path}: {e}")
                processed_files += 1
                overall_progress = int((processed_files / total_files) * 100)
                
                self.file_progress_signal.emit(i, 100, f"Error: {str(e)}")
                self.overall_progress_signal.emit(overall_progress)
                self.file_finished_signal.emit(i, False, str(e))
        
        # Signal completion
        success = successful_files > 0
        self.status_signal.emit(f"Processed {successful_files} of {total_files} files successfully.")
        self.all_finished_signal.emit(success, self.output_paths)


class ELISAParserGUI(QMainWindow):
    """Main GUI window for the ELISA Kit Datasheet Parser."""
    
    def __init__(self):
        """Initialize the main window."""
        super().__init__()
        
        self.setWindowTitle("ELISA Kit Datasheet Parser")
        self.setMinimumSize(800, 600)
        
        # Initialize the main layout
        self.init_ui()
        
        # Initialize state
        self.source_path = None
        self.template_path = DEFAULT_TEMPLATE
        self.output_path = None
        self.processing_worker = None
        self.batch_processing_worker = None
        self.batch_source_paths = []
    
    def init_ui(self):
        """Initialize the user interface."""
        # Create a tab widget for different modes
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Create tabs
        self.single_tab = QWidget()
        self.batch_tab = QWidget()
        self.settings_tab = QWidget()
        
        # Add tabs to widget
        self.tabs.addTab(self.single_tab, "Single File Process")
        self.tabs.addTab(self.batch_tab, "Batch Process")
        self.tabs.addTab(self.settings_tab, "Settings")
        
        # Setup each tab
        self.setup_single_tab()
        self.setup_batch_tab()
        self.setup_settings_tab()
    
    def setup_single_tab(self):
        """Setup the single file processing tab."""
        layout = QVBoxLayout()
        
        # Input section
        input_group = QGroupBox("Input")
        input_layout = QFormLayout()
        
        # Source file selection
        self.source_path_edit = QLineEdit()
        self.source_path_edit.setReadOnly(True)
        source_browse_btn = QPushButton("Browse...")
        source_browse_btn.clicked.connect(self.browse_source)
        
        source_layout = QHBoxLayout()
        source_layout.addWidget(self.source_path_edit, 80)
        source_layout.addWidget(source_browse_btn, 20)
        input_layout.addRow("Source ELISA Datasheet:", source_layout)
        
        # Template selection
        self.template_combo = QComboBox()
        self.load_templates()
        input_layout.addRow("Template:", self.template_combo)
        
        # Additional fields
        self.kit_name_edit = QLineEdit()
        self.catalog_number_edit = QLineEdit()
        self.lot_number_edit = QLineEdit()
        
        input_layout.addRow("Kit Name (Optional):", self.kit_name_edit)
        input_layout.addRow("Catalog Number (Optional):", self.catalog_number_edit)
        input_layout.addRow("Lot Number (Optional):", self.lot_number_edit)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Output section
        output_group = QGroupBox("Output")
        output_layout = QFormLayout()
        
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setReadOnly(True)
        output_browse_btn = QPushButton("Browse...")
        output_browse_btn.clicked.connect(self.browse_output)
        
        output_layout.addRow("Output Folder:", QLabel(str(OUTPUT_FOLDER)))
        
        # Output naming options
        self.use_metadata_check = QCheckBox("Use extracted metadata for output filename")
        self.use_metadata_check.setChecked(True)
        output_layout.addRow("", self.use_metadata_check)
        
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
        # Process button and progress
        process_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process Datasheet")
        self.process_btn.setMinimumHeight(40)
        self.process_btn.clicked.connect(self.process_single)
        process_layout.addWidget(self.process_btn)
        layout.addLayout(process_layout)
        
        # Progress bar
        progress_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.status_label = QLabel("Ready")
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        layout.addLayout(progress_layout)
        
        # Results section
        results_group = QGroupBox("Results")
        results_layout = QVBoxLayout()
        
        self.results_label = QLabel("No results yet")
        self.open_output_btn = QPushButton("Open Output File")
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.clicked.connect(self.open_output)
        
        results_layout.addWidget(self.results_label)
        results_layout.addWidget(self.open_output_btn)
        
        results_group.setLayout(results_layout)
        layout.addWidget(results_group)
        
        # Set layout for the tab
        self.single_tab.setLayout(layout)
    
    def setup_batch_tab(self):
        """Setup the batch processing tab."""
        layout = QVBoxLayout()
        
        # Input section
        input_group = QGroupBox("Batch Input")
        input_layout = QFormLayout()
        
        # Source files selection
        self.batch_files_label = QLabel("No files selected")
        batch_browse_btn = QPushButton("Browse...")
        batch_browse_btn.clicked.connect(self.browse_batch)
        
        input_layout.addRow("Source ELISA Datasheets:", batch_browse_btn)
        input_layout.addRow("Selected Files:", self.batch_files_label)
        
        # Template selection
        self.batch_template_combo = QComboBox()
        self.load_templates(self.batch_template_combo)
        input_layout.addRow("Template:", self.batch_template_combo)
        
        # Batch options
        self.batch_use_metadata_check = QCheckBox("Use extracted metadata for output filenames")
        self.batch_use_metadata_check.setChecked(True)
        self.batch_parallel_check = QCheckBox("Process files in parallel")
        self.batch_parallel_check.setChecked(True)
        
        input_layout.addRow("", self.batch_use_metadata_check)
        input_layout.addRow("", self.batch_parallel_check)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Process button and progress
        process_layout = QHBoxLayout()
        self.batch_process_btn = QPushButton("Process Batch")
        self.batch_process_btn.setMinimumHeight(40)
        self.batch_process_btn.clicked.connect(self.process_batch)
        process_layout.addWidget(self.batch_process_btn)
        layout.addLayout(process_layout)
        
        # Overall progress
        progress_layout = QVBoxLayout()
        progress_layout.addWidget(QLabel("Overall Progress:"))
        self.batch_progress_bar = QProgressBar()
        self.batch_progress_bar.setValue(0)
        self.batch_status_label = QLabel("Ready")
        progress_layout.addWidget(self.batch_progress_bar)
        progress_layout.addWidget(self.batch_status_label)
        layout.addLayout(progress_layout)
        
        # File progress table
        self.files_table = QTableWidget(0, 4)
        self.files_table.setHorizontalHeaderLabels(["File", "Status", "Progress", "Actions"])
        self.files_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.files_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        layout.addWidget(self.files_table)
        
        # Set layout for the tab
        self.batch_tab.setLayout(layout)
    
    def setup_settings_tab(self):
        """Setup the settings tab."""
        layout = QVBoxLayout()
        
        # Settings sections
        templates_group = QGroupBox("Templates")
        templates_layout = QVBoxLayout()
        
        # Template information
        templates_info = QLabel("Available Templates:")
        self.templates_list = QTableWidget(0, 2)
        self.templates_list.setHorizontalHeaderLabels(["Name", "Description"])
        self.templates_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.templates_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        
        # Populate templates list
        self.populate_templates_list()
        
        # Template upload
        template_upload_layout = QHBoxLayout()
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setReadOnly(True)
        template_browse_btn = QPushButton("Browse...")
        template_browse_btn.clicked.connect(self.browse_template)
        template_upload_btn = QPushButton("Upload Template")
        template_upload_btn.clicked.connect(self.upload_template)
        
        template_upload_layout.addWidget(QLabel("Template File:"))
        template_upload_layout.addWidget(self.template_path_edit, 60)
        template_upload_layout.addWidget(template_browse_btn, 20)
        template_upload_layout.addWidget(template_upload_btn, 20)
        
        templates_layout.addWidget(templates_info)
        templates_layout.addWidget(self.templates_list)
        templates_layout.addLayout(template_upload_layout)
        templates_group.setLayout(templates_layout)
        
        # About section
        about_group = QGroupBox("About")
        about_layout = QVBoxLayout()
        
        about_text = """<h3>ELISA Kit Datasheet Parser</h3>
        <p>Version: 1.0</p>
        <p>A tool for extracting data from ELISA kit datasheets and populating DOCX templates.</p>
        <p>This application provides both GUI and command-line interfaces for processing ELISA datasheets.</p>
        <p>&copy; 2025 ELISA Parser</p>
        """
        
        about_label = QLabel(about_text)
        about_label.setTextFormat(Qt.RichText)
        about_label.setWordWrap(True)
        about_layout.addWidget(about_label)
        
        about_group.setLayout(about_layout)
        
        # Add groups to layout
        layout.addWidget(templates_group, 70)
        layout.addWidget(about_group, 30)
        
        # Set layout for the tab
        self.settings_tab.setLayout(layout)
    
    def load_templates(self, combo_box=None):
        """Load available templates into the combo box."""
        if combo_box is None:
            combo_box = self.template_combo
        
        combo_box.clear()
        templates = get_available_templates(TEMPLATE_FOLDER)
        
        for template in templates:
            combo_box.addItem(template['description'], template['name'])
            
            # Set enhanced template as default
            if template['name'] == DEFAULT_TEMPLATE.name:
                combo_box.setCurrentIndex(combo_box.count() - 1)
    
    def populate_templates_list(self):
        """Populate the templates list in the settings tab."""
        self.templates_list.setRowCount(0)
        templates = get_available_templates(TEMPLATE_FOLDER)
        
        for i, template in enumerate(templates):
            self.templates_list.insertRow(i)
            self.templates_list.setItem(i, 0, QTableWidgetItem(template['name']))
            self.templates_list.setItem(i, 1, QTableWidgetItem(template['description']))
    
    def browse_source(self):
        """Browse for a source ELISA datasheet."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select ELISA Datasheet", str(Path.home()),
            "DOCX Files (*.docx);;All Files (*)"
        )
        
        if file_path:
            self.source_path = Path(file_path)
            self.source_path_edit.setText(str(self.source_path))
            
            # If catalog/lot number not manually set, try to extract from filename
            if not self.catalog_number_edit.text() or not self.lot_number_edit.text():
                filename = self.source_path.stem
                if '-' in filename:
                    parts = filename.split('-')
                    if len(parts) >= 2:
                        # Assume format like "EK1586-6058725" (catalog-lot)
                        if not self.catalog_number_edit.text():
                            self.catalog_number_edit.setText(parts[0])
                        if not self.lot_number_edit.text() and len(parts) > 1:
                            self.lot_number_edit.setText(parts[1])
    
    def browse_output(self):
        """Browse for output file location."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Output File", str(OUTPUT_FOLDER),
            "DOCX Files (*.docx);;All Files (*)"
        )
        
        if file_path:
            self.output_path = Path(file_path)
            if not self.output_path.suffix:
                self.output_path = self.output_path.with_suffix('.docx')
            self.output_path_edit.setText(str(self.output_path))
    
    def browse_batch(self):
        """Browse for multiple source ELISA datasheets."""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Select ELISA Datasheets", str(Path.home()),
            "DOCX Files (*.docx);;All Files (*)"
        )
        
        if file_paths:
            self.batch_source_paths = [Path(path) for path in file_paths]
            self.batch_files_label.setText(f"{len(self.batch_source_paths)} files selected")
            
            # Prepare the files table
            self.files_table.setRowCount(len(self.batch_source_paths))
            for i, path in enumerate(self.batch_source_paths):
                self.files_table.setItem(i, 0, QTableWidgetItem(path.name))
                self.files_table.setItem(i, 1, QTableWidgetItem("Queued"))
                
                # Add progress bar
                progress_bar = QProgressBar()
                progress_bar.setValue(0)
                self.files_table.setCellWidget(i, 2, progress_bar)
                
                # Add placeholder for actions
                self.files_table.setItem(i, 3, QTableWidgetItem(""))
    
    def browse_template(self):
        """Browse for a template file to upload."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Template File", str(Path.home()),
            "DOCX Files (*.docx);;All Files (*)"
        )
        
        if file_path:
            self.template_path_edit.setText(file_path)
    
    def upload_template(self):
        """Upload a new template."""
        template_path = self.template_path_edit.text()
        if not template_path:
            QMessageBox.warning(self, "Warning", "Please select a template file first.")
            return
        
        try:
            # Copy the template to the templates folder
            source_path = Path(template_path)
            dest_path = TEMPLATE_FOLDER / source_path.name
            
            # Copy file
            import shutil
            shutil.copy(source_path, dest_path)
            
            # Refresh templates
            self.load_templates()
            self.load_templates(self.batch_template_combo)
            self.populate_templates_list()
            
            QMessageBox.information(self, "Success", f"Template {source_path.name} uploaded successfully.")
            self.template_path_edit.clear()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error uploading template: {str(e)}")
    
    def process_single(self):
        """Process a single ELISA datasheet."""
        if not self.source_path:
            QMessageBox.warning(self, "Warning", "Please select a source ELISA datasheet first.")
            return
        
        # Get template path
        template_name = self.template_combo.currentData()
        template_path = TEMPLATE_FOLDER / template_name
        
        if not template_path.exists():
            QMessageBox.warning(self, "Warning", f"Template {template_name} not found.")
            return
        
        # Determine output path
        if self.output_path:
            output_path = self.output_path
        else:
            # Generate output filename based on catalog and lot numbers if provided
            catalog_number = self.catalog_number_edit.text().strip()
            lot_number = self.lot_number_edit.text().strip()
            
            if catalog_number and lot_number and self.use_metadata_check.isChecked():
                output_filename = f"{catalog_number}-{lot_number}.docx"
            else:
                # Fall back to default naming
                unique_id = os.urandom(4).hex()
                output_filename = f"output_{unique_id}.docx"
            
            output_path = OUTPUT_FOLDER / output_filename
        
        # Get optional user-provided values
        kit_name = self.kit_name_edit.text().strip() or None
        catalog_number = self.catalog_number_edit.text().strip() or None
        lot_number = self.lot_number_edit.text().strip() or None
        
        # Create and start the worker thread
        self.processing_worker = ProcessingWorker(
            self.source_path,
            template_path,
            output_path,
            kit_name=kit_name,
            catalog_number=catalog_number,
            lot_number=lot_number
        )
        
        # Connect signals
        self.processing_worker.progress_signal.connect(self.update_progress)
        self.processing_worker.status_signal.connect(self.update_status)
        self.processing_worker.finished_signal.connect(self.processing_finished)
        
        # Disable the process button during processing
        self.process_btn.setEnabled(False)
        self.process_btn.setText("Processing...")
        
        # Start the worker
        self.processing_worker.start()
    
    def process_batch(self):
        """Process multiple ELISA datasheets."""
        if not self.batch_source_paths:
            QMessageBox.warning(self, "Warning", "Please select source ELISA datasheets first.")
            return
        
        # Get template path
        template_name = self.batch_template_combo.currentData()
        template_path = TEMPLATE_FOLDER / template_name
        
        if not template_path.exists():
            QMessageBox.warning(self, "Warning", f"Template {template_name} not found.")
            return
        
        # Create a batch output directory
        unique_id = os.urandom(4).hex()
        batch_output_dir = BATCH_FOLDER / unique_id
        batch_output_dir.mkdir(exist_ok=True)
        
        # Create and start the worker thread
        self.batch_processing_worker = BatchProcessingWorker(
            self.batch_source_paths,
            template_path,
            batch_output_dir,
            use_metadata=self.batch_use_metadata_check.isChecked()
        )
        
        # Connect signals
        self.batch_processing_worker.overall_progress_signal.connect(self.update_batch_progress)
        self.batch_processing_worker.status_signal.connect(self.update_batch_status)
        self.batch_processing_worker.file_progress_signal.connect(self.update_file_progress)
        self.batch_processing_worker.file_finished_signal.connect(self.file_finished)
        self.batch_processing_worker.all_finished_signal.connect(self.batch_finished)
        
        # Disable the process button during processing
        self.batch_process_btn.setEnabled(False)
        self.batch_process_btn.setText("Processing...")
        
        # Start the worker
        self.batch_processing_worker.start()
    
    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(value)
    
    def update_status(self, message):
        """Update the status label."""
        self.status_label.setText(message)
    
    def update_batch_progress(self, value):
        """Update the batch progress bar."""
        self.batch_progress_bar.setValue(value)
    
    def update_batch_status(self, message):
        """Update the batch status label."""
        self.batch_status_label.setText(message)
    
    def update_file_progress(self, file_index, progress, status):
        """Update progress for a specific file in the batch."""
        # Update status text
        status_item = QTableWidgetItem(status)
        self.files_table.setItem(file_index, 1, status_item)
        
        # Update progress bar
        progress_bar = self.files_table.cellWidget(file_index, 2)
        if progress_bar and isinstance(progress_bar, QProgressBar):
            progress_bar.setValue(progress)
    
    def file_finished(self, file_index, success, output_path):
        """Handle when a file in the batch is finished processing."""
        if success:
            # Add open button for the file
            open_btn = QPushButton("Open")
            open_btn.clicked.connect(lambda: self.open_batch_file(output_path))
            self.files_table.setCellWidget(file_index, 3, open_btn)
    
    def processing_finished(self, success, output_path, error):
        """Handle when processing is finished."""
        self.process_btn.setEnabled(True)
        self.process_btn.setText("Process Datasheet")
        
        if success:
            self.results_label.setText(f"Output saved to: {output_path}")
            self.open_output_btn.setEnabled(True)
            self.output_path = Path(output_path)
            
            QMessageBox.information(self, "Success", f"ELISA datasheet processed successfully. Output saved to: {output_path}")
        else:
            self.results_label.setText(f"Error: {error}")
            self.open_output_btn.setEnabled(False)
            
            QMessageBox.critical(self, "Error", f"Error processing ELISA datasheet: {error}")
    
    def batch_finished(self, success, output_paths):
        """Handle when batch processing is finished."""
        self.batch_process_btn.setEnabled(True)
        self.batch_process_btn.setText("Process Batch")
        
        if success:
            message = f"Batch processing completed successfully. {len(output_paths)} files processed."
            QMessageBox.information(self, "Success", message)
            
            # Ask if the user wants to open the output directory
            reply = QMessageBox.question(
                self, "Open Output Directory", 
                "Do you want to open the output directory?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.open_batch_directory()
        else:
            QMessageBox.warning(self, "Warning", "Batch processing completed with errors. Check the status of individual files.")
    
    def open_output(self):
        """Open the output file."""
        if self.output_path and self.output_path.exists():
            self.open_file(self.output_path)
    
    def open_batch_file(self, path):
        """Open a file from the batch processing."""
        file_path = Path(path)
        if file_path.exists():
            self.open_file(file_path)
    
    def open_batch_directory(self):
        """Open the batch output directory."""
        if self.batch_processing_worker and self.batch_processing_worker.output_folder.exists():
            self.open_directory(self.batch_processing_worker.output_folder)
    
    def open_file(self, path):
        """Open a file with the default application."""
        if platform.system() == 'Windows':
            os.startfile(path)
        elif platform.system() == 'Darwin':  # macOS
            os.system(f'open "{path}"')
        else:  # Linux
            os.system(f'xdg-open "{path}"')
    
    def open_directory(self, path):
        """Open a directory with the default file manager."""
        if platform.system() == 'Windows':
            os.startfile(path)
        elif platform.system() == 'Darwin':  # macOS
            os.system(f'open "{path}"')
        else:  # Linux
            os.system(f'xdg-open "{path}"')


def main():
    """Main entry point for the GUI application."""
    app = QApplication(sys.argv)
    window = ELISAParserGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()