#!/usr/bin/env python3
"""
ELISA Kit Datasheet Web Application
-----------------------------------
Web interface for extracting data from ELISA kit datasheets and populating DOCX templates.
"""

import os
import uuid
import json
import zipfile
import logging
import threading
import hashlib
from pathlib import Path
from functools import wraps
from typing import Dict, List, Any
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, send_from_directory, jsonify, session

from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator
from docx_templates import initialize_templates, get_available_templates
from batch_processor import BatchProcessor

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create the Flask application
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")

# Set the application password (in a real app, this would be stored securely, not hardcoded)
# The hashed version of "ElisaParser2025!" - in production, this should come from environment variables
APP_PASSWORD_HASH = "3b8fc838840530f5acee33eeef31785c41fc9502"

# Create upload folders if they don't exist
UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
TEMPLATE_FOLDER = Path('templates_docx')
ASSETS_FOLDER = Path('attached_assets')
BATCH_FOLDER = Path('batch_outputs')
DEFAULT_TEMPLATE = TEMPLATE_FOLDER / 'enhanced_template.docx'

# Store batch processing tasks
batch_tasks = {}

for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_FOLDER, BATCH_FOLDER]:
    folder.mkdir(exist_ok=True)

# Initialize templates
initialize_templates(TEMPLATE_FOLDER, ASSETS_FOLDER)

# Make sure the enhanced template is the default
if not DEFAULT_TEMPLATE.exists():
    logger.warning(f"Default enhanced template not found at {DEFAULT_TEMPLATE}")
    logger.info("Looking for any available template to use as default...")
    # Find any template to use as a fallback
    templates = list(TEMPLATE_FOLDER.glob('*.docx'))
    if templates:
        DEFAULT_TEMPLATE = templates[0]
        logger.info(f"Using {DEFAULT_TEMPLATE.name} as the default template")
    else:
        logger.warning("No templates found. The application may not work correctly.")

# Define a simple login required decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('authenticated'):
            flash('Please log in to access this page.', 'info')
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Handle password protection"""
    if session.get('authenticated'):
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        password = request.form.get('password')
        remember_me = 'remember_me' in request.form
        
        # Hash the input password using SHA-1 for comparison
        password_hash = hashlib.sha1(password.encode()).hexdigest()
        
        if password_hash == APP_PASSWORD_HASH:
            session['authenticated'] = True
            if remember_me:
                # Set a longer session lifetime (30 days)
                session.permanent = True
            
            next_page = request.args.get('next')
            if not next_page or not next_page.startswith('/'):
                next_page = url_for('index')
                
            flash('Login successful!', 'success')
            return redirect(next_page)
        else:
            flash('Invalid password. Please try again.', 'error')
            return redirect(url_for('login'))
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Handle logout"""
    session.pop('authenticated', None)
    flash('Logged out successfully.', 'success')
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    """Render the home page"""
    # Get available templates with descriptions
    templates = get_available_templates(TEMPLATE_FOLDER)
    
    # Mark the default enhanced template
    default_template_name = DEFAULT_TEMPLATE.name if DEFAULT_TEMPLATE.exists() else None
    for template in templates:
        if template['name'] == default_template_name:
            template['is_default'] = "yes"
            template['description'] += " (Default)"
        else:
            template['is_default'] = "no"
    
    # List recent outputs if any
    recent_outputs = list(OUTPUT_FOLDER.glob('*.docx'))
    recent_outputs = sorted(recent_outputs, key=lambda x: x.stat().st_mtime, reverse=True)[:5]
    recent_output_names = [output.name for output in recent_outputs]
    
    return render_template('index.html', templates=templates, recent_outputs=recent_output_names, default_template=default_template_name)

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    """Handle file upload and processing"""
    if 'source_file' not in request.files:
        flash('No file part', 'error')
        return redirect(request.url)
    
    source_file = request.files['source_file']
    if source_file.filename == '':
        flash('No selected file', 'error')
        return redirect(request.url)
    
    if not source_file.filename.lower().endswith('.docx'):
        flash('Only DOCX files are supported', 'error')
        return redirect(request.url)
    
    # Get selected template or use enhanced template as default
    template_name = request.form.get('template', 'enhanced_template.docx')
    
    if template_name:
        template_path = TEMPLATE_FOLDER / template_name
        if not template_path.exists():
            logger.warning(f"Selected template {template_name} not found, using default")
            template_path = DEFAULT_TEMPLATE
    else:
        # No template selected, use enhanced template
        template_path = DEFAULT_TEMPLATE
    
    # Log which template is being used
    logger.info(f"Using template: {template_path.name}")
        
    if not template_path.exists():
        flash(f'Template not found. Please upload a template first.', 'error')
        return redirect(request.url)
    
    try:
        # Save the uploaded file
        unique_id = str(uuid.uuid4())[:8]
        source_filename = f"{unique_id}_{source_file.filename}"
        source_path = UPLOAD_FOLDER / source_filename
        source_file.save(source_path)
        
        # Get catalog number and lot number for filename
        catalog_number = request.form.get('catalog_number', '').strip()
        lot_number = request.form.get('lot_number', '').strip()
        
        # Generate output filename based on catalog and lot numbers if provided
        if catalog_number and lot_number:
            output_filename = f"{catalog_number}-{lot_number}.docx"
        else:
            # Fall back to default naming if either is missing
            output_filename = f"output_{unique_id}.docx"
            
        output_path = OUTPUT_FOLDER / output_filename
        
        # Get optional user-provided values
        kit_name = request.form.get('kit_name')
        catalog_number = request.form.get('catalog_number')
        lot_number = request.form.get('lot_number')
        
        # Process the file
        parser = ELISADatasheetParser(source_path)
        data = parser.extract_data()
        
        # Populate template with user-provided values
        populator = TemplatePopulator(template_path)
        populator.populate(
            data, 
            output_path,
            kit_name=kit_name,
            catalog_number=catalog_number,
            lot_number=lot_number
        )
        
        # Redirect to download page
        return redirect(url_for('download_file', filename=output_filename))
    
    except Exception as e:
        logger.exception(f"Error processing file: {e}")
        flash(f"Error processing file: {str(e)}", 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
@login_required
def download_file(filename):
    """Download a processed file"""
    try:
        # Make sure the file name only contains safe characters
        safe_filename = os.path.basename(filename)
        output_path = OUTPUT_FOLDER / safe_filename
        
        # Additional check to ensure file exists and is accessible
        if not output_path.exists():
            logger.error(f'File {safe_filename} not found at {output_path}')
            flash(f'File {safe_filename} not found', 'error')
            return redirect(url_for('index'))
        
        logger.info(f'Sending file: {output_path}, size: {output_path.stat().st_size} bytes')
        
        # Use send_from_directory with more explicit parameters
        return send_from_directory(
            directory=str(OUTPUT_FOLDER),
            path=safe_filename,
            as_attachment=True,
            download_name=f"ELISA_Kit_Datasheet_{safe_filename}"
        )
    except Exception as e:
        logger.exception(f"Error downloading file: {e}")
        flash(f"Error downloading file: {str(e)}", 'error')
        return redirect(url_for('index'))

@app.route('/upload_template', methods=['POST'])
@login_required
def upload_template():
    """Handle template upload"""
    if 'template_file' not in request.files:
        flash('No file part', 'error')
        return redirect(url_for('index'))
    
    template_file = request.files['template_file']
    if template_file.filename == '':
        flash('No selected file', 'error')
        return redirect(url_for('index'))
    
    if not template_file.filename.lower().endswith('.docx'):
        flash('Only DOCX files are supported', 'error')
        return redirect(url_for('index'))
    
    try:
        # Save the uploaded template
        template_path = TEMPLATE_FOLDER / template_file.filename
        template_file.save(template_path)
        flash(f'Template {template_file.filename} uploaded successfully', 'success')
    except Exception as e:
        logger.exception(f"Error uploading template: {e}")
        flash(f"Error uploading template: {str(e)}", 'error')
    
    return redirect(url_for('index'))

@app.route('/view_source')
@login_required
def view_source():
    """View the source structure page"""
    # Extract structure from default source file to show as an example
    try:
        source_path = Path('attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx')
        parser = ELISADatasheetParser(source_path)
        data = parser.extract_data()
        
        # Convert data to a more readable format for display
        readable_data = {}
        for key, value in data.items():
            if isinstance(value, dict):
                readable_data[key] = value
            elif isinstance(value, list):
                if all(isinstance(item, dict) for item in value):
                    readable_data[key] = value
                else:
                    readable_data[key] = ", ".join(str(item) for item in value)
            else:
                # Truncate long text values for display
                if isinstance(value, str) and len(value) > 100:
                    readable_data[key] = value[:100] + "..."
                else:
                    readable_data[key] = value
        
        return render_template('view_source.html', data=readable_data)
    except Exception as e:
        logger.exception(f"Error viewing source structure: {e}")
        flash(f"Error viewing source structure: {str(e)}", 'error')
        return redirect(url_for('index'))

@app.route('/batch_process')
@login_required
def batch_process():
    """Show batch processing page"""
    # Get available templates with descriptions
    templates = get_available_templates(TEMPLATE_FOLDER)
    
    # Mark the default enhanced template
    default_template_name = DEFAULT_TEMPLATE.name if DEFAULT_TEMPLATE.exists() else None
    for template in templates:
        if template['name'] == default_template_name:
            template['is_default'] = "yes"
            template['description'] += " (Default)"
        else:
            template['is_default'] = "no"
    
    # List available source files
    source_files = list(UPLOAD_FOLDER.glob('*.docx'))
    source_file_names = [source.name for source in source_files]
    
    return render_template('batch_process.html', templates=templates, source_files=source_file_names, default_template=default_template_name)

@app.route('/about')
@login_required
def about():
    """Show about page with information about the application"""
    return render_template('about.html')

@app.route('/upload_batch', methods=['POST'])
@login_required
def upload_batch():
    """Handle batch file upload and processing"""
    if 'source_files' not in request.files:
        flash('No files found', 'error')
        return redirect(url_for('batch_process'))
    
    files = request.files.getlist('source_files')
    if not files or (len(files) == 1 and files[0].filename == ''):
        flash('No files selected', 'error')
        return redirect(url_for('batch_process'))
    
    # Get selected template or use enhanced template as default
    template_name = request.form.get('template', 'enhanced_template.docx')
    
    if template_name:
        template_path = TEMPLATE_FOLDER / template_name
        if not template_path.exists():
            logger.warning(f"Selected template {template_name} not found, using default")
            template_path = DEFAULT_TEMPLATE
    else:
        # No template selected, use enhanced template
        template_path = DEFAULT_TEMPLATE
    
    # Log which template is being used
    logger.info(f"Using template: {template_path.name} for batch processing")
    
    if not template_path.exists():
        flash(f'Template not found. Please upload a template first.', 'error')
        return redirect(url_for('batch_process'))
    
    # Process in parallel if requested
    process_parallel = 'process_parallel' in request.form
    use_metadata = 'use_metadata' in request.form
    
    # Create a unique task ID
    task_id = str(uuid.uuid4())
    batch_output_dir = BATCH_FOLDER / task_id
    batch_output_dir.mkdir(exist_ok=True)
    
    # Save the files
    file_paths = []
    for file in files:
        if file.filename.lower().endswith('.docx'):
            # Save the file with a unique prefix
            unique_id = str(uuid.uuid4())[:8]
            filename = f"{unique_id}_{file.filename}"
            file_path = UPLOAD_FOLDER / filename
            file.save(file_path)
            file_paths.append(file_path)
    
    # Create a batch processor
    processor = BatchProcessor(template_path, batch_output_dir)
    
    # Start the batch processing in a background thread
    def process_files_async():
        try:
            if process_parallel:
                results = processor.process_batch_parallel(file_paths)
            else:
                results = processor.process_batch(file_paths)
            
            # Store the results
            batch_tasks[task_id] = {
                'status': 'completed',
                'template': template_path.name,
                'total': len(file_paths),
                'successful': results['successful'],
                'failed': results['failed'],
                'files': results['files'],
                'output_dir': str(batch_output_dir)
            }
            
            # Create a ZIP file with all the outputs if there are successful results
            if results['successful'] > 0:
                zip_path = BATCH_FOLDER / f"{task_id}.zip"
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for output_file in batch_output_dir.glob('*.docx'):
                        zipf.write(output_file, arcname=output_file.name)
                
                batch_tasks[task_id]['zip_path'] = str(zip_path)
        
        except Exception as e:
            logger.exception(f"Error processing batch: {e}")
            batch_tasks[task_id]['status'] = 'failed'
            batch_tasks[task_id]['error'] = str(e)
    
    # Initialize the task status
    batch_tasks[task_id] = {
        'status': 'processing',
        'template': template_path.name,
        'total': len(file_paths),
        'successful': 0,
        'failed': 0,
        'files': []
    }
    
    # Start processing in the background
    thread = threading.Thread(target=process_files_async)
    thread.daemon = True
    thread.start()
    
    # Return the task ID for status checking
    return jsonify({'task_id': task_id})

@app.route('/batch_status/<task_id>')
@login_required
def batch_status(task_id):
    """Get the status of a batch processing task"""
    if task_id not in batch_tasks:
        return jsonify({'status': 'not_found'})
    
    task = batch_tasks[task_id]
    
    # Add progress information from the processor if available
    processor_progress = batch_tasks.get(task_id, {}).get('processor', None)
    if processor_progress:
        task['progress'] = processor_progress.get_progress()
    
    return jsonify(task)

@app.route('/download_batch/<task_id>')
@login_required
def download_batch(task_id):
    """Download a ZIP file containing all batch outputs"""
    if task_id not in batch_tasks:
        flash('Batch task not found', 'error')
        return redirect(url_for('batch_process'))
    
    task = batch_tasks[task_id]
    if 'zip_path' not in task:
        flash('No ZIP file found for this batch', 'error')
        return redirect(url_for('batch_process'))
    
    try:
        zip_path = Path(task['zip_path'])
        if not zip_path.exists():
            flash('ZIP file not found', 'error')
            return redirect(url_for('batch_process'))
        
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=f"ELISA_Kit_Datasheets_Batch_{task_id[:8]}.zip"
        )
    
    except Exception as e:
        logger.exception(f"Error downloading batch: {e}")
        flash(f"Error downloading batch: {str(e)}", 'error')
        return redirect(url_for('batch_process'))

@app.route('/api/templates')
@login_required
def api_templates():
    """API to get available templates"""
    templates = get_available_templates(TEMPLATE_FOLDER)
    return jsonify({'templates': templates})

@app.route('/api/recent_outputs')
@login_required
def api_recent_outputs():
    """API to get recent outputs"""
    recent_outputs = list(OUTPUT_FOLDER.glob('*.docx'))
    recent_outputs = sorted(recent_outputs, key=lambda x: x.stat().st_mtime, reverse=True)[:10]
    output_list = [{'name': output.name, 'size': output.stat().st_size, 'date': output.stat().st_mtime} for output in recent_outputs]
    return jsonify({'outputs': output_list})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)