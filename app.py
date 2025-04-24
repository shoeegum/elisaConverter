#!/usr/bin/env python3
"""
ELISA Kit Datasheet Web Application
-----------------------------------
Web interface for extracting data from ELISA kit datasheets and populating DOCX templates.
"""

import os
import uuid
import logging
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, flash, send_file

from elisa_parser import ELISADatasheetParser
from template_populator_enhanced import TemplatePopulator
from docx_templates import initialize_templates, get_available_templates

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create the Flask application
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")

# Create upload folders if they don't exist
UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
TEMPLATE_FOLDER = Path('templates_docx')
ASSETS_FOLDER = Path('attached_assets')
DEFAULT_TEMPLATE = TEMPLATE_FOLDER / 'enhanced_template.docx'

for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_FOLDER]:
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

@app.route('/')
def index():
    """Render the home page"""
    # Get available templates with descriptions
    templates = get_available_templates(TEMPLATE_FOLDER)
    
    # Mark the default enhanced template
    default_template_name = DEFAULT_TEMPLATE.name if DEFAULT_TEMPLATE.exists() else None
    for template in templates:
        if template['name'] == default_template_name:
            template['is_default'] = True
            template['description'] += " (Default)"
        else:
            template['is_default'] = False
    
    # List recent outputs if any
    recent_outputs = list(OUTPUT_FOLDER.glob('*.docx'))
    recent_outputs = sorted(recent_outputs, key=lambda x: x.stat().st_mtime, reverse=True)[:5]
    recent_output_names = [output.name for output in recent_outputs]
    
    return render_template('index.html', templates=templates, recent_outputs=recent_output_names, default_template=default_template_name)

@app.route('/upload', methods=['POST'])
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
    template_name = request.form.get('template')
    
    if template_name:
        template_path = TEMPLATE_FOLDER / template_name
        if not template_path.exists():
            logger.warning(f"Selected template {template_name} not found, using default")
            template_path = DEFAULT_TEMPLATE
    else:
        # No template selected, use enhanced template
        template_path = DEFAULT_TEMPLATE
        
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
def download_file(filename):
    """Download a processed file"""
    output_path = OUTPUT_FOLDER / filename
    
    if not output_path.exists():
        flash(f'File {filename} not found', 'error')
        return redirect(url_for('index'))
    
    return send_file(output_path, as_attachment=True)

@app.route('/upload_template', methods=['POST'])
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
def batch_process():
    """Show batch processing page"""
    # Get available templates with descriptions
    templates = get_available_templates(TEMPLATE_FOLDER)
    
    # Mark the default enhanced template
    default_template_name = DEFAULT_TEMPLATE.name if DEFAULT_TEMPLATE.exists() else None
    for template in templates:
        if template['name'] == default_template_name:
            template['is_default'] = True
            template['description'] += " (Default)"
        else:
            template['is_default'] = False
    
    # List available source files
    source_files = list(UPLOAD_FOLDER.glob('*.docx'))
    source_file_names = [source.name for source in source_files]
    
    return render_template('batch_process.html', templates=templates, source_files=source_file_names, default_template=default_template_name)

@app.route('/about')
def about():
    """Show about page with information about the application"""
    return render_template('about.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)