"""
DOCX Template Management
-----------------------
Handles conversion between Jinja text templates and actual DOCX templates.
"""

import logging
import os
import shutil
from pathlib import Path
from typing import List, Dict

import docx
from docxtpl import DocxTemplate

logger = logging.getLogger(__name__)

def create_docx_template_from_text(text_path: Path, output_path: Path) -> bool:
    """
    Create a DOCX template from a text-based template.
    
    Args:
        text_path: Path to the text template
        output_path: Path where the DOCX template will be saved
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create a new document
        doc = docx.Document()
        
        # Read the template text
        with open(text_path, 'r', encoding='utf-8') as f:
            template_text = f.read()
        
        # Add the template text to the document
        doc.add_paragraph(template_text)
        
        # Save the document
        doc.save(output_path)
        
        return True
    except Exception as e:
        logger.exception(f"Error creating DOCX template from text: {e}")
        return False

def get_available_templates(template_dir: Path) -> List[Dict[str, str]]:
    """
    Get a list of available templates in the template directory.
    
    Args:
        template_dir: Path to the template directory
        
    Returns:
        List of dictionaries with template names and descriptions
    """
    templates = []
    
    # List all DOCX files in the template directory
    docx_files = list(template_dir.glob('*.docx'))
    
    for docx_file in docx_files:
        name = docx_file.name
        
        # Skip temporary files
        if name.startswith('~$') or name.startswith('.'):
            continue
        
        # Create a description based on the filename
        if name == 'default_template.docx':
            description = "Default Boster Template"
        elif name == 'innovative_template.docx':
            description = "Innovative Research Template"
        elif name == 'innovative_formatted_template.docx':
            description = "Innovative Research Template (Formatted)"
        else:
            # Remove extension and replace underscores with spaces
            description = name.replace('.docx', '').replace('_', ' ').title()
        
        templates.append({
            'name': name,
            'description': description
        })
    
    return templates

def initialize_templates(template_dir: Path, assets_dir: Path) -> None:
    """
    Initialize the template directory with default templates.
    
    Args:
        template_dir: Path to the template directory
        assets_dir: Path to the assets directory containing source templates
    """
    # Create the template directory if it doesn't exist
    template_dir.mkdir(exist_ok=True)
    
    # Copy default templates if they don't exist
    templates = [
        ('boster_template_ready.docx', 'default_template.docx'),
        ('IMSKLK1KT-Sample.docx', 'innovative_template.docx')
    ]
    
    for source_name, dest_name in templates:
        source_path = assets_dir / source_name
        dest_path = template_dir / dest_name
        
        if source_path.exists() and not dest_path.exists():
            shutil.copy(source_path, dest_path)
            logger.info(f"Copied {source_name} to {dest_name}")
    
    # Convert text-based templates to DOCX if needed
    text_templates = list(template_dir.glob('*.jinja.docx'))
    for text_path in text_templates:
        # Create output path by removing .jinja from the filename
        output_name = text_path.name.replace('.jinja.docx', '.docx')
        output_path = template_dir / output_name
        
        # Only create if the output doesn't exist or is older than the text template
        if not output_path.exists() or (output_path.stat().st_mtime < text_path.stat().st_mtime):
            if create_docx_template_from_text(text_path, output_path):
                logger.info(f"Created DOCX template {output_name} from {text_path.name}")

def get_template_path(template_dir: Path, template_name: str) -> Path:
    """
    Get the path to a template.
    
    Args:
        template_dir: Path to the template directory
        template_name: Name of the template
        
    Returns:
        Path to the template
    """
    return template_dir / template_name