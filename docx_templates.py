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
    template_order = []
    
    # Pre-define template descriptions and their order
    template_info = {
        'enhanced_template.docx': "Enhanced Template with Fixed Tables and Native Word Formatting",
        'default_template.docx': "Default Boster Template",
        'boster_template_ready.docx': "Boster Template with Standard Formatting",
        'innovative_template.docx': "Innovative Research Template",
        'innovative_formatted_template.docx': "Innovative Research Template (Formatted)",
        'innovative_direct_template.docx': "Innovative Research Direct Format Template",
        'innovative_proper_template.docx': "Innovative Research Proper Template",
        'innovative_exact_template.docx': "Innovative Research Exact Format Template",
        'red_dot_template.docx': "Innovative Research Template"
    }
    
    # Set the order of templates to display
    template_order = [
        'enhanced_template.docx',  # Put enhanced template first 
        'boster_template_ready.docx',  # Second choice
        'red_dot_template.docx',   # Third choice
        'innovative_exact_template.docx',  # Fourth choice
        'default_template.docx',
        'innovative_template.docx',
        'innovative_formatted_template.docx',
        'innovative_direct_template.docx',
        'innovative_proper_template.docx'
    ]
    
    # First add templates in the defined order (if they exist)
    for template_name in template_order:
        template_path = template_dir / template_name
        if template_path.exists():
            templates.append({
                'name': template_name,
                'description': template_info.get(template_name, template_name.replace('.docx', '').replace('_', ' ').title())
            })
    
    # Then add any other templates that might be in the directory but not in our ordered list
    docx_files = list(template_dir.glob('*.docx'))
    for docx_file in docx_files:
        name = docx_file.name
        
        # Skip temporary files and already added templates
        if name.startswith('~$') or name.startswith('.') or name in [t['name'] for t in templates]:
            continue
        
        # Create a description based on the filename or use from template_info if available
        description = template_info.get(name, name.replace('.docx', '').replace('_', ' ').title())
        
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
    
    # First try to find and use the enhanced template with proper date stamp
    enhanced_template_exists = False
    
    # Check for todays or other dated Innovative template
    enhance_path = Path('IMSKLK1KT-20250424.docx')
    dest_path = template_dir / 'enhanced_template.docx'
    output_template = Path('output_populated_template.docx')
    
    # Try to find any other dated template if today's doesn't exist
    if not enhance_path.exists():
        # Look for any recent IMSKLK1KT-*.docx file
        dated_templates = list(Path('.').glob('IMSKLK1KT-*.docx'))
        if dated_templates:
            enhance_path = sorted(dated_templates, key=lambda x: x.stat().st_mtime, reverse=True)[0]
            logger.info(f"Found alternative enhanced template: {enhance_path}")
    
    # Copy the enhanced template if found and it doesn't already exist
    if enhance_path.exists() and not dest_path.exists():
        shutil.copy(enhance_path, dest_path)
        logger.info(f"Copied enhanced template from {enhance_path} to {dest_path}")
        enhanced_template_exists = True
    elif dest_path.exists():
        enhanced_template_exists = True
        logger.info(f"Enhanced template already exists at {dest_path}")
    elif output_template.exists():
        # If no enhanced template but we have a recently generated output, use that
        shutil.copy(output_template, dest_path)
        logger.info(f"Copied output template to {dest_path}")
        enhanced_template_exists = True
    
    # Copy other default templates if they don't exist
    templates = [
        ('boster_template_ready.docx', 'default_template.docx'),
        ('IMSKLK1KT-Sample.docx', 'innovative_template.docx'),
        ('RDR-LMNB2-Hu.docx', 'red_dot_template.docx'),
    ]
    
    # If enhanced template still doesn't exist, add it to the list
    if not enhanced_template_exists:
        templates.append(('boster_template_ready.docx', 'enhanced_template.docx'))
        logger.warning("Could not find an enhanced template, will use boster_template_ready.docx as fallback")
    
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