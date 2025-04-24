#!/usr/bin/env python3
"""
Fix how materials are added to the template by directly modifying the template population logic
to not duplicate bullets.
"""

import logging
import re
from pathlib import Path
from docx import Document

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def fix_template_populator():
    """Fix the template populator to correctly handle material bullet points."""
    target_file = "template_populator_enhanced.py"
    original_content = Path(target_file).read_text()
    
    # Find the section that processes required materials
    materials_pattern = r"# Map required materials to individual bullet points.*?# Map standard curve data"
    materials_section = re.search(materials_pattern, original_content, re.DOTALL)
    
    if not materials_section:
        logger.error("Could not find materials section in template populator")
        return False
    
    # Create a replacement that properly handles material bullets
    replacement = '''
            # Map required materials to individual bullet points
            if 'required_materials' in processed_data:
                req_materials = processed_data['required_materials']
                self.logger.info(f"Processing {len(req_materials)} required materials for template")
                
                # Ensure we have enough materials
                if len(req_materials) < 5:
                    # Add default items if needed
                    default_items = [
                        "Microplate reader capable of measuring absorbance at 450 nm",
                        "Automated plate washer (optional)",
                        "Adjustable pipettes and pipette tips capable of precisely dispensing volumes",
                        "Tubes for sample preparation",
                        "Deionized or distilled water"
                    ]
                    
                    for item in default_items:
                        if item not in req_materials:
                            req_materials.append(item)
                            if len(req_materials) >= 5:
                                break
                
                # Clean up the material items (only keep the text, no bullets)
                clean_materials = []
                for material in req_materials:
                    # Remove any existing bullet points
                    material_text = material.strip()
                    if material_text.startswith('•') or material_text.startswith('-'):
                        material_text = material_text[1:].strip()
                    elif material_text.startswith('\\u2022'):  # Unicode bullet
                        material_text = material_text[1:].strip()
                    
                    # Only add non-empty materials
                    if material_text:
                        clean_materials.append(material_text)
                
                # Add individual material entries (WITHOUT bullets, template already has them)
                for i in range(min(len(clean_materials), 10)):
                    processed_data[f'req_material_{i+1}'] = clean_materials[i]
                
                # Clear any unused material slots
                for i in range(len(clean_materials) + 1, 11):
                    processed_data[f'req_material_{i}'] = ''
            
            # Map standard curve data'''
    
    # Replace the old section with the new one
    modified_content = original_content.replace(materials_section.group(0), replacement)
    
    # Write the modified content back to the file
    Path(target_file).write_text(modified_content)
    logger.info(f"Updated template populator in {target_file}")
    
    return True

if __name__ == "__main__":
    fix_template_populator()
    print("Fixed template populator to correctly handle material bullet points.")