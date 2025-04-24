#!/usr/bin/env python3
"""
Fix how materials are added to the template by directly modifying the template population logic
to not duplicate bullets.
"""

import logging
import re
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def fix_template_populator():
    """Fix the template populator to correctly handle material bullet points."""
    template_file = "template_populator_enhanced.py"
    with open(template_file, 'r') as f:
        content = f.read()
    
    # Find the section that populates material entries
    pattern = r"(# Add individual material entries\s+for i in range\(min\(len\(req_materials\), 10\)\):\s+processed_data\[f'req_material_\{i\+1\}'\] = req_materials\[i\])"
    
    # New version that prepends a bullet point character and removes existing bullets
    replacement = """# Add individual material entries
                for i in range(min(len(req_materials), 10)):
                    # Clean up the material text (remove existing bullets)
                    material_text = req_materials[i]
                    material_text = material_text.strip()
                    # Remove existing bullet characters
                    if material_text.startswith('•'):
                        material_text = material_text[1:].strip()
                    processed_data[f'req_material_{i+1}'] = material_text"""
    
    # Replace the pattern
    new_content = re.sub(pattern, replacement, content)
    
    # Write it back
    with open(template_file, 'w') as f:
        f.write(new_content)
    
    logger.info(f"Updated {template_file} to better handle material bullets")
    
    # Now fix the fix_bullet_points.py script
    bullet_file = "fix_bullet_points.py"
    with open(bullet_file, 'r') as f:
        content = f.read()
    
    # Find the function that adds bullets
    pattern = r"def add_bullet_to_paragraph\(paragraph\):\s+\"\"\"Add a bullet character to the start of a paragraph.\"\"\"\s+run = paragraph.add_run\(\"• \", 0\)  # Insert at the beginning\s+run.font.size = Pt\(11\)"
    
    # Replace with a better version
    replacement = """def add_bullet_to_paragraph(paragraph):
    \"\"\"Add a bullet character to the start of a paragraph.\"\"\"
    # First remove existing runs
    for _ in range(len(paragraph.runs)):
        paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)
    
    # Now add just the bullet
    run = paragraph.add_run("• ")
    run.font.size = Pt(11)"""
    
    # Replace the pattern
    new_content = re.sub(pattern, replacement, content)
    
    # Write it back
    with open(bullet_file, 'w') as f:
        f.write(new_content)
    
    logger.info(f"Updated {bullet_file} to better handle material bullets")
    
    return True

if __name__ == "__main__":
    fix_template_populator()