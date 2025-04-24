"""
Template Populator
-----------------
Populates DOCX templates with extracted ELISA datasheet data.
"""

import logging
from pathlib import Path
from typing import Dict, Any

from docxtpl import DocxTemplate

class TemplatePopulator:
    """
    Populates DOCX templates with extracted ELISA datasheet data.
    
    Uses the docxtpl library to fill templates with the structured data
    extracted from ELISA kit datasheets.
    """
    
    def __init__(self, template_path: Path):
        """
        Initialize the template populator with the path to the template.
        
        Args:
            template_path: Path to the DOCX template file
        """
        self.template_path = template_path
        self.logger = logging.getLogger(__name__)
        self.template = DocxTemplate(template_path)
        
    def populate(self, data: Dict[str, Any], output_path: Path) -> None:
        """
        Populate the template with the extracted data and save to the output path.
        
        Args:
            data: Dictionary containing structured data to populate the template
            output_path: Path where the populated template will be saved
        """
        self.logger.info(f"Populating template {self.template_path} with extracted data")
        
        try:
            # Render the template with the data
            self.template.render(data)
            
            # Save the populated template
            self.template.save(output_path)
            
            self.logger.info(f"Template successfully populated and saved to {output_path}")
            
        except Exception as e:
            self.logger.exception(f"Error populating template: {e}")
            raise
