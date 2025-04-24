"""
Template Populator
-----------------
Populates DOCX templates with extracted ELISA datasheet data.
"""

import logging
import re
from pathlib import Path
from typing import Dict, Any, Optional, List

from docxtpl import DocxTemplate

logger = logging.getLogger(__name__)

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
        
    def _clean_data(self, data: Dict[str, Any], kit_name: Optional[str] = None, 
                   catalog_number: Optional[str] = None, lot_number: Optional[str] = None) -> Dict[str, Any]:
        """
        Clean and prepare data for template population.
        
        Args:
            data: Raw extracted data dictionary
            kit_name: Optional kit name provided by user
            catalog_number: Optional catalog number provided by user
            lot_number: Optional lot number provided by user
            
        Returns:
            Processed data dictionary ready for template population
        """
        # Create a copy of the data to avoid modifying the original
        processed_data = data.copy()
        
        # Override with user-provided values if available
        if kit_name:
            processed_data['kit_name'] = kit_name
        elif 'catalog_number' in processed_data:
            # Try to construct a kit name from existing data
            catalog = processed_data.get('catalog_number', '')
            if catalog and 'description' in processed_data:
                description = processed_data.get('description', '')
                kit_name_match = re.search(r'(Mouse|Rat|Human|Canine|Bovine|Porcine)\s+([A-Za-z0-9]+)', description)
                if kit_name_match:
                    processed_data['kit_name'] = f"{kit_name_match.group(0)} ELISA Kit"
                else:
                    processed_data['kit_name'] = f"ELISA Kit ({catalog})"
        
        if catalog_number:
            processed_data['catalog_number'] = catalog_number
            
        if lot_number:
            processed_data['lot_number'] = lot_number
        else:
            processed_data['lot_number'] = "LOT#_______"  # Placeholder for user to fill manually
            
        # Extract intended use from assay principle (first paragraph)
        if 'assay_principle' in processed_data:
            assay_principle = processed_data['assay_principle']
            # Split by paragraph breaks and take the first paragraph
            paragraphs = assay_principle.split('\n\n')
            if paragraphs:
                processed_data['intended_use'] = paragraphs[0].strip()
                # Use the rest of paragraphs (minus the last sentence) for principle_of_assay
                if len(paragraphs) > 1:
                    principle_text = paragraphs[1].strip()
                    # Remove the last sentence if it contains Boster reference
                    sentences = re.split(r'(?<=[.!?])\s+', principle_text)
                    if sentences and any(word in sentences[-1].lower() for word in ['boster', 'picokine']):
                        principle_text = ' '.join(sentences[:-1])
                    processed_data['principle_of_assay'] = principle_text
        
        # Process background section
        if 'background' in processed_data:
            background_text = processed_data['background']
            # If user provided a kit name, use it to create a background title
            if kit_name:
                # Extract key identifier from kit name (e.g., "KLK1" from "Mouse KLK1 ELISA Kit")
                identifier_match = re.search(r'(Mouse|Rat|Human|Canine|Bovine|Porcine)\s+([A-Za-z0-9]+)', kit_name)
                if identifier_match:
                    identifier = identifier_match.group(2)
                    processed_data['background_title'] = f"Background on {identifier}"
            else:
                processed_data['background_title'] = "Background"
            
            processed_data['background_text'] = background_text
            
        # Process standard curve section
        if 'standard_curve' in processed_data:
            standard_curve = processed_data['standard_curve']
            
            # Extract product name for standard curve title
            if kit_name:
                # Extract product identifier (e.g., "Mouse KLK1" from "Mouse KLK1 ELISA Kit")
                product_match = re.search(r'(Mouse|Rat|Human|Canine|Bovine|Porcine)\s+([A-Za-z0-9]+)', kit_name)
                if product_match:
                    product_id = product_match.group(0)
                    processed_data['standard_curve_title'] = f"{product_id} ELISA Standard Curve Example"
                else:
                    processed_data['standard_curve_title'] = "ELISA Standard Curve Example"
            else:
                processed_data['standard_curve_title'] = "ELISA Standard Curve Example"
                
            # Ensure standard curve concentrations and OD values are properly formatted
            if 'concentrations' in standard_curve and 'od_values' in standard_curve:
                # Create a formatted table for the template
                std_curve_table = []
                for i, (conc, od) in enumerate(zip(
                    standard_curve['concentrations'], 
                    standard_curve['od_values']
                )):
                    # Make sure first concentration is 0.0
                    if i == 0 and conc != "0.0":
                        std_curve_table.append({
                            'concentration': "0.0",
                            'od_value': od
                        })
                    else:
                        std_curve_table.append({
                            'concentration': conc,
                            'od_value': od
                        })
                
                processed_data['standard_curve_table'] = std_curve_table
        
        # Process data analysis section - remove Boster reference
        if 'data_analysis' in processed_data:
            data_analysis = processed_data['data_analysis']
            # Remove first two sentences if they contain Boster references
            sentences = re.split(r'(?<=[.!?])\s+', data_analysis)
            if len(sentences) > 2 and any(word in ' '.join(sentences[:2]).lower() for word in ['boster', 'biocompare', 'online']):
                processed_data['data_analysis'] = ' '.join(sentences[2:])
            else:
                processed_data['data_analysis'] = data_analysis
                
        # Replace "Boster" with "Innovative Research" in all text fields
        for key, value in processed_data.items():
            if isinstance(value, str):
                # Replace "Boster" with "Innovative Research"
                value = re.sub(r'\bBoster\b', 'Innovative Research', value)
                # Remove all variations of PicoKine®
                value = re.sub(r'PicoKine\s*®', '', value)
                value = re.sub(r'PicoKine', '', value)
                processed_data[key] = value
            elif isinstance(value, list) and all(isinstance(item, dict) for item in value):
                # Handle lists of dictionaries (like reagents, tables, etc.)
                for item in value:
                    for item_key, item_value in item.items():
                        if isinstance(item_value, str):
                            # Apply the same replacements to dictionary values
                            item[item_key] = (re.sub(r'\bBoster\b', 'Innovative Research', 
                                             re.sub(r'PicoKine\s*®', '', 
                                             re.sub(r'PicoKine', '', item_value))))
        
        return processed_data
        
    def populate(self, data: Dict[str, Any], output_path: Path, 
                kit_name: Optional[str] = None, 
                catalog_number: Optional[str] = None, 
                lot_number: Optional[str] = None) -> None:
        """
        Populate the template with the extracted data and save to the output path.
        
        Args:
            data: Dictionary containing structured data to populate the template
            output_path: Path where the populated template will be saved
            kit_name: Optional kit name provided by user
            catalog_number: Optional catalog number provided by user
            lot_number: Optional lot number provided by user
        """
        self.logger.info(f"Populating template {self.template_path} with extracted data")
        
        try:
            # Clean and prepare the data
            processed_data = self._clean_data(data, kit_name, catalog_number, lot_number)
            
            # Render the template with the processed data
            self.template.render(processed_data)
            
            # Save the populated template
            self.template.save(output_path)
            
            self.logger.info(f"Template successfully populated and saved to {output_path}")
            
        except Exception as e:
            self.logger.exception(f"Error populating template: {e}")
            raise
