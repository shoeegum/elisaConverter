"""
Template Populator Enhanced
-----------------
Enhanced version of the template populator to handle the new technical details table format.
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
            
        # Extract and ensure intended use is populated
        if not processed_data.get('intended_use') or processed_data.get('intended_use') == "For research use only. Not for use in diagnostic procedures.":
            # First check if assay_principle exists to extract from there
            if 'assay_principle' in processed_data and processed_data['assay_principle']:
                assay_principle = processed_data['assay_principle']
                
                # Try different splitting patterns to find the first paragraph
                # First try splitting by double newlines
                paragraphs = assay_principle.split('\n\n')
                
                if paragraphs:
                    processed_data['intended_use'] = paragraphs[0].strip()
                    
                    # If split didn't work (whole text in one paragraph), try to get the first sentence
                    if len(paragraphs) == 1 and len(paragraphs[0].split('.')) > 1:
                        first_sentence = paragraphs[0].split('.')[0].strip() + '.'
                        if len(first_sentence) > 20:  # Make sure it's substantive
                            processed_data['intended_use'] = first_sentence
                
                # Extract principle of assay from remaining paragraphs
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
        
        # Process data analysis section - remove Boster reference and unwanted sections
        if 'data_analysis' in processed_data:
            data_analysis = processed_data['data_analysis']
            
            # Remove first two sentences if they contain Boster references
            sentences = re.split(r'(?<=[.!?])\s+', data_analysis)
            if len(sentences) > 2 and any(word in ' '.join(sentences[:2]).lower() for word in ['boster', 'biocompare', 'online']):
                processed_data['data_analysis'] = ' '.join(sentences[2:])
            else:
                processed_data['data_analysis'] = data_analysis
                
            # Remove the Publications and Submit a Product Review sections
            processed_data['data_analysis'] = re.sub(r'Publications.*?using this product.*?$', '', processed_data['data_analysis'], flags=re.DOTALL | re.IGNORECASE)
            processed_data['data_analysis'] = re.sub(r'Submit a Product Review to Biocompare.*?$', '', processed_data['data_analysis'], flags=re.DOTALL | re.IGNORECASE)
        
        # Handle required materials which should already be a list from the parser
        if 'required_materials' in processed_data:
            # This is now returned directly as a list from the parser - just copy to materials_list
            processed_data['required_materials_list'] = processed_data['required_materials']
            # Also keep original format for compatibility
            processed_data['required_materials_text'] = "\n".join(processed_data['required_materials'])
            
            # Prepare materials as a simple list
            materials = processed_data['required_materials']
            if materials:
                # Clean up the items
                clean_materials = []
                for item in materials:
                    if item.strip():
                        clean_materials.append(item.strip())
                processed_data['required_materials_list_items'] = clean_materials
                
        # Format assay protocol as numbered steps
        if 'assay_protocol' in processed_data and processed_data['assay_protocol']:
            protocol = processed_data['assay_protocol']
            if protocol:
                # Keep the original protocol steps
                processed_data['assay_protocol_steps'] = protocol
                
                # Also create a numbered version for text display
                numbered_steps = []
                for i, step in enumerate(protocol, 1):
                    numbered_steps.append(f"{i}. {step}")
                processed_data['assay_protocol_numbered'] = "\n".join(numbered_steps)
                
        # Format standard curve data for table display - just use the original data
        if 'standard_curve_table' in processed_data and processed_data['standard_curve_table']:
            # Make a copy to avoid unwanted modifications
            processed_data['standard_curve_data'] = processed_data['standard_curve_table']
            
        # Process overview specifications table data
        if 'overview_specifications' in processed_data and processed_data['overview_specifications']:
            # Clean up the specifications data for display in the template
            cleaned_specs = []
            for spec in processed_data['overview_specifications']:
                if 'property' in spec and 'value' in spec:
                    # Replace "Boster" with "Innovative Research" in values
                    value = re.sub(r'\bBoster\b', 'Innovative Research', spec['value'])
                    value = re.sub(r'\bBOSTER\b', 'INNOVATIVE RESEARCH', value)
                    value = re.sub(r'\bboster\b', 'innovative research', value)
                    
                    # Remove trademark symbols
                    value = re.sub(r'®', '', value)
                    value = re.sub(r'™', '', value)
                    
                    # Remove any "PicoKine" references
                    value = re.sub(r'PicoKine\s*®', '', value)
                    value = re.sub(r'Picokine\s*®', '', value)
                    value = re.sub(r'PicoKine', '', value)
                    value = re.sub(r'Picokine', '', value)
                    
                    # Skip empty values
                    if value.strip():
                        cleaned_specs.append({
                            'property': spec['property'],
                            'value': value.strip()
                        })
            
            processed_data['overview_specifications_table'] = cleaned_specs
                
        # Process technical details (now it's a dictionary with 'text' and 'technical_table')
        if 'technical_details' in processed_data:
            technical_details = processed_data['technical_details']
            
            # Handle the text part
            if isinstance(technical_details, dict) and 'text' in technical_details:
                text_content = technical_details['text']
                # Clean up any extraneous text like "Cross-reactivity:" or empty lines
                tech_lines = []
                for line in text_content.split('\n'):
                    line = line.strip()
                    if line:
                        # If it's just a header line with no data, skip it
                        if line.endswith(':') and len(line) < 30:
                            continue
                        # Remove specific supplier references
                        line = re.sub(r'from (Boster|PicoKine|EK[0-9]+)', '', line)
                        tech_lines.append(line)
                processed_data['technical_details'] = '\n\n'.join(tech_lines)
            elif isinstance(technical_details, str):
                # If technical_details is already a string, keep it as is
                processed_data['technical_details'] = technical_details
            else:
                processed_data['technical_details'] = ''
            
            # Handle the table part
            if isinstance(technical_details, dict) and 'technical_table' in technical_details:
                # Ensure all fields have values
                for item in technical_details['technical_table']:
                    if not item['value']:
                        item['value'] = 'N/A'
                processed_data['technical_details_table'] = technical_details['technical_table']
            else:
                # Fallback empty table with placeholder values
                processed_data['technical_details_table'] = [
                    {'property': 'Capture/Detection Antibodies', 'value': 'N/A'},
                    {'property': 'Specificity', 'value': 'N/A'},
                    {'property': 'Standard Protein', 'value': 'N/A'},
                    {'property': 'Cross-reactivity', 'value': 'N/A'}
                ]
                
        # Process preparations before assay
        if 'preparations_before_assay' in processed_data:
            prep_data = processed_data['preparations_before_assay']
            
            # If it's a dictionary with 'text' and 'steps' keys
            if isinstance(prep_data, dict) and 'text' in prep_data and 'steps' in prep_data:
                # Extract the non-step portions of the text
                non_step_text = prep_data['text']
                # Find all steps in the text
                for step in prep_data['steps']:
                    # Remove the numbered steps from the main text
                    step_text = f"{step['number']}. {step['text']}"
                    non_step_text = non_step_text.replace(step_text, "")
                
                # Clean up the non-step text by removing extra whitespace and empty lines
                non_step_text_lines = [line.strip() for line in non_step_text.split('\n') if line.strip()]
                processed_data['preparations_text'] = "\n\n".join(non_step_text_lines)
                
                # Check if we have numbered steps
                if prep_data['steps']:
                    # Make sure the steps are sorted by number
                    sorted_steps = sorted(prep_data['steps'], key=lambda x: x['number'])
                    
                    # Make sure we have a proper sequence (1, 2, 3, 4, etc.)
                    fixed_steps = []
                    for i, step in enumerate(sorted_steps, 1):
                        fixed_steps.append({
                            'number': i,
                            'text': step['text']
                        })
                    
                    # Create a numbered list for text display (but we'll use the actual step objects for rendering)
                    numbered_steps = []
                    for step in fixed_steps:
                        numbered_steps.append(f"{step['number']}. {step['text']}")
                    
                    processed_data['preparations_numbered'] = "\n".join(numbered_steps)
                    # Use the fixed and sorted steps for the template
                    processed_data['preparations_steps'] = fixed_steps
                else:
                    # No numbered steps, use the same text
                    processed_data['preparations_numbered'] = processed_data['preparations_text']
                    processed_data['preparations_steps'] = []
            elif isinstance(prep_data, str):
                # Handle the old format where prep_data is a string
                processed_data['preparations_text'] = prep_data
                processed_data['preparations_numbered'] = prep_data
                processed_data['preparations_steps'] = []
                
        # Define patterns to remove for all text processing
        patterns_to_remove = [
            r'For more information on.*?\.', 
            r'For additional information.*?\.', 
            r'Visit (?:our|the) (?:website|resource center).*?\.', 
            r'Please refer to (?:our|the) (?:website|resource center).*?\.', 
            r'More details can be found at.*?\.', 
            r'Technical support (?:is|can be) available.*?\.', 
            r'Visit.*?\.(?:com|org|net).*?\.', 
            r'.*?resource center at.*?\.',
            r'.*?ELISA Resource Center.*?\.',
            r'.*?technical resource center.*?\.',
            r'For more information on assay principle, protocols, and troubleshooting tips, see.*'
        ]

        # Clean up data to remove unwanted content and replace company names
        for key, value in processed_data.items():
            if isinstance(value, str):
                # Replace "Boster" with "Innovative Research"
                value = re.sub(r'\bBoster\b', 'Innovative Research', value)
                value = re.sub(r'\bBOSTER\b', 'INNOVATIVE RESEARCH', value)
                value = re.sub(r'\bboster\b', 'innovative research', value)
                
                # Remove all trademark and registered trademark symbols
                value = re.sub(r'®', '', value)
                value = re.sub(r'™', '', value)
                value = re.sub(r'©', '', value)
                
                # Remove all variations of PicoKine®
                value = re.sub(r'PicoKine\s*®', '', value)
                value = re.sub(r'Picokine\s*®', '', value)
                value = re.sub(r'PicoKine', '', value)
                value = re.sub(r'Picokine', '', value)
                
                # Remove references to online tools and Biocompare product reviews
                value = re.sub(r'offers an easy-to-use online ELISA data analysis tool\. Try it out at.*?\.com.*?online', '', value)
                value = re.sub(r'Submit a (?:product )?review (?:of this product )?to Biocompare\.com.*?contribution\.', '', value, flags=re.IGNORECASE | re.DOTALL)
                value = re.sub(r'Submit a (?:product )?review (?:of this product )?to Biocompare.*?gift card.*', '', value, flags=re.IGNORECASE | re.DOTALL)
                value = re.sub(r'.*?receive a \$[0-9]+ Amazon\.com gift card.*', '', value, flags=re.IGNORECASE | re.DOTALL)
                
                # Remove references to resource centers and external URLs
                for pattern in patterns_to_remove:
                    value = re.sub(pattern, '', value, flags=re.IGNORECASE | re.DOTALL)
                
                # Final cleanup
                value = re.sub(r'\s+', ' ', value)  # Replace multiple spaces with single space
                value = value.strip()
                
                processed_data[key] = value
            elif isinstance(value, list):
                if all(isinstance(item, dict) for item in value):
                    # Handle lists of dictionaries (like reagents, tables, etc.)
                    for item in value:
                        for item_key, item_value in item.items():
                            if isinstance(item_value, str):
                                # Apply the same replacements to dictionary values
                                replaced_value = item_value
                                replaced_value = re.sub(r'\bBoster\b', 'Innovative Research', replaced_value)
                                replaced_value = re.sub(r'\bBOSTER\b', 'INNOVATIVE RESEARCH', replaced_value)
                                replaced_value = re.sub(r'\bboster\b', 'innovative research', replaced_value)
                                
                                # Remove all trademark and registered trademark symbols
                                replaced_value = re.sub(r'®', '', replaced_value)
                                replaced_value = re.sub(r'™', '', replaced_value)
                                replaced_value = re.sub(r'©', '', replaced_value)
                                
                                # Remove all variations of PicoKine®
                                replaced_value = re.sub(r'PicoKine\s*®', '', replaced_value)
                                replaced_value = re.sub(r'Picokine\s*®', '', replaced_value)
                                replaced_value = re.sub(r'PicoKine', '', replaced_value)
                                replaced_value = re.sub(r'Picokine', '', replaced_value)
                                
                                # Remove references to online tools
                                replaced_value = re.sub(r'offers an easy-to-use online ELISA data analysis tool\. Try it out at.*?\.com.*?online', '', replaced_value)
                                replaced_value = re.sub(r'Submit a (?:product )?review (?:of this product )?to Biocompare', '', replaced_value, flags=re.IGNORECASE)
                                
                                item[item_key] = replaced_value
                elif all(isinstance(item, str) for item in value):
                    # Handle lists of strings (like required_materials_list)
                    processed_list = []
                    for item in value:
                        # Apply all the same replacements and cleanup to list items
                        item = re.sub(r'\bBoster\b', 'Innovative Research', item)
                        item = re.sub(r'\bBOSTER\b', 'INNOVATIVE RESEARCH', item)
                        item = re.sub(r'\bboster\b', 'innovative research', item)
                        
                        # Remove all trademark and registered trademark symbols
                        item = re.sub(r'®', '', item)
                        item = re.sub(r'™', '', item)
                        item = re.sub(r'©', '', item)
                        
                        # Remove all variations of PicoKine®
                        item = re.sub(r'PicoKine\s*®', '', item)
                        item = re.sub(r'Picokine\s*®', '', item)
                        item = re.sub(r'PicoKine', '', item)
                        item = re.sub(r'Picokine', '', item)
                        
                        # Remove references to Biocompare
                        item = re.sub(r'Submit a (?:product )?review (?:of this product )?to Biocompare\.com.*', '', item, flags=re.IGNORECASE | re.DOTALL)
                        
                        # Final cleanup
                        item = re.sub(r'\s+', ' ', item)  # Replace multiple spaces with single space
                        item = item.strip()
                        
                        processed_list.append(item)
                    
                    processed_data[key] = processed_list
        
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
        try:
            # Clean and prepare the data
            processed_data = self._clean_data(data, kit_name, catalog_number, lot_number)
            
            # Map reagent data to static individual fields in the template
            if 'reagents' in processed_data:
                reagents = processed_data['reagents']
                # Add individual reagent entries for up to 7 rows
                for i in range(min(len(reagents), 7)):
                    reagent = reagents[i]
                    # Fill in each column for this reagent
                    if isinstance(reagent, dict):
                        processed_data[f'reagent_{i+1}_name'] = reagent.get('name', '')
                        processed_data[f'reagent_{i+1}_quantity'] = reagent.get('quantity', '')
                        processed_data[f'reagent_{i+1}_volume'] = reagent.get('volume', '')
                        processed_data[f'reagent_{i+1}_storage'] = reagent.get('storage', '')
                
                # Clear any unused reagent slots
                for i in range(len(reagents) + 1, 8):
                    processed_data[f'reagent_{i}_name'] = ''
                    processed_data[f'reagent_{i}_quantity'] = ''
                    processed_data[f'reagent_{i}_volume'] = ''
                    processed_data[f'reagent_{i}_storage'] = ''
            
            # Map required materials to individual bullet points
            if 'required_materials' in processed_data:
                req_materials = processed_data['required_materials']
                # Add individual material entries
                for i in range(min(len(req_materials), 10)):
                    processed_data[f'req_material_{i+1}'] = req_materials[i]
                
                # Clear any unused material slots
                for i in range(len(req_materials) + 1, 11):
                    processed_data[f'req_material_{i}'] = ''
            
            # Map standard curve data to individual fields
            if 'standard_curve' in processed_data:
                # Check format of standard curve data
                if 'concentration' in processed_data['standard_curve'] and 'od' in processed_data['standard_curve']:
                    conc_values = processed_data['standard_curve']['concentration']
                    od_values = processed_data['standard_curve']['od']
                elif 'concentrations' in processed_data['standard_curve'] and 'od_values' in processed_data['standard_curve']:
                    conc_values = processed_data['standard_curve']['concentrations']
                    od_values = processed_data['standard_curve']['od_values']
                else:
                    conc_values = []
                    od_values = []
                
                # Map to individual fields in template
                for i in range(min(len(conc_values), len(od_values), 8)):
                    processed_data[f'std_conc_{i+1}'] = conc_values[i]
                    processed_data[f'std_od_{i+1}'] = od_values[i]
                
                # Clear any unused slots
                for i in range(min(len(conc_values), len(od_values)) + 1, 9):
                    processed_data[f'std_conc_{i}'] = ''
                    processed_data[f'std_od_{i}'] = ''
            
            # Map assay protocol steps to numbered list items
            if 'assay_protocol' in processed_data:
                protocol_steps = processed_data['assay_protocol']
                # Add individual protocol step entries
                for i in range(min(len(protocol_steps), 20)):
                    processed_data[f'protocol_step_{i+1}'] = protocol_steps[i]
                
                # Clear any unused steps
                for i in range(len(protocol_steps) + 1, 21):
                    processed_data[f'protocol_step_{i}'] = ''
            
            # Render the template with the context data
            self.template.render(processed_data)
            
            # Save the rendered template to the output path
            self.template.save(output_path)
            
            self.logger.info(f"Template successfully populated and saved to {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error populating template: {e}")
            raise