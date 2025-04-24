"""
Template Populator Enhanced
-----------------
Enhanced version of the template populator to handle the new technical details table format.
"""

import re
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional

import docx
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor

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
        self.template = DocxTemplate(template_path)
        self.logger = logging.getLogger(__name__)
    
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
        # Start with a copy of the data to avoid modifying the original
        processed_data = dict(data)
        
        # Override with user-provided values if available
        if kit_name:
            processed_data['kit_name'] = kit_name
        elif 'catalog_number' in processed_data:
            # Default kit name based on catalog number
            catalog = processed_data['catalog_number']
            processed_data['kit_name'] = f"Mouse KLK1/Kallikrein 1 ELISA Kit ({catalog})"
        
        if catalog_number:
            processed_data['catalog_number'] = catalog_number
            
        if lot_number:
            processed_data['lot_number'] = lot_number
        
        # Process background text for the template
        if 'background' in processed_data:
            processed_data['background_text'] = processed_data['background']
        
        # Process intended use with additional information if needed
        if 'intended_use' in processed_data:
            processed_data['intended_use'] = processed_data['intended_use']
            
        # Process required materials list
        if 'required_materials' in processed_data:
            materials = processed_data['required_materials']
            if isinstance(materials, list):
                # Format as a bulleted list with proper bullet points
                formatted_materials = []
                for item in materials:
                    formatted_materials.append(f"• {item}")
                
                processed_data['required_materials_with_bullets'] = "\n".join(formatted_materials)
            else:
                # Single string with each item on a new line, prefixed with bullet
                processed_data['required_materials_with_bullets'] = f"• {materials}"
                
        # Process assay protocol steps
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
                    
                    # Remove all trademark and registered trademark symbols
                    value = re.sub(r'®', '', value)
                    value = re.sub(r'™', '', value)
                    value = re.sub(r'©', '', value)
                    
                    # Clean up the value
                    value = value.strip()
                    
                    cleaned_specs.append({'property': spec['property'], 'value': value})
            
            processed_data['overview_specifications'] = cleaned_specs
            
        # Process technical details for the enhanced template format
        if 'technical_details' in processed_data and processed_data['technical_details']:
            technical_details = processed_data['technical_details']
            if isinstance(technical_details, dict) and 'technical_table' in technical_details:
                # Make sure values are not None
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
        
        # Add structured variability data for the new template format
        processed_data['variability'] = {
            'intra_assay': {
                'sample_1': {
                    'n': processed_data.get('intra_var_sample1_n', '24'),
                    'mean': processed_data.get('intra_var_sample1_mean', '145'),
                    'sd': processed_data.get('intra_var_sample1_sd', '10.15'),
                    'cv': processed_data.get('intra_var_sample1_cv', '7.0%')
                },
                'sample_2': {
                    'n': processed_data.get('intra_var_sample2_n', '24'),
                    'mean': processed_data.get('intra_var_sample2_mean', '329'),
                    'sd': processed_data.get('intra_var_sample2_sd', '23.03'),
                    'cv': processed_data.get('intra_var_sample2_cv', '7.0%')
                },
                'sample_3': {
                    'n': processed_data.get('intra_var_sample3_n', '24'),
                    'mean': processed_data.get('intra_var_sample3_mean', '1062'),
                    'sd': processed_data.get('intra_var_sample3_sd', '65.84'),
                    'cv': processed_data.get('intra_var_sample3_cv', '6.2%')
                }
            },
            'inter_assay': {
                'sample_1': {
                    'n': processed_data.get('inter_var_sample1_n', '24'),
                    'mean': processed_data.get('inter_var_sample1_mean', '145'),
                    'sd': processed_data.get('inter_var_sample1_sd', '13.05'),
                    'cv': processed_data.get('inter_var_sample1_cv', '9.0%')
                },
                'sample_2': {
                    'n': processed_data.get('inter_var_sample2_n', '24'),
                    'mean': processed_data.get('inter_var_sample2_mean', '329'),
                    'sd': processed_data.get('inter_var_sample2_sd', '29.61'),
                    'cv': processed_data.get('inter_var_sample2_cv', '9.0%')
                },
                'sample_3': {
                    'n': processed_data.get('inter_var_sample3_n', '24'),
                    'mean': processed_data.get('inter_var_sample3_mean', '1062'),
                    'sd': processed_data.get('inter_var_sample3_sd', '95.58'),
                    'cv': processed_data.get('inter_var_sample3_cv', '9.0%')
                }
            }
        }
        
        # Set up reproducibility data with standard deviation
        processed_data['reproducibility'] = [
            {
                'sample': 'Sample 1',
                'lot1': processed_data.get('repro_sample1_lot1', '150'),
                'lot2': processed_data.get('repro_sample1_lot2', '154'),
                'lot3': processed_data.get('repro_sample1_lot3', '170'),
                'lot4': processed_data.get('repro_sample1_lot4', '150'),
                'sd': processed_data.get('repro_sample1_sd', '9.4'),
                'mean': processed_data.get('repro_sample1_mean', '156'),
                'cv': processed_data.get('repro_sample1_cv', '5.2%')
            },
            {
                'sample': 'Sample 2',
                'lot1': processed_data.get('repro_sample2_lot1', '600'),
                'lot2': processed_data.get('repro_sample2_lot2', '580'),
                'lot3': processed_data.get('repro_sample2_lot3', '595'),
                'lot4': processed_data.get('repro_sample2_lot4', '605'),
                'sd': processed_data.get('repro_sample2_sd', '11.3'),
                'mean': processed_data.get('repro_sample2_mean', '595'),
                'cv': processed_data.get('repro_sample2_cv', '1.9%')
            },
            {
                'sample': 'Sample 3',
                'lot1': processed_data.get('repro_sample3_lot1', '1010'),
                'lot2': processed_data.get('repro_sample3_lot2', '970'),
                'lot3': processed_data.get('repro_sample3_lot3', '990'),
                'lot4': processed_data.get('repro_sample3_lot4', '1030'),
                'sd': processed_data.get('repro_sample3_sd', '25.7'),
                'mean': processed_data.get('repro_sample3_mean', '1000'),
                'cv': processed_data.get('repro_sample3_cv', '2.6%')
            }
        ]
        
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
        # Direct access to manipulate tables after Jinja template rendering
        self._post_process_doc = None
        try:
            # Clean and prepare the data
            processed_data = self._clean_data(data, kit_name, catalog_number, lot_number)
            
            # Map reagent data to static individual fields in the template
            if 'reagents' in processed_data:
                reagents = processed_data['reagents']
                self.logger.info(f"Processing {len(reagents)} reagents for template")
                
                # Add individual reagent entries for up to 12 rows (increased from 7)
                for i in range(min(len(reagents), 12)):
                    reagent = reagents[i]
                    # Fill in each column for this reagent
                    processed_data[f'reagent_{i+1}_name'] = reagent.get('name', '')
                    processed_data[f'reagent_{i+1}_quantity'] = reagent.get('quantity', '')
                    processed_data[f'reagent_{i+1}_volume'] = reagent.get('volume', '')
                    processed_data[f'reagent_{i+1}_storage'] = reagent.get('storage', '')
            
            # Process required materials for the template
            if 'required_materials' in processed_data:
                materials = processed_data['required_materials']
                self.logger.info(f"Processing {len(materials)} required materials for template")
                
                # Format as a bulleted list for display in the template
                if isinstance(materials, list):
                    processed_data['required_materials_with_bullets'] = "• " + "\n• ".join(materials)
                else:
                    processed_data['required_materials_with_bullets'] = f"• {materials}"
            
            # Process standard curve data for the template
            if 'standard_curve' in processed_data:
                standard_curve = processed_data['standard_curve']
                if isinstance(standard_curve, dict) and 'concentrations' in standard_curve and 'od_values' in standard_curve:
                    self.logger.info(f"Processing standard curve data: {len(standard_curve['concentrations'])} concentrations, {len(standard_curve['od_values'])} OD values")
            
            # Process assay protocol steps for the template and individual step fields
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
            
            # Load the document for post-processing
            doc = Document(output_path)
            
            # Format the document header and first page
            self._format_document_header(doc)
            
            # Save the formatted document
            doc.save(output_path)
            
            # Post-process the document to directly modify the kit components table
            self._post_process_kit_components(output_path, processed_data)
            
            self.logger.info(f"Template successfully populated and saved to {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error populating template: {e}")
            raise
            
    def _format_document_header(self, doc):
        """
        Format the document header to be size 36pt with Title style.
        Also ensure the first page only contains title, catalog number, lot number, 
        and intended use by adding page breaks.
        
        Args:
            doc: The Document object to modify
        """
        # Format the document title (first paragraph should be the title)
        if len(doc.paragraphs) > 0:
            title_para = doc.paragraphs[0]
            title_para.style = 'Title'
            
            # Set font size to 36pt for the title
            for run in title_para.runs:
                run.font.size = Pt(36)
                run.font.bold = True
            
            # Make sure the title is centered
            title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # Find the intended use section (should be within first few paragraphs)
        intended_use_idx = None
        for i, para in enumerate(doc.paragraphs[:10]):  # Check the first 10 paragraphs
            if 'intended use' in para.text.lower() or 'purpose' in para.text.lower():
                intended_use_idx = i
                break
        
        # If found, add page break after the intended use section
        if intended_use_idx is not None:
            # Look for the end of the intended use section (usually a paragraph or two)
            # We'll look for the next section heading as the end marker
            end_idx = intended_use_idx
            for i in range(intended_use_idx + 1, min(intended_use_idx + 5, len(doc.paragraphs))):
                # Look for the next heading or all-caps paragraph 
                # (common formatting for section headings)
                if (doc.paragraphs[i].style.name.startswith('Heading') or 
                    doc.paragraphs[i].text.isupper() or
                    'TECHNICAL' in doc.paragraphs[i].text or
                    'OVERVIEW' in doc.paragraphs[i].text):
                    # Found the next section, so put page break at previous paragraph
                    end_idx = i - 1
                    break
                
                # Include this paragraph as part of intended use
                end_idx = i
            
            # If there are runs in the paragraph
            if len(doc.paragraphs[end_idx].runs) > 0:
                # Add page break after the intended use section
                doc.paragraphs[end_idx].runs[-1].add_break(docx.enum.text.WD_BREAK.PAGE)
            else:
                # No runs, add a run with page break
                run = doc.paragraphs[end_idx].add_run()
                run.add_break(docx.enum.text.WD_BREAK.PAGE)
        else:
            # If intended use not found, just add page break after first few paragraphs
            if len(doc.paragraphs) > 5:
                if len(doc.paragraphs[3].runs) > 0:  # After intended use description (usually paragraph 3)
                    doc.paragraphs[3].runs[-1].add_break(docx.enum.text.WD_BREAK.PAGE)
                else:
                    run = doc.paragraphs[3].add_run()
                    run.add_break(docx.enum.text.WD_BREAK.PAGE)

    def _post_process_kit_components(self, output_path: Path, processed_data: Dict[str, Any]) -> None:
        """
        Perform post-processing on the populated template to handle the kit components table.
        This directly modifies the DOCX after the Jinja2 template rendering is complete.
        
        Args:
            output_path: Path to the populated template file
            processed_data: Dictionary containing the processed data used for template population
        """
        try:
            if 'reagents' not in processed_data:
                self.logger.warning("No reagents data found for post-processing")
                return
                
            # Load the document to modify tables directly
            doc = Document(output_path)
            
            # Find the kit components section
            kit_components_section_idx = None
            for i, para in enumerate(doc.paragraphs):
                text = para.text.strip().lower()
                if "kit components" in text or "materials provided" in text:
                    self.logger.info(f"Found Kit Components section at paragraph {i}: {para.text}")
                    kit_components_section_idx = i
                    break
            
            if kit_components_section_idx is None:
                self.logger.warning("Kit Components section not found, cannot update table")
                return
            
            # Identify the correct kit components table
            kit_components_table_idx = None
            
            # First check if there's a 4-column table (preferred format)
            for i, table in enumerate(doc.tables):
                if len(table.columns) == 4:
                    # Check headers
                    try:
                        header_row = [cell.text.strip().lower() for cell in table.rows[0].cells]
                        if len(header_row) == 4 and any(keyword in " ".join(header_row) for keyword in 
                                                      ["description", "quantity", "volume", "storage"]):
                            self.logger.info(f"Found 4-column kit components table at index {i}")
                            kit_components_table_idx = i
                            break
                    except:
                        pass
            
            # If 4-column table not found, use the first table after the kit components section
            if kit_components_table_idx is None:
                # Just take the first table after the section (usually Table 3)
                kit_components_table_idx = 2
                self.logger.info(f"Using table at index {kit_components_table_idx} for kit components")
            
            if kit_components_table_idx >= len(doc.tables):
                self.logger.warning(f"Table index {kit_components_table_idx} is out of bounds")
                return
                
            # Get the kit components table
            kit_table = doc.tables[kit_components_table_idx]
            
            # Clear out existing content in kit components table (keep header row)
            for row_idx in range(1, len(kit_table.rows)):
                for cell in kit_table.rows[row_idx].cells:
                    for paragraph in cell.paragraphs:
                        paragraph.clear()
            
            # Fill in the table with the reagent data
            reagents = processed_data['reagents']
            
            # If we need more rows, add them
            while len(kit_table.rows) < len(reagents) + 1:  # +1 for header row
                kit_table.add_row()
            
            # Populate reagent rows
            for i, reagent in enumerate(reagents):
                if i >= len(kit_table.rows) - 1:  # Skip header row
                    break
                    
                row_idx = i + 1  # Skip header row
                
                # Check if enough cells in row
                if len(kit_table.rows[row_idx].cells) >= 4:
                    if 'name' in reagent:
                        kit_table.rows[row_idx].cells[0].text = reagent['name']
                    if 'quantity' in reagent:
                        kit_table.rows[row_idx].cells[1].text = reagent['quantity']
                    if 'volume' in reagent:
                        kit_table.rows[row_idx].cells[2].text = reagent['volume']
                    if 'storage' in reagent:
                        kit_table.rows[row_idx].cells[3].text = reagent['storage']
            
            # Save the modified document
            doc.save(output_path)
            self.logger.info(f"Updated kit components table with {len(reagents)} reagents")
            
        except Exception as e:
            self.logger.error(f"Error in post-processing kit components: {e}")
            # Continue anyway - this is just an enhancement