#!/usr/bin/env python3
"""
Update the template populator to handle the new template format.
"""

import logging
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def update_template_populator():
    """Update the template populator with code for the new template format."""
    template_path = Path('template_populator_enhanced.py')
    
    # Read the current file
    with open(template_path, 'r') as f:
        content = f.read()
    
    # Add new variability handling
    variability_code = """
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
"""
    
    # Add the new code before the render template line
    # Find the position to insert the new code
    insertion_point = content.find("            # Render the template with the context data")
    
    if insertion_point == -1:
        logger.error("Could not find the insertion point in the template populator")
        return
    
    # Insert the new code
    new_content = content[:insertion_point] + variability_code + content[insertion_point:]
    
    # Write the updated file
    with open(template_path, 'w') as f:
        f.write(new_content)
    
    logger.info(f"Updated template populator at {template_path}")

if __name__ == "__main__":
    update_template_populator()