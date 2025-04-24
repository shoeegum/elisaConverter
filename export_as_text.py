#!/usr/bin/env python3
"""
Export the extracted data as a simple text file.
This approach avoids Word document corruption issues completely.
"""

import sys
import argparse
import logging
from pathlib import Path

from elisa_parser import ELISADatasheetParser

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def export_as_text(data, output_path, kit_name=None, catalog_number=None, lot_number=None):
    """
    Export the extracted data as a text file.
    
    Args:
        data: Dictionary containing the extracted data
        output_path: Path where the text file will be saved
        kit_name: Optional kit name provided by user
        catalog_number: Optional catalog number provided by user
        lot_number: Optional lot number provided by user
    """
    # Use provided values or extract from data
    kit_name = kit_name or data.get('kit_name', 'Mouse KLK1 ELISA Kit')
    catalog_number = catalog_number or data.get('catalog_number', 'N/A')
    lot_number = lot_number or data.get('lot_number', 'SAMPLE')
    
    # Format the text content
    content = [
        f"{kit_name}\n",
        f"Catalog #: {catalog_number} | Lot #: {lot_number}\n",
        f"\n{'=' * 60}\n",
        "\nINTENDED USE\n",
        f"{data.get('intended_use', 'No intended use information available.')}\n",
        f"\n{'=' * 60}\n",
        "\nBACKGROUND\n",
        f"{data.get('background', 'No background information available.')}\n",
        f"\n{'=' * 60}\n",
        "\nPRINCIPLE OF THE ASSAY\n",
        f"{data.get('assay_principle', 'No assay principle information available.')}\n",
        f"\n{'=' * 60}\n",
        "\nTECHNICAL DETAILS\n",
        f"Sensitivity: {data.get('sensitivity', 'N/A')}\n",
        f"Detection Range: {data.get('detection_range', 'N/A')}\n",
        f"Specificity: {data.get('specificity', 'N/A')}\n",
        f"Cross-reactivity: {data.get('cross_reactivity', 'N/A')}\n",
        f"\n{'=' * 60}\n",
        "\nKIT COMPONENTS\n"
    ]
    
    # Add reagents
    reagents = data.get('reagents', [])
    if reagents:
        for reagent in reagents:
            content.append(f"* {reagent.get('name', '')}: {reagent.get('quantity', '')}\n")
    else:
        content.append("No reagent information available.\n")
    
    content.extend([
        f"\n{'=' * 60}\n",
        "\nMATERIALS REQUIRED BUT NOT PROVIDED\n"
    ])
    
    # Add materials
    materials = data.get('required_materials', [])
    if materials:
        for material in materials:
            content.append(f"* {material}\n")
    else:
        content.append("No materials information available.\n")
    
    content.extend([
        f"\n{'=' * 60}\n",
        "\nASSAY PROTOCOL\n"
    ])
    
    # Add protocol steps
    protocol = data.get('assay_protocol', [])
    if protocol:
        for i, step in enumerate(protocol, 1):
            content.append(f"{i}. {step}\n")
    else:
        content.append("No protocol information available.\n")
    
    content.extend([
        f"\n{'=' * 60}\n",
        "\nINNOVATIVE RESEARCH\n",
        "35200 Schoolcraft Rd, Livonia, MI 48150 | Phone: (248) 896-0142 | Fax: (248) 896-0148\n"
    ])
    
    # Write to the output file
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.writelines(content)
        logger.info(f"Text file successfully created at {output_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving text file: {e}")
        return False

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description='Export ELISA datasheet data as a text file'
    )
    
    parser.add_argument(
        '--source', 
        required=True,
        help='Path to the source ELISA kit datasheet DOCX file'
    )
    
    parser.add_argument(
        '--output', 
        required=True,
        help='Path where the text file will be saved'
    )
    
    parser.add_argument(
        '--kit-name',
        help='Name of the ELISA kit'
    )
    
    parser.add_argument(
        '--catalog-number',
        help='Catalog number of the ELISA kit'
    )
    
    parser.add_argument(
        '--lot-number',
        help='Lot number of the ELISA kit'
    )
    
    return parser.parse_args()

def main():
    """Main entry point"""
    args = parse_arguments()
    
    try:
        # Parse the ELISA datasheet
        logger.info(f"Parsing ELISA datasheet: {args.source}")
        parser = ELISADatasheetParser(args.source)
        data = parser.extract_data()
        
        # Export as text
        logger.info(f"Exporting data as text file: {args.output}")
        success = export_as_text(
            data, 
            args.output, 
            kit_name=args.kit_name, 
            catalog_number=args.catalog_number, 
            lot_number=args.lot_number
        )
        
        if success:
            print(f"✅ Text file successfully created: {args.output}")
            return 0
        else:
            print(f"❌ Failed to create text file")
            return 1
    
    except Exception as e:
        logger.exception(f"Error processing file: {e}")
        print(f"❌ Error: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())