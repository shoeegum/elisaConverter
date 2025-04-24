"""
Utility functions for ELISA datasheet processing.
"""

import re
import logging
from typing import Dict, Any, List, Optional, Union

def clean_text(text: str) -> str:
    """
    Clean text by removing extra whitespace and normalizing newlines.
    
    Args:
        text: The text to clean
        
    Returns:
        Cleaned text
    """
    if not text:
        return ""
        
    # Replace multiple spaces with a single space
    text = re.sub(r'\s+', ' ', text)
    
    # Remove leading/trailing whitespace
    text = text.strip()
    
    return text

def extract_numeric_value(text: str) -> Optional[str]:
    """
    Extract a numeric value from text.
    
    Args:
        text: The text to extract a numeric value from
        
    Returns:
        The extracted numeric value as a string, or None if no numeric value is found
    """
    match = re.search(r'\d+(?:\.\d+)?', text)
    return match.group(0) if match else None

def format_table_data(data: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Format table data for template rendering.
    
    Args:
        data: List of dictionaries containing table data
        
    Returns:
        Formatted table data
    """
    formatted_data = []
    for row in data:
        formatted_row = {}
        for key, value in row.items():
            if isinstance(value, (int, float)):
                formatted_row[key] = str(value)
            else:
                formatted_row[key] = clean_text(str(value))
        formatted_data.append(formatted_row)
    return formatted_data

def find_nearest_paragraph(paragraphs: List[Any], index: int, text: str, forward: bool = True) -> Optional[int]:
    """
    Find the nearest paragraph containing the specified text.
    
    Args:
        paragraphs: List of paragraphs to search
        index: The index to start searching from
        text: The text to search for
        forward: Whether to search forward (True) or backward (False)
        
    Returns:
        The index of the nearest paragraph containing the text, or None if not found
    """
    step = 1 if forward else -1
    start = index
    end = len(paragraphs) if forward else -1
    
    for i in range(start, end, step):
        if text.lower() in paragraphs[i].text.lower():
            return i
    
    return None

def convert_units(value: str, from_unit: str, to_unit: str) -> str:
    """
    Convert a value from one unit to another.
    
    Args:
        value: The value to convert
        from_unit: The unit to convert from
        to_unit: The unit to convert to
        
    Returns:
        The converted value as a string
    """
    logger = logging.getLogger(__name__)
    
    # Extract numeric part
    try:
        numeric_value = float(re.search(r'\d+(?:\.\d+)?', value).group(0))
    except (AttributeError, ValueError):
        logger.warning(f"Could not extract numeric value from {value}")
        return value
    
    # Define conversion factors
    conversions = {
        # Concentration conversions
        'pg_to_ng': 0.001,
        'ng_to_pg': 1000,
        'ng_to_ug': 0.001,
        'ug_to_ng': 1000,
        'ug_to_mg': 0.001,
        'mg_to_ug': 1000,
        
        # Volume conversions
        'ul_to_ml': 0.001,
        'ml_to_ul': 1000,
        'ml_to_l': 0.001,
        'l_to_ml': 1000,
    }
    
    # Normalize units
    from_unit = from_unit.lower().replace('μ', 'u').replace('µ', 'u')
    to_unit = to_unit.lower().replace('μ', 'u').replace('µ', 'u')
    
    # Create conversion key
    conversion_key = f"{from_unit.split('/')[0]}_to_{to_unit.split('/')[0]}"
    
    # Convert value if conversion exists
    if conversion_key in conversions:
        converted_value = numeric_value * conversions[conversion_key]
        # Format the number appropriately
        if converted_value < 0.01:
            formatted_value = f"{converted_value:.2e}"
        elif converted_value < 1:
            formatted_value = f"{converted_value:.3f}"
        else:
            formatted_value = f"{converted_value:.1f}"
        
        # Replace unit
        result = re.sub(r'\d+(?:\.\d+)?', formatted_value, value)
        result = result.replace(from_unit, to_unit)
        
        return result
    else:
        logger.warning(f"No conversion defined for {from_unit} to {to_unit}")
        return value
