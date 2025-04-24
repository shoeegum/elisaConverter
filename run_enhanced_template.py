#!/usr/bin/env python3
"""
Enhanced Template Runner

This script runs the ELISA parser with the enhanced template.
"""

import os
import sys
import subprocess

def main():
    """Run the ELISA parser with the enhanced template"""
    source_file = "attached_assets/EK1586_Mouse_KLK1Kallikrein_1_ELISA_Kit_PicoKine_Datasheet.docx"
    enhanced_template = "templates_docx/enhanced_template.docx"
    output_file = "enhanced_output.docx"
    
    # Make sure the template file exists
    if not os.path.exists(enhanced_template):
        print(f"Error: Enhanced template not found at {enhanced_template}")
        return 1
    
    # Run the parser with the enhanced template
    cmd = [
        "python", "main.py",
        "--source", source_file,
        "--template", enhanced_template,
        "--output", output_file
    ]
    
    result = subprocess.run(cmd, check=False)
    
    if result.returncode == 0:
        print(f"Successfully generated enhanced template output at: {output_file}")
        return 0
    else:
        print(f"Error running the parser: {result.returncode}")
        return result.returncode

if __name__ == "__main__":
    sys.exit(main())