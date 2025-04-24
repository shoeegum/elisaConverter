"""
Create a proper template with placeholders from the sample document.
"""

import logging
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Set up logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_innovative_template():
    """
    Create a proper Innovative Research template with proper styles and placeholders.
    """
    doc = Document()
    
    # Set document styles
    styles = doc.styles
    
    # Set default font for the entire document
    font = styles['Normal'].font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # Create styles for headers
    header_style = styles.add_style('Header Style', styles['Normal'].type)
    header_style.font.name = 'Calibri'
    header_style.font.size = Pt(12)
    header_style.font.bold = True
    header_style.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
    
    # Add the title at the top
    title_para = doc.add_paragraph()
    title_run = title_para.add_run("{{ kit_name }}")
    title_run.font.name = 'Calibri'
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Add catalog number and lot number
    cat_lot_para = doc.add_paragraph()
    cat_lot_para.add_run("Catalog Number: ").bold = True
    cat_lot_para.add_run("{{ catalog_number }}")
    cat_lot_para.add_run("          Lot Number: ").bold = True
    cat_lot_para.add_run("{{ lot_number }}")
    cat_lot_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Add sections with proper formatting
    
    # Intended Use
    doc.add_paragraph()
    intended_use_heading = doc.add_paragraph("Intended Use:", style='Header Style')
    doc.add_paragraph("{{ intended_use }}")
    doc.add_paragraph()
    
    # Background
    background_heading = doc.add_paragraph("Background:", style='Header Style')
    doc.add_paragraph("{{ background }}")
    doc.add_paragraph()
    
    # Assay Principle
    assay_heading = doc.add_paragraph("Principle of the Assay:", style='Header Style')
    doc.add_paragraph("{{ assay_principle }}")
    doc.add_paragraph()
    
    # Kit Components
    components_heading = doc.add_paragraph("Kit Components:", style='Header Style')
    component_para = doc.add_paragraph()
    component_para.add_run("The following reagents are included:").bold = True
    
    # Add a table for kit components
    comp_table = doc.add_table(rows=1, cols=3)
    comp_table.style = 'Table Grid'
    
    # Add headers to the table
    header_cells = comp_table.rows[0].cells
    header_cells[0].text = "Component"
    header_cells[1].text = "Volume"
    header_cells[2].text = "Storage"
    
    # Make headers bold
    for cell in header_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    # Add placeholder rows with Jinja tags
    for i in range(5):  # We'll add 5 placeholder rows
        row = comp_table.add_row().cells
        row[0].text = f"{{{{ reagents[{i}].name if {i} < reagents|length else '' }}}}"
        row[1].text = f"{{{{ reagents[{i}].volume if {i} < reagents|length else '' }}}}"
        row[2].text = f"{{{{ reagents[{i}].storage if {i} < reagents|length else '' }}}}"
    
    doc.add_paragraph()
    
    # Materials Required
    materials_heading = doc.add_paragraph("Materials Required But Not Supplied:", style='Header Style')
    doc.add_paragraph("{{ required_materials }}")
    doc.add_paragraph()
    
    # Reagent Preparation
    reagent_heading = doc.add_paragraph("Reagent Preparation:", style='Header Style')
    doc.add_paragraph("{{ reagent_preparation }}")
    doc.add_paragraph()
    
    # Sample Collection and Storage
    sample_heading = doc.add_paragraph("Sample Collection and Storage:", style='Header Style')
    doc.add_paragraph("{{ sample_collection_notes }}")
    doc.add_paragraph()
    
    # Assay Procedure
    procedure_heading = doc.add_paragraph("Assay Procedure:", style='Header Style')
    
    # Add each step of the procedure as a numbered list
    for i in range(10):  # We'll add 10 placeholder steps
        para = doc.add_paragraph(style='List Number')
        para.add_run(f"{{{{ assay_protocol[{i}] if {i} < assay_protocol|length else '' }}}}")
    
    doc.add_paragraph()
    
    # Calculation of Results
    calculation_heading = doc.add_paragraph("Calculation of Results:", style='Header Style')
    doc.add_paragraph("{{ data_analysis }}")
    doc.add_paragraph()
    
    # Standard Curve
    curve_heading = doc.add_paragraph("Standard Curve:", style='Header Style')
    doc.add_paragraph("{{ standard_curve_title }}")
    
    # Add a table for standard curve data
    curve_table = doc.add_table(rows=1, cols=2)
    curve_table.style = 'Table Grid'
    
    # Add headers to the curve table
    curve_headers = curve_table.rows[0].cells
    curve_headers[0].text = "Concentration"
    curve_headers[1].text = "OD Value"
    
    # Make headers bold
    for cell in curve_headers:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    # Add placeholder rows with Jinja tags
    for i in range(8):  # We'll add 8 placeholder rows for standard curve
        row = curve_table.add_row().cells
        row[0].text = f"{{{{ standard_curve_table[{i}].concentration if {i} < standard_curve_table|length else '' }}}}"
        row[1].text = f"{{{{ standard_curve_table[{i}].od_value if {i} < standard_curve_table|length else '' }}}}"
    
    doc.add_paragraph()
    
    # Sensitivity
    sensitivity_heading = doc.add_paragraph("Sensitivity:", style='Header Style')
    doc.add_paragraph("{{ sensitivity }}")
    doc.add_paragraph()
    
    # Precision
    precision_heading = doc.add_paragraph("Precision:", style='Header Style')
    
    # Intra-Assay Precision
    doc.add_paragraph("Intra-Assay Precision (Precision within an assay)")
    
    # Intra-Assay Precision table
    intra_table = doc.add_table(rows=1, cols=3)
    intra_table.style = 'Table Grid'
    
    # Add headers to the intra-assay table
    intra_headers = intra_table.rows[0].cells
    intra_headers[0].text = "Sample"
    intra_headers[1].text = "n"
    intra_headers[2].text = "CV%"
    
    # Make headers bold
    for cell in intra_headers:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    # Add placeholder rows with Jinja tags
    for i in range(3):  # We'll add 3 placeholder rows for intra-assay precision
        row = intra_table.add_row().cells
        row[0].text = f"{{{{ intra_precision[{i}].sample if intra_precision and {i} < intra_precision|length else '' }}}}"
        row[1].text = f"{{{{ intra_precision[{i}].n if intra_precision and {i} < intra_precision|length else '' }}}}"
        row[2].text = f"{{{{ intra_precision[{i}].cv if intra_precision and {i} < intra_precision|length else '' }}}}"
    
    doc.add_paragraph()
    
    # Inter-Assay Precision
    doc.add_paragraph("Inter-Assay Precision (Precision between assays)")
    
    # Inter-Assay Precision table
    inter_table = doc.add_table(rows=1, cols=3)
    inter_table.style = 'Table Grid'
    
    # Add headers to the inter-assay table
    inter_headers = inter_table.rows[0].cells
    inter_headers[0].text = "Sample"
    inter_headers[1].text = "n"
    inter_headers[2].text = "CV%"
    
    # Make headers bold
    for cell in inter_headers:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    # Add placeholder rows with Jinja tags
    for i in range(3):  # We'll add 3 placeholder rows for inter-assay precision
        row = inter_table.add_row().cells
        row[0].text = f"{{{{ inter_precision[{i}].sample if inter_precision and {i} < inter_precision|length else '' }}}}"
        row[1].text = f"{{{{ inter_precision[{i}].n if inter_precision and {i} < inter_precision|length else '' }}}}"
        row[2].text = f"{{{{ inter_precision[{i}].cv if inter_precision and {i} < inter_precision|length else '' }}}}"
    
    doc.add_paragraph()
    
    # Specificity
    specificity_heading = doc.add_paragraph("Specificity:", style='Header Style')
    doc.add_paragraph("{{ specificity }}")
    doc.add_paragraph()

    # Save the template
    template_path = Path('templates_docx/innovative_proper_template.docx')
    doc.save(template_path)
    
    logger.info(f"Created template at {template_path}")
    return template_path

if __name__ == "__main__":
    create_innovative_template()