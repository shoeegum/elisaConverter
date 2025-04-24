# ELISA Kit Datasheet Processor

This tool automates the conversion of ELISA kit datasheets into professionally formatted research documents with consistent branding and styling.

## Features

- **Data Extraction**: Extracts key information from ELISA kit datasheets including catalog numbers, intended use, background information, etc.
- **Formatted Output**: Creates well-formatted documents with proper styling, margins, and consistent branding
- **Command-line Interface**: Simple command-line interface for batch processing
- **Web Interface**: User-friendly web interface for non-technical users

## Quick Start

To generate a properly formatted document from an ELISA kit datasheet:

```bash
python complete_parser_and_document_creator.py --source [SOURCE_FILE] --output [OUTPUT_FILE] 
    --kit-name "Mouse KLK1 ELISA Kit" --catalog-number "IMSKLK1KT" --lot-number "20250424"
```

For example:
```bash
python complete_parser_and_document_creator.py --source datasheets/EK1586.docx --output IMSKLK1KT-20250424.docx 
    --kit-name "Mouse KLK1 ELISA Kit" --catalog-number "IMSKLK1KT" --lot-number "20250424"
```

## Web Application

The web interface can be started using:

```bash
gunicorn --bind 0.0.0.0:5000 --reuse-port --reload main:app
```

Then navigate to the provided URL to use the web interface for uploading and processing files.

## Available Scripts

The project includes several utilities:

1. **complete_parser_and_document_creator.py** - A standalone script for parsing ELISA datasheets and creating formatted documents
2. **simple_output_creator.py** - Creates a simple document with basic formatting
3. **create_better_template.py** - Generates a clean template file
4. **generate_from_scratch.py** - Creates a document from scratch using the data from a datasheet

## Document Formatting

Output documents follow a consistent format with:

- Blue section headings (RGB 0,70,180) in Calibri font
- Bold 36pt Calibri for the main title
- Company name in the footer (Calibri 24pt bold)
- Contact information in the footer (Calibri Light 12pt)
- Formatted tables for reagents and technical details
- Bulleted lists for required materials
- Numbered lists for protocol steps

## Output Files

Output files are named using the format `CatalogNumber-LotNumber.docx` (e.g., `IMSKLK1KT-20250424.docx`).

## Troubleshooting

If the main application produces documents that cannot be opened in Microsoft Word, try using the standalone `complete_parser_and_document_creator.py` script which uses a more reliable approach.