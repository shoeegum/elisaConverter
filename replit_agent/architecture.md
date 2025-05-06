# Architecture Overview - ELISA Kit Datasheet Parser

## Overview

The ELISA Kit Datasheet Parser is a specialized document processing application designed to extract structured data from ELISA (Enzyme-Linked Immunosorbent Assay) kit datasheets in DOCX format and generate standardized output documents based on templates. The application serves as a document transformation tool for scientific/medical documentation, providing both web-based and desktop interfaces.

The primary goal of the application is to standardize and automate the extraction of technical specifications, protocols, and other metadata from various manufacturer formats into a consistent, branded output format that can be used by researchers and laboratory staff.

## System Architecture

The application is built with a modular architecture that consists of the following main components:

1. **Core Parser Engine** - The central component that analyzes DOCX files and extracts structured data
2. **Template Processor** - Manages document templates and populates them with extracted data
3. **Web Interface** - Flask-based web application for browser access
4. **CLI Interface** - Command-line interface for script-based automation
5. **GUI Interface** - PyQt5-based desktop application
6. **Batch Processor** - Handles processing of multiple documents in sequence

These components work together to provide multiple interfaces to the same underlying document processing capabilities, allowing for flexibility in deployment and usage scenarios.

```
+----------------+    +----------------+    +----------------+
|  Web Interface |    | CLI Interface  |    | GUI Interface  |
|    (Flask)     |    |                |    |    (PyQt5)     |
+-------+--------+    +-------+--------+    +-------+--------+
        |                     |                     |
        v                     v                     v
+-------+---------------------+---------------------+--------+
|                      Core Parser Engine                    |
|                  (ELISADatasheetParser)                    |
+-------+---------------------+---------------------+--------+
        |                     |                     |
        v                     v                     v
+-------+--------+    +-------+--------+    +-------+--------+
| Template       |    | Batch          |    | Document       |
| Processor      |    | Processor      |    | Storage        |
+----------------+    +----------------+    +----------------+
```

## Key Components

### Core Parser Engine (elisa_parser.py)

The `ELISADatasheetParser` class forms the foundation of the application, responsible for:

- Parsing DOCX files using the python-docx library
- Extracting structured data from various sections (technical details, overview, reagents, etc.)
- Identifying and extracting tables, lists, and other formatted content
- Normalizing and cleaning data from different manufacturer formats

This component is designed to be independent of the interface, allowing it to be used from web, CLI, or GUI contexts.

### Template Processor (template_populator_enhanced.py, updated_template_populator.py)

The template processing components handle:

- Managing document templates with placeholders (using docxtpl library)
- Populating templates with extracted data
- Formatting text and tables according to specified styles
- Generating final output documents

The system supports multiple template formats (enhanced_template.docx, red_dot_template.docx, etc.) to accommodate different output styling requirements.

### Web Interface (app.py)

The Flask-based web interface provides:

- Password-protected access to the application
- File upload and processing capabilities
- Template selection options
- Download of processed output documents
- Batch processing capabilities

The web interface is designed to be accessible to non-technical users who need to process documents without programming knowledge.

### CLI & GUI Interfaces (elisa_cli.py, elisa_gui.py)

Alternative interfaces for different use cases:

- CLI: Scriptable interface for automation and integration with other systems
- GUI: Desktop application built with PyQt5 for users who prefer a native application experience

### Batch Processor (batch_processor.py)

Handles processing of multiple documents, with:

- Concurrent processing using thread pools
- Progress tracking
- Aggregated results reporting

## Data Flow

1. **Input** - User uploads ELISA kit datasheet(s) in DOCX format via web, CLI, or GUI interface
2. **Parsing** - Core parser extracts structured data from the document
3. **Template Selection** - User selects or provides a template for the output
4. **Data Population** - Template processor fills the template with extracted data
5. **Output Generation** - Final document is generated and provided to the user
6. **Storage** - Input and output files are temporarily stored in designated folders

```
+---------------+    +----------------+    +-------------------+
| User uploads  | -> | Parser extracts| -> | Template selected |
| DOCX datasheet|    | structured data|    | by user           |
+---------------+    +----------------+    +-------------------+
                                                    |
+---------------+    +----------------+    +-------------------+
| User downloads| <- | Final document | <- | Template populated|
| result        |    | generated      |    | with data         |
+---------------+    +----------------+    +-------------------+
```

## Data Storage

The application uses a simple file-based storage system with organized directories:

- `/uploads` - Temporary storage for uploaded input files
- `/outputs` - Generated output documents
- `/templates_docx` - Document templates
- `/attached_assets` - Static assets used in templates
- `/batch_outputs` - Results from batch processing operations

No database is currently used, as the application operates primarily on a per-request basis without requiring persistent data storage beyond the file system.

## Authentication and Authorization

The web interface implements a simple password-based authentication mechanism:

- Password is stored as a hash (SHA1) in an environment variable or falls back to a default development hash
- Session management is handled through Flask sessions
- No multi-user functionality or role-based authorization is implemented

## External Dependencies

The application relies on several key external libraries:

1. **python-docx** - For parsing and manipulating DOCX files
2. **docxtpl** - Template-based document generation with Jinja2-like syntax
3. **Flask** - Web framework for the browser interface
4. **PyQt5** - GUI framework for the desktop interface
5. **gunicorn** - WSGI HTTP server for production deployment of the web interface

## Deployment Strategy

The application is designed to be deployable in multiple ways:

1. **Web Application** - Deployed using Gunicorn as a WSGI server
2. **Desktop Application** - Packaged as a standalone PyQt5 application
3. **CLI Tool** - Used as a command-line utility in scripted environments

The repository includes a `.replit` file indicating it can be deployed on Replit's platform, with Gunicorn configured to serve the Flask application.

The deployment configuration in the Replit file indicates:
- Python 3.11 as the runtime
- OpenSSL, PostgreSQL and unzip as system dependencies
- Gunicorn as the WSGI server, binding to 0.0.0.0:5000

## Security Considerations

Several security aspects are addressed in the architecture:

1. **Authentication** - Password-protected web interface
2. **Secure Sessions** - HTTP-only cookies with configurable lifetime
3. **File Validation** - Input file type checking and validation
4. **Environment Variables** - Sensitive configuration stored in environment variables

## Limitations and Future Considerations

1. **Scalability** - The current file-based storage approach has limitations for high-volume scenarios
2. **Multi-user Support** - No built-in user management or access control beyond a single password
3. **Error Handling** - While basic error handling exists, a more robust error management system could be implemented
4. **Testing** - No automated testing framework is evident in the repository

## Conclusion

The ELISA Kit Datasheet Parser is designed as a document transformation tool with multiple interfaces to accommodate different usage scenarios. Its modular architecture allows for separation of concerns between parsing, template management, and user interfaces, while maintaining a consistent core functionality across different interaction methods.