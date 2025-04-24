# ELISA Kit Datasheet Parser

A specialized Python tool for extracting data from ELISA kit datasheets and generating standardized, formatted output documents.

## Features

- Extracts key information from ELISA kit datasheets in DOCX format
- Populates professional templates with extracted data
- Customizable template system with multiple formatting options
- Web interface for easy use without programming knowledge
- Batch processing capabilities for multiple datasheets
- Password-protected access for secure usage

## Technologies Used

- Python 3.10+
- Flask web framework
- python-docx for DOCX document manipulation
- docxtpl for template processing
- Gunicorn WSGI server
- Bootstrap for the web interface

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/your-username/elisa-kit-parser.git
   cd elisa-kit-parser
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   gunicorn --bind 0.0.0.0:5000 main:app
   ```

## Usage

### Web Interface

1. Access the web interface at `http://localhost:5000`
2. Enter the password to log in (default: "IRelisa2017!")
3. Upload an ELISA kit datasheet in DOCX format
4. Select a template format
5. Optionally provide a custom kit name, catalog number, and lot number
6. Click "Process" to generate the formatted output
7. Download the resulting document

### Command Line

For command-line usage:

```
python elisa_cli.py
```

This will start an interactive CLI that guides you through the extraction process.

## Customizing the Password

Run the included password generator script:

```
python generate_password_hash.py
```

Follow the instructions to create a new password, then add it to your environment variables:

```
export APP_PASSWORD_HASH=your_generated_hash
```

## License

[Insert your license information here]

## Contributing

Contributions welcome! Please feel free to submit a Pull Request.