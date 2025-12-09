# PDF to Word Converter

## Overview
The PDF to Word Converter is a Python application designed to convert PDF files containing Arabic text and images into Word documents while preserving the original page order. This project utilizes various libraries to handle PDF processing and Word document creation.

## Features
- Converts PDF files with Arabic text to Word documents.
- Preserves the original layout and page order of the PDF.
- Extracts images from the PDF and includes them in the Word document.

## Project Structure
```
PDF-to-Word-Converter
├── src
│   ├── index.py          # Entry point of the application
│   ├── converter.py      # Contains PDFConverter class for conversion logic
│   ├── utils.py          # Utility functions for text and image extraction
│   └── types
│       └── __init__.py   # Custom types and data structures
├── requirements.txt      # Project dependencies
└── README.md             # Project documentation
```

## Installation
To set up the project, clone the repository and install the required dependencies:

```bash
git clone <repository-url>
cd PDF-to-Word-Converter
pip install -r requirements.txt
```

## Usage
1. Run the application:
   ```bash
   python src/index.py
   ```
2. Follow the prompts to select the PDF file you wish to convert.
3. The converted Word document will be saved in the same directory as the input PDF.

## Dependencies
The project requires the following Python libraries:
- PyPDF2
- python-docx
- Any other libraries needed for PDF processing and Word document creation.

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.