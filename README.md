# PDF Extractor Tool

A Python desktop application that extracts PDF hyperlinks from Word documents, organizes them, and creates a searchable index.

![PDF Extractor Tool Screenshot](screenshot.png)

## Description

PDF Extractor Tool helps organize and manage PDF documents that are referenced in Word files. The tool extracts PDF hyperlinks from Word documents, copies the linked PDFs to a destination folder, adds cover pages, and generates an Excel report for easy reference.

Created by Haytham Abo Abdallah for internal company use.

## Features

- Extract PDF hyperlinks from Word documents
- Copy linked PDFs from a central folder to a destination folder
- Add custom cover pages with "المستند رقم" (Document Number) in Arabic
- Rename PDFs with a consistent numbering system
- Create an Excel report with:
  - Clickable links to PDFs
  - Document numbers
  - Status indicators (found/missing)
  - Original link text from the Word document
- Handles duplicate references to the same PDF
- Arabic text support with proper right-to-left rendering

## Installation

### Prerequisites

- Python 3.6 or higher
- Required Python packages:
  ```
  pip install python-docx openpyxl ttkbootstrap reportlab PyPDF2 arabic-reshaper python-bidi
  ```

### Setup

1. Clone this repository:
   ```
   git clone https://github.com/ICEBETA4/PDF-Extractor-Tool.git
   cd pdf-extractor-tool
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python pdf_extractor.py
   ```

## Usage

1. Launch the application
2. Select a Word document containing PDF hyperlinks
3. Choose the central folder where the original PDFs are stored
4. Select a destination folder for the processed files
5. Configure options:
   - Enable/disable renaming with "المستند رقم" prefix
   - Enable/disable cover page addition
6. Click "Process Word File" to start
7. After processing, an Excel report will be generated with links to all PDFs

## Arabic Support

For proper Arabic text rendering, make sure to install the following packages:
```
pip install arabic-reshaper python-bidi
```

The application will work without these packages but may not display Arabic text correctly in the cover pages.

## System Requirements

- Windows, macOS, or Linux
- For Arabic text support on Windows, the following fonts are recommended:
  - Arial
  - Tahoma
  - Arabic Typesetting
  - Simplified Arabic

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Created by Haytham Abo Abdallah
- Built with Python and TTKBootstrap
- Uses python-docx for Word document processing
- Uses reportlab for PDF generation
