# NLP-Text-Table_Extraction
Text/Table Extractor from Docx/pdf/excel to Excel
# Table Extractor from Docx to Excel

## Overview

This Python script extracts tables from a Word document (.docx) and saves them into separate sheets in an Excel file. It also extracts the text from paragraphs in the document and saves it to an 'Text' sheet in the same Excel file.

## Prerequisites

- Python 3.x
- Install required Python packages using:

  ```bash
  pip install pandas openpyxl python-docx

  Usage
Clone the repository or download the script docx_to_excel.py.

Place your .docx file in the same directory as the script.

Open a terminal and navigate to the script's directory.

Run the script using:
python docx_to_excel.py
Replace your_docx_file.docx with the actual name of your Word document.

The script will generate an Excel file (output_excel_file.xlsx) in the same directory, containing separate sheets for each table and a 'Text' sheet for extracted text.

Script Explanation
The script uses the python-docx library to parse the Word document.
Tables are extracted from the document and saved into separate sheets in an Excel file using the pandas library.
The script also extracts text from paragraphs and saves it to a 'Text' sheet in the Excel file.
Column headers in each table are ensured to be unique to prevent duplication.
Customization
You can modify the script to suit your specific needs, such as changing the output file name or adding additional features.
License
This project is licensed under the MIT License - see the LICENSE.md file for details.
