# Parent-Teacher-Interviews-Summary-Report

This Python script generates individualised Word documents for each student, containing data from an Excel spreadsheet. Each document includes the student's details in a table format, with specific fields excluded. The output files are saved with a standardized naming convention.

## Prerequisites

Make sure the following packages are installed before running the script:

+  pandas: for handling the Excel file data.
+  openpyxl: for reading .xlsx files in pandas.
+  python-docx: for generating Word documents.

You can install these packages via pip:

bash
pip install pandas openpyxl python-docx

## How It Works
+ Excel File Input: The script reads data from an Excel file where each row corresponds to a student, and the columns contain various details about each student.

+ Excluded Fields: Fields like 'NESA' and 'Home_Room' are excluded from the generated Word document. You can customize the excluded_fields list as necessary.

+ Word Document Generation: A Word document is generated for each student with their data in a table format. The filename of each document is based on the student's last and first names.

+ Document Title: The title of each document can be customized by modifying the ***DOCUMENT TITLE*** placeholder in the script.

+ Table Formatting: A table is created where each row corresponds to a column in the Excel file, except for the excluded fields. The table uses a 'Table Grid' style with a font size of 10pt for table content.

+ Output: The documents are saved to a specified directory.


## Template
This is a template with fictional data. 
[Student_Data_Template.csv](https://github.com/user-attachments/files/17441950/Student_Data_Template.csv)
