# Excel to CSV Hexadecimal Data Extractor

This Python program processes an input Excel (*.xlsx) file containing multiple worksheets with hexadecimal data and extracts specific columns from each worksheet, generating a separate CSV file for each worksheet.

## Features

- Automatically locates and extracts three consecutive columns containing at least 8-digit hexadecimal numbers from each worksheet
- Ignores any rows above the first row and columns to the left of the first column containing the required data
- Generates a CSV file for each worksheet, named &lt;worksheet_name&gt;.csv
- Saves generated CSV files in a subdirectory named `data` within the script directory
- Logs processing details and any errors encountered in a log file named `getser.log`

## Requirements

- Python 3.x
- openpyxl library
- csv library

## Installation

1. Install Python 3.x: https://www.python.org/downloads/
2. Install the required libraries:

```
pip install openpyxl
```

## Usage

1. Place the Python script (script_name.py) in a directory where you want the generated CSV files to be saved
2. Run the script from the command line with the input Excel file path as an argument:

```
python script_name.py input_file.xlsx
```

The program will process the input file and generate the CSV files in the `data` subdirectory. The `getser.log` file will be created in the same directory as the script, containing processing details and any errors encountered.

## Output

The generated CSV files will have the following column headers:

- generator-serial
- sensor-serial
- pairing-serial

Each CSV file will be named after the corresponding worksheet in the input Excel file and saved in the `data` subdirectory.

The program also prints a brief summary of the extraction status for each worksheet to stdout, indicating whether the extraction was successful or if an error occurred.

## Detailed Description

This Python program takes an input Excel file (*.xlsx) with multiple worksheets containing hexadecimal data and extracts specific columns from each worksheet, generating a separate CSV file for each worksheet. The program is designed to locate and extract three consecutive columns containing at least 8-digit hexadecimal numbers from each worksheet, ignoring any rows above the first row and columns to the left of the first column containing the required data.

The program starts by accepting the input Excel file path as a command-line argument and opens the file using the openpyxl library. It then iterates through all the worksheets in the Excel file, processing them one by one. For each worksheet, the program searches for the first cell containing an 8-digit (or more) hexadecimal number. Once located, it identifies the first column and the two columns to its right as the data columns to be extracted. It then reads all the rows that follow, extracting the values from the identified columns.

The extracted data is saved in a CSV file named <worksheet_name>.csv, with each CSV file containing the following column headers: "generator-serial", "sensor-serial", and "pairing-serial". The generated CSV files are saved in a subdirectory called data within the directory where the script is located. If the data directory does not exist, the program creates it before writing the CSV files.

In addition to generating the CSV files, the program logs processing details and any errors encountered while processing the worksheets in a log file named getser.log. If a worksheet does not have the expected three consecutive columns with hexadecimal data, the program logs an error message and skips that worksheet. The program also prints to stdout a brief summary of the extraction status for each worksheet, indicating whether the extraction was successful or if an error occurred.

The program requires the openpyxl and csv libraries for processing Excel and CSV files, respectively.