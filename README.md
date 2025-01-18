# Excel Automation Script: Correcting Prices and Adding Charts

This Python script automates the process of reading Excel spreadsheets, correcting prices in the data, and adding bar charts to visualize the corrected values. It utilizes the openpyxl library to manipulate the Excel files and is designed for efficient processing of multiple spreadsheets.

## Features
- **Price Correction**: Reads prices from an Excel sheet, applies a correction factor (10% discount), and writes the corrected prices to a new column.
- **Bar Chart Generation**: Creates a bar chart to visualize the corrected prices and embeds it into the spreadsheet.
- **Efficient Processing**: Handles large Excel files by iterating over rows and performing operations dynamically.

# Requirements
To use this script, ensure you have the following installed:
- **Python 3.8+**
- **openpyxl library** (install using pip)

**Install the required library**:
- pip install openpyxl

# How to Use

1. **Prepare Your Excel File:**
    - Ensure your Excel file has a sheet named Sheet1.
    - The data should be organized such that the third column (C) contains the original prices.
2. **Place the Script and File:**
    - Place the script and the Excel file you want to process in the same directory.
3. **Run the Script:**
    - Use the following command in the terminal or command prompt:
      - python script_name.py
    - Replace script_name.py with the name of the Python script file.
    - Pass the Excel file name as an argument to the process_workbook function within the script.
4. **Check the Output:**
   - The script updates the original Excel file, adding corrected prices in the fourth column (D) and a bar chart starting from cell E2.
     
# Contributing
If you'd like to contribute to improving this script, feel free to fork the repository, make changes, and submit a pull request.

# Contact
For questions or suggestions, feel free to reach out:
      **Email:** nelsonmmorales9@gmail.com
      **Github:** Sonofneli7

