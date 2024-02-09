# Visual Basic .NET - Relatório Financeiro Automático

## Overview

This Visual Basic .NET script automates the generation of financial reports based on data from an external source (Omie). The script performs various tasks such as data manipulation, creating dynamic tables, and exporting the final report as an Excel file and PDF. Additionally, it includes functionality to handle file overwriting and copying to a designated folder.

## Key Features

- **Data Processing**: Imports and manipulates data from an external Excel file (`dados.xlsx`).
- **Dynamic Table Creation**: Generates a dynamic pivot table and timeline to analyze financial data.
- **Styling and Formatting**: Applies formatting to enhance readability, including bolding headers and formatting numeric values.
- **File Export**: Exports the financial report as an Excel file and PDF.
- **File Management**: Handles overwriting options and copies files to a specified destination folder.

## Usage

1. **Data Preparation**:
    - Ensure the required data file (`dados.xlsx`) is present in the same directory as the script.

2. **Run the Script**:
    - Execute the script, which will perform data processing, create tables, and export the financial report.

3. **Additional Module**:
    - For the script to work correctly, create another module and call the `Exportar()` function.

4. **Review and Modify**:
    - Review the generated report and make any necessary modifications based on your specific requirements.

## Important Note

Before running the script, make sure to create another module and call the `Exportar()` function for it to work correctly.

## Error Handling

The script includes error handling mechanisms to address common issues such as missing data files and file overwriting. If any errors occur, the script will provide appropriate messages to guide the user.

Feel free to customize the script according to your specific needs and data sources.

Happy reporting!
