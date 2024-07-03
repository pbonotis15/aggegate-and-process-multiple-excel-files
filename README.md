# Excel File Aggregator and Processor

## Overview
This project is a Python-based tool designed to automate the aggregation and processing of multiple Excel files. It provides a user-friendly interface for selecting input and output directories, reads and concatenates data from multiple Excel sheets, standardizes column names, and removes duplicates based on specified criteria. The processed data is then saved into a new Excel file.

## Key Features
- **Multi-File Aggregation**: Reads multiple Excel files from a user-specified directory and aggregates data by sheet name.
- **Column Standardization**: Standardizes column names for consistency across different data sets.
- **Duplicate Removal**: Removes duplicate entries based on a unique identifier while retaining the entry with the highest specified value.
- **User-Friendly Interface**: Utilizes a graphical user interface (GUI) for easy selection of input and output directories.
- **Output Generation**: Saves the processed and cleaned data into a new Excel file in the user-specified output directory.

## Requirements
- Python 3.x
- pandas
- openpyxl
- tkinter

## Detailed Description

1. **File Reading and Aggregation**:
    - Uses the `glob` module to identify and read all Excel files within the selected directory.
    - Utilizes `pandas` to read and parse each sheet within the Excel files.
    - Aggregates data from sheets with the same name across different files into a single DataFrame per sheet name.

2. **Column Standardization**:
    - Standardizes variations of column names (e.g., `SR`, `sr`, `SRID` to `SR ID`, and `AGECLASS` to `AGE`) to ensure uniformity across all data sets.

3. **Data Processing**:
    - Sorts the data by the unique identifier (`SR ID`) and a specified value (`AGE`) to prioritize entries with the highest value.
    - Removes duplicate entries based on the unique identifier, retaining the entry with the highest specified value.

4. **Output Generation**:
    - Saves the cleaned and processed data into a new Excel file using `pandas`' `ExcelWriter`.
    - Each sheet name in the original data is preserved in the output file.

5. **User Interaction**:
    - Utilizes `tkinter` to provide a GUI for folder selection, making it easy for users to specify input and output directories without manually entering paths.

## Installation
1. Clone the repository:
    ```sh
    git clone https://github.com/pbonotis15/aggregate-and-process-multiple-excel-files.git
    ```
2. Navigate to the project directory:
    ```sh
    cd to_your_folder
    ```
3. Install the required packages:
    ```sh
    pip install pandas openpyxl
    ```

## Usage Instructions

### Setup
- Ensure Python and necessary libraries (`pandas`, `openpyxl`, `tkinter`) are installed.
- Install `pyinstaller` for creating an executable.

### Running the Script
- Run the script directly using Python to test and verify functionality.
- Use `pyinstaller` to create an executable file:
  ```sh
  pyinstaller --onefile aggregate-tpbe.py

## Execution
1. **Launch the executable or run the script.**
2. **Use the GUI to select the folder containing the Excel files for input.**
3. **Select the output folder where the processed Excel file will be saved.**
4. **The script will read, process, and save the data, displaying a confirmation message upon completion.**

## Example Use Case
A company regularly receives Excel reports from multiple departments. Each report contains various sheets with task data, but column names may vary and some tasks are duplicated. This tool can automate the aggregation of these reports, standardize the column names, remove duplicates, and produce a clean, consolidated Excel file for further analysis.

This project simplifies data consolidation, enhances data consistency, and saves time, making it ideal for organizations dealing with large volumes of Excel data from multiple sources.
