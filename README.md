# Calibration Summary Extraction from Fluorescence Data

## Project Overview
This project is designed to extract and summarize calibration data from fluorescence measurements stored in an Excel file. The script processes the Excel file containing data for multiple wells, extracting statistical values (e.g., Median, Mean, Standard Deviation) for three fluorescence channels: Green, CY5, and CY7. The extracted data is then cleaned, processed, and compiled into a final summary file.

## Technologies Used
- **Python** (Pandas, openpyxl)
- **Excel** (input data format)

## Features
- Reads multiple sheets from an Excel file.
- Extracts statistics (e.g., Mean, Median, 95% CI) for each of the fluorescence channels: Green, CY5, and CY7.
- Standardizes column names to make the data easier to process.
- Generates a clean summary table and exports it to a new Excel file.
- Skips rows related to "Calibration range summary."

## Installation and Setup

### Prerequisites
You need to have Python installed on your machine. You can download and install Python from [here](https://www.python.org/downloads/).

#### Required Libraries:
1. Pandas: A powerful data manipulation and analysis library.
2. Openpyxl: A library for reading and writing Excel files.

You can install the required dependencies using pip:

```bash
pip install pandas openpyxl


# Fluorescence Calibration Summary Extraction Script

## Overview
This Python script processes fluorescence measurement data from an Excel file, extracts relevant statistics from each sheet, and generates a summarized report. The report includes statistics for each channel (Green, CY5, and CY7) and is saved to a new Excel file.

## Input Data
The script expects an Excel file (`data.xlsx`) with multiple sheets. Each sheet should contain data for fluorescence measurements with the following structure:

- **Channels**: Green, CY5, and CY7 are represented as columns.
- **Rows**: Each row corresponds to a specific well.
- **Statistics**: The rows should contain various statistics (e.g., Median, Mean, etc.) related to each channel.

### Expected Columns:
- Green
- CY5
- CY7

### Example Structure:

| Well | Green | CY5  | CY7  |
|------|-------|------|------|
| A1   | 100   | 200  | 300  |
| A2   | 110   | 210  | 310  |
| ...  | ...   | ...  | ...  |

## Running the Script

1. Place your Excel file (`data.xlsx`) in the same directory as the Python script.
   
2. To run the script, execute the following command in the terminal:

   ```bash
   python extract_calibration_summary.py
