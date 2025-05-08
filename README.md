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





## Input Data
The script expects an Excel file (`data.xlsx`) with multiple sheets. Each sheet should contain data for fluorescence measurements and should include columns corresponding to the channels: **Green**, **CY5**, and **CY7**.

The data in each sheet should follow this structure:
- The channels (**Green**, **CY5**, **CY7**) are represented as columns.
- Each row corresponds to a specific well, and the statistics (e.g., Median, Mean, etc.) are provided in one of the rows.

## Running the Script
1. Place your Excel file (`data.xlsx`) in the same directory as the Python script.
2. Run the Python script:

   ```bash
   python extract_calibration_summary.py

