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

### Running the Script
Once you've installed the necessary libraries, and placed the script in your project directory, follow these steps to execute the script:

1. Place your **Excel file** (`data.xlsx`) containing the fluorescence data in the same directory as the Python script.
   
2. Open your **command line terminal** and navigate to the directory where your Python script is located. For example:

   ```bash
   cd path/to/your/project
