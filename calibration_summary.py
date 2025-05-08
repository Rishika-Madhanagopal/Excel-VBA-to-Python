
import pandas as pd

# Load the Excel file
input_excel = r"C:\Users\rishi\OneDrive\Desktop\Aeirtec_Freelance\Excel project\CALIBRATION RAW DATA table for R260335.xlsx"
xls = pd.ExcelFile(input_excel)

# Define channels and statistics
channels = ['Green', 'CY5', 'CY7']
stats = ['Median', 'Mean', 'Std dev', 'Sample size', '95% CI', 'lower range', 'upper range', 'Q1', 'Q3', 'ICR']

summary = []

# Function to extract stat values from each sheet
def extract_values(df, well_name):
    result = {'Well': well_name}
    for channel in channels:
        for stat in stats:
            value = None
            try:
                # Find the row where channel and stat match
                row_match = df[df[channel].astype(str).str.strip().str.lower() == stat.lower()]
                if not row_match.empty:
                    channel_idx = df.columns.get_loc(channel)
                    next_col = df.columns[channel_idx + 1]
                    value = pd.to_numeric(row_match[next_col], errors='coerce').dropna().values[0]
            except Exception:
                pass
            result[f"{channel}_{stat.lower().replace(' ', '_')}"] = value
    return result

# Loop through all sheets
for sheet in xls.sheet_names:
    print(f"[{sheet}] Processing...")
    df = xls.parse(sheet)
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace('%', 'pct').str.replace(' ', '_').str.lower()

    # Standardize channel column names
    for col in df.columns:
        if 'green' in col.lower():
            df.rename(columns={col: 'Green'}, inplace=True)
        elif 'cy5' in col.lower():
            df.rename(columns={col: 'CY5'}, inplace=True)
        elif 'cy7' in col.lower():
            df.rename(columns={col: 'CY7'}, inplace=True)

    # Only process if at least one channel is found
    if any(ch in df.columns for ch in ['Green', 'CY5', 'CY7']):
        summary.append(extract_values(df, sheet))
    else:
        print(f" No fluorescence channels found in {sheet}")

# Convert summary to DataFrame
summary_df = pd.DataFrame(summary)

# Drop the row where the 'Well' column contains "Calibration range summary"
summary_df = summary_df[summary_df['Well'].str.contains('Calibration range summary', case=False) == False]

# Save output
output_file = "calibration_summary_final.xlsx"
summary_df.to_excel(output_file, index=False)
print(f"\n Calibration summary saved to: {output_file}")

