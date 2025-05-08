
import pandas as pd

# Load the full summary
df = pd.read_excel("calibration_summary_final.xlsx")

# Filter only 'Well' + lower/upper range columns for each channel
range_cols = ['Well'] + [col for col in df.columns if 'lower_range' in col or 'upper_range' in col]

# Create the simplified DataFrame
range_only_df = df[range_cols]

# Save to a new file
output_file = "calibration_range_summary.xlsx"
range_only_df.to_excel(output_file, index=False)


print(f"\n Final range-only summary saved to: {output_file}")
