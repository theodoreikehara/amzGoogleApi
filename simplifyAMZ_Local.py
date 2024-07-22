import pandas as pd

# Set the base paths for download and desktop directories
download_path = r'C:\Users\tedik\Downloads\\'
desktop_path = r'C:\Users\tedik\OneDrive\Desktop\DA Daily Performance AMZ\\'

# Only the file name is variable
file_name = 'DSP_Overview_Dashboard_MSCO_DUR3_2023-W51.csv'

# Full file paths are constructed by joining base paths with the file name
input_file_path = download_path + file_name
output_file_path = desktop_path + file_name

# Read the CSV data into a DataFrame
df = pd.read_csv(input_file_path)

# Create a new DataFrame with only the 'Delivery Associate' and 'Overall Standing' columns
# Ensure the column names match exactly what's in your CSV file
simplified_df = df[['Delivery Associate ', 'Overall Standing']]

# Save the simplified DataFrame to a new CSV file
simplified_df.to_csv(output_file_path, index=False)