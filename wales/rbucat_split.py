import pandas as pd

# Read excel without header
df = pd.read_excel("your_file.xlsx", sheet_name="Engine", header=None)

# -------------------------------------------------
# STEP 1: Detect actual header row dynamically
# -------------------------------------------------

header_row = None

for i in range(len(df)):
    row_values = df.iloc[i].astype(str).str.upper().tolist()

    # Check for expected column names
    if "PART NO." in row_values and "DESCRIPTION" in row_values:
        header_row = i
        break

print("Header row:", header_row)

# Set headers
headers = df.iloc[header_row]

# Data after header
data_df = df.iloc[header_row + 1:].copy()
data_df.columns = headers
data_df = data_df.reset_index(drop=True)

# -------------------------------------------------
# STEP 2: Find separator row dynamically
# -------------------------------------------------

separator_index = None

for i in range(len(data_df)):
    row = data_df.iloc[i].astype(str)

    # Example conditions
    if row.str.contains("MANDATORY", case=False).any():
        separator_index = i
        break

print("Separator row:", separator_index)

# -------------------------------------------------
# STEP 3: Split into 2 DataFrames
# -------------------------------------------------

material_df = data_df.iloc[:separator_index].reset_index(drop=True)

labor_df = data_df.iloc[separator_index + 1:].reset_index(drop=True)

# Remove fully empty rows
material_df = material_df.dropna(how="all")
labor_df = labor_df.dropna(how="all")

print(material_df.head())
print(labor_df.head())