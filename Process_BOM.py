import pandas as pd

# ======= CONFIGURATION =======
INPUT_FILE = "BOM_2.xlsx"
OUTPUT_FILE = "output.xlsx"

# Mapping of output columns:
# Keys = column labels in output file
# Values = (source_column_name, row_offset)
columns_to_extract = {
    'Component Name': ('Description', 0),       # column F, same row
    'Component Weight': ('Weight', 0),          # column Q, same row
    'Material Name': ('Description', 1),        # column F, next row
    'Material Fraction': ('Weight', 1),         # column Q, next row
}
# =============================

print("Loading BOM...")
bom_df = pd.read_excel(INPUT_FILE)

# Determine type of each row (single, paired, skip)
def classify_rows(df):
    classifications = []
    for i in range(len(df)):
        level = df.iloc[i]['Lvl']
        unit = str(df.iloc[i]['U/M']).strip().lower()

        prev_level = df.iloc[i - 1]['Lvl'] if i > 0 else None
        next_level = df.iloc[i + 1]['Lvl'] if i + 1 < len(df) else None
        next_unit = str(df.iloc[i + 1]['U/M']).strip().lower() if i + 1 < len(df) else ''

        if level == 1 and unit in ['pcs', 'm', 'm3']:
            if next_level == 2 and next_unit == 'kg':
                classifications.append('paired')
            else:
                classifications.append('single')
        elif level == 2 and prev_level == 1 and unit == 'kg':
            classifications.append('skip')
        else:
            classifications.append('single')  # fallback
    return classifications

# Add a new classification column
bom_df['row_type'] = classify_rows(bom_df)

# Process rows
output_rows = []
i = 0
while i < len(bom_df):
    row_type = bom_df.iloc[i]['row_type']

    if row_type == 'single':
        row = bom_df.iloc[i]
        data = {}
        for label, (col_name, offset) in columns_to_extract.items():
            if offset == 0:
                data[label] = row[col_name]
            else:
                data[label] = ''
        output_rows.append(data)
        i += 1

    elif row_type == 'paired':
        if i + 1 < len(bom_df):
            component_row = bom_df.iloc[i]
            material_row = bom_df.iloc[i + 1]
            data = {}
            for label, (col_name, offset) in columns_to_extract.items():
                source_row = component_row if offset == 0 else material_row
                data[label] = source_row[col_name]
            output_rows.append(data)
            i += 2
        else:
            print(f"Warning: 'Paired' row at {i} has no next row.")
            i += 1

    elif row_type == 'skip':
        i += 1  # skip

    else:
        print(f"Unknown classification at row {i}, skipping.")
        i += 1

# Write output
output_df = pd.DataFrame(output_rows)
output_df.to_excel(OUTPUT_FILE, index=False)
print(f"Done! Output written to {OUTPUT_FILE}")
