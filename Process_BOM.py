import pandas as pd

# ======= CONFIGURATION =======
INPUT_FILE = "BOM.xlsx"
OUTPUT_FILE = "output.xlsx"

# Mapping of output columns:
# Keys = column labels in output file
# Values = (source_column_letter, row_offset)
# row_offset: 0 = current row, 1 = next row (for material, etc.)

columns_to_extract = {
    'Component Name': ('Description', 0),              # column F, same row
    'Component Weight': ('Weight', 0),            # column Q, same row
    'Material Name': ('Description', 1),               # column F, next row
    'Material Fraction': ('Weight', 1),           # column Q, next row
}
# =============================

# Load the BOM file
print("Loading BOM...")
bom_df = pd.read_excel(INPUT_FILE)

# Prepare output rows
output_rows = []
i = 0
while i < len(bom_df):
    row = bom_df.iloc[i]
    p_value = str(row['Description 7']).strip().lower()

    if p_value == 'single':
        data = {}
        for label, (col_name, offset) in columns_to_extract.items():
            if offset == 0:
                data[label] = row[col_name]
            else:
                data[label] = ''  # No material info for single
        output_rows.append(data)
        i += 1

    elif p_value == 'paired':
        if i + 1 < len(bom_df):
            component_row = row
            material_row = bom_df.iloc[i + 1]
            data = {}
            for label, (col_name, offset) in columns_to_extract.items():
                source_row = component_row if offset == 0 else material_row
                data[label] = source_row[col_name]
            output_rows.append(data)
            i += 2  # skip material row next time
        else:
            print(f"Warning: 'Paired' row at {i} has no next row.")
            i += 1

    elif p_value == 'skip':
        i += 1  # skip this row

    else:
        print(f"Warning: Unknown P value '{row['Description 7']}' at row {i}")
        i += 1


# Create output DataFrame and write to Excel
output_df = pd.DataFrame(output_rows)
output_df.to_excel(OUTPUT_FILE, index=False)
print(f"Done! Output written to {OUTPUT_FILE}")

