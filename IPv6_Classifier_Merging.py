import pandas as pd

# --- CONFIGURATION ---
# The exact file you already generated from the previous script
INPUT_EXCEL_FILE = r"D:\2Work Task\IPv6 Collection\IPv6_Traffic_Family_Counts.xlsx"

# The name of the NEW file it will create with everything on one sheet
OUTPUT_EXCEL_FILE = r"D:\2Work Task\IPv6 Collection\IPv6_Traffic_SingleSheet.xlsx"


# ---------------------

def merge_excel_sheets(input_path, output_path):
    print(f"Reading existing Excel file: {input_path}")

    try:
        # sheet_name=None tells pandas to read EVERY tab and put them in a dictionary
        all_sheets = pd.read_excel(input_path, sheet_name=None)
    except FileNotFoundError:
        print("Error: Could not find the input Excel file. Please check the path!")
        return

    # Combine all the individual sheet DataFrames into one massive DataFrame
    print(f"Found {len(all_sheets)} tabs. Stacking them together...")
    combined_df = pd.concat(all_sheets.values(), ignore_index=True)

    # Sort it nicely by Switch Family and then by Subnet Block
    combined_df = combined_df.sort_values(by=['Switch Family', 'Subnet Block'])

    print("Generating new single-sheet Excel file...")

    # Write it back out to a new Excel file
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='All_Traffic_Data')

        # Auto-adjust the column widths so it looks clean and readable
        worksheet = writer.sheets['All_Traffic_Data']
        for col in worksheet.columns:
            max_len = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
            worksheet.column_dimensions[column_letter].width = max_len + 5

    print(f"Success! 100% of your data is now on a single sheet here:\n{output_path}")


# --- RUN SCRIPT ---
if __name__ == "__main__":
    merge_excel_sheets(INPUT_EXCEL_FILE, OUTPUT_EXCEL_FILE)
