import pandas as pd

def split_excel_to_multiple_sheets(input_file, output_file, rows_per_sheet=2000):
    # Read the entire Excel file into a DataFrame
    data = pd.read_excel(input_file)
    
    # Calculate the number of sheets needed
    num_sheets = (len(data) // rows_per_sheet) + (1 if len(data) % rows_per_sheet != 0 else 0)
    
    # Create a Pandas Excel writer object to write multiple sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i in range(num_sheets):
            # Define the start and end rows for the current sheet
            start_row = i * rows_per_sheet
            end_row = min((i + 1) * rows_per_sheet, len(data))
            
            # Slice the DataFrame for the current sheet
            sheet_data = data.iloc[start_row:end_row]
            
            # Write the sliced data to a new sheet in the Excel file
            sheet_name = f'Sheet_{i + 1}'
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Data split into {num_sheets} sheets and saved to '{output_file}'")

# Example usage
input_file = (r'C:\\Users\\Smart Rental\\Documents\\RI Export (Odoo Finale)\\Contact(Dec).xlsx')  # Replace with the path to your input Excel file
output_file = 'output_file.xlsx'  # Replace with the path to your output Excel file
split_excel_to_multiple_sheets(input_file, output_file)
print("sucessful")
