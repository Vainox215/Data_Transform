
import pandas as pd

# Load data from Excel files
invoices_df = pd.read_excel(r'C:\\Users\\Smart Rental\\Documents\\RI Export\\Invoic updated.xlsx')
invoice_lines_df = pd.read_excel(r'C:\\Users\\Smart Rental\\Documents\\RI Export\\Invoice Line.xlsx')

# Rename 'Id' in invoice_lines_df to 'Invoice_Line_ID' for clarity in the output
invoice_lines_df = invoice_lines_df.rename(columns={'Id': 'Invoice_Line_ID'})

# Check for duplicates in 'Invoice_Line_ID' and 'Invoice__c' columns and remove them
invoice_lines_df = invoice_lines_df.drop_duplicates(subset=['Invoice_Line_ID', 'Invoice__c'], keep='first')

# Define the columns we want from Invoice_Lines, including 'Invoice Line ID'
desired_columns = [
    'Invoice_Line_ID', 'Invoice__c', 'Stock_Item__c',  
    'Quantity__c', 'Amount__c', 'Selected__c', 'Amount_Local__c'
]

# Retain only the columns that exist in the loaded DataFrame
existing_columns = [col for col in desired_columns if col in invoice_lines_df.columns]
invoice_lines_df = invoice_lines_df[existing_columns]

# Merge the two DataFrames on the invoice ID with all columns from Invoices and existing columns from Invoice_Lines
merged_df = pd.merge(
    invoices_df, 
    invoice_lines_df, 
    left_on='Id', 
    right_on='Invoice__c', 
    how='left'
)

# Initialize a list to hold the final output
output = []

# Group by Invoice ID (which is 'Id' in invoices_df)
for invoice_id, group in merged_df.groupby('Id'):
    # Skip modifying invoices with only one line
    if len(group) > 1:
        # Sort the group by 'Amount__c' in descending order
        group = group.sort_values(by='Amount__c', ascending=False, na_position='last')

        # Check if any Amount__c matches the Invoice_Amount__c
        invoice_amount = group.iloc[0]['Invoice_Amount__c']

        # Identify the row where Amount__c matches Invoice_Amount__c and set its Quantity__c to 0
        group.loc[group['Amount__c'] == invoice_amount, 'Quantity__c'] = 0
    
    # Append rows for the current invoice
    for i in range(len(group)):
        if i == 0:
            # For the first row, add all invoice data and product line data
            row_data = group.iloc[i].tolist()  # Use all data as-is
        else:
            # For remaining product lines, add only product-specific data, leaving invoice fields blank
            row_data = [""] * len(invoices_df.columns) + group.iloc[i][existing_columns].tolist()
        
        output.append(row_data)

# Define columns for the final DataFrame
final_columns = list(invoices_df.columns) + existing_columns

# Create a new DataFrame for the final output
final_df = pd.DataFrame(output, columns=final_columns)

# Save the result to a new Excel file
final_df.to_excel('Final_Invoice_Report_Quantity_Adjusted(final)2.xlsx', index=False)

print("Final report has been created successfully!")

