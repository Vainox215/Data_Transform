'''
import pandas as pd

# Load each sheet into a DataFrame
invoice_line_df = pd.read_excel(r'C:\\Users\\Smart Rental\\Documents\\test\\Invoice Line Item(RI).xlsx')
stock_df = pd.read_excel(r'C:\\Users\\Smart Rental\\Documents\\RI Export(100 Sample)\\Stock_Item__c.xlsx')

# Step 1: Find the maximum price for each stock item in the invoice line sheet
# Group by 'Stock_Item__c' and get the maximum 'Price' for each unique stock item
max_price_per_stock_item = invoice_line_df.groupby('Stock_Item__c')['Amount_Local__c'].max().reset_index()

# Step 2: Merge the max prices with the stock data
# Assuming 'Id' in stock_df matches 'Stock_Item__c' in invoice_line_df
stock_with_max_price = pd.merge(stock_df, max_price_per_stock_item, left_on='Id', right_on='Stock_Item__c', how='left')

# Rename columns for clarity
stock_with_max_price.rename(columns={'Amount_Local__c': 'MaxPrice'}, inplace=True)

# Drop 'Stock_Item__c' column if it's not needed in the final output
stock_with_max_price.drop(columns=['Stock_Item__c'], inplace=True)

# Step 3: Save the result to a new Excel file
output_file = 'stock_with_max_price.xlsx'  # Specify the output file path
stock_with_max_price.to_excel(output_file, index=False)

print(f"The output has been saved to {output_file}")
'''
import pandas as pd

# Load each sheet into a DataFrame
invoice_line_df = pd.read_excel(r'C:\\Users\\Smart Rental\\Documents\\RI Export (Odoo Finale)\\Invoice line(Dec).xlsx')
stock_df = pd.read_excel(r'C:\\Users\\Smart Rental\\Documents\\RI Export (Odoo Finale)\\Stock item (Dec).xlsx')

# Step 1: Find the most frequent price for each stock item in the invoice line sheet
# Group by 'Stock_Item__c' and calculate the mode for the 'Price' column
# We use lambda to handle cases where there may be multiple modes, keeping only the first mode for simplicity
most_frequent_price_per_stock_item = (
    invoice_line_df.groupby('Stock_Item__c')['Amount_Local__c']
    .agg(lambda x: x.mode().iloc[0])  # Select the first mode if there are multiple
    .reset_index()
)

# Step 2: Merge the most frequent prices with the stock data
# Assuming 'Id' in stock_df matches 'Stock_Item__c' in invoice_line_df
stock_with_most_frequent_price = pd.merge(stock_df, most_frequent_price_per_stock_item, left_on='Id', right_on='Stock_Item__c', how='left')

# Rename columns for clarity
stock_with_most_frequent_price.rename(columns={'Amount_Local__c': 'MostFrequentPrice'}, inplace=True)

# Drop 'Stock_Item__c' column if it's not needed in the final output
stock_with_most_frequent_price.drop(columns=['Stock_Item__c'], inplace=True)

# Step 3: Save the result to a new Excel file
output_file = 'stock_with_most_frequent_price.xlsx'  # Specify the output file path
stock_with_most_frequent_price.to_excel(output_file, index=False)

print(f"The output has been saved to {output_file}")
