# Data-Transformation
This project contains ETL scripts and data pipelines for migrating invoice data from CRM to CRM, ensuring proper formatting and transformation. The process includes:

Invoice Formatting: Converting Salesforce invoice structures to match Odoo’s requirements.
Product Value Modification: Adjusting product price, tax rates, or other attributes during migration.
Batch Import: Efficiently loading transformed data into Odoo in bulk.

Technologies Used
Python (pandas, re)
CSV/XLSX Processing (for data exports and imports)

Prerequisites
Ensure you have the following installed and configured:

Python 3.11
Required libraries (install with):
pip install pandas openpyxl

Input Data: Ensure your Salesforce exports (invoices, products, sales data) are available in Excel (.xlsx) format.
Configuration: Update file paths inside the scripts as needed.
