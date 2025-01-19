### AUTOMATED SUPPLIER PO REPORT EMAIL BOT ###
# Author: Shubham Ekatpure
# Purpose: Automatically send open PO reports to suppliers via Outlook

# SECTION 1: Import Required Libraries
import win32com.client as win32  # For Outlook automation
import pandas as pd  # For data manipulation and analysis
import re  # For regular expression operations (text cleaning)
from pathlib import Path  # For handling file paths cross-platform
import ntpath as os  # For operating system operations

# SECTION 2: Load and Prepare PO Report Data
# Read the PO report Excel file
df = pd.read_excel(r'F:\Input Data\PO_Report_Simulated.xlsx')

# Display information about the dataframe (columns, data types, etc.)
df.info()

# SECTION 3: Clean Supplier Names
# Remove special characters from supplier names for consistent matching
rowrange = len(df['Supplier Name'])
for i in range(0, rowrange):
    df['Supplier Name'][i] = str(df['Supplier Name'][i])  # Convert to string
    # Keep only alphanumeric characters and spaces
    df['Supplier Name'][i] = re.sub('[^A-Za-z0-9 ]+', '', df['Supplier Name'][i])

# SECTION 4: Clean Buyer Names
# Convert buyer names to string format
rowrange_buyer = len(df['Buyer'])
for i in range(0, rowrange_buyer):
    df['Buyer'][i] = str(df['Buyer'][i])

# SECTION 5: Date Formatting
# Convert PO Creation Date to datetime format
# Multiple approaches provided for different date formats
df['Po Creation Date'] = pd.to_datetime(df['Po Creation Date'])

# SECTION 6: Filter Open POs
# Create filter for open POs (where PO Qty Due > 0)
df_filter_open = df['PO Qty Due'] > 0
# Apply filter to get only open POs
df_openpo = df[df_filter_open]

# SECTION 7: Get Unique Suppliers
# Extract supplier names from filtered data
df_supplier = df_openpo[['Supplier Name']]
# Remove duplicate supplier entries
df_supplier_unique = df_supplier.drop_duplicates()

# SECTION 8: Load Supplier Email Directory
# Read Excel file containing supplier email addresses
df_supp_dir = pd.read_excel(r'F:\Input Data\Simulated Supplier Emails.xlsx')

# SECTION 9: Clean Supplier Directory Data
# Clean supplier names in email directory
rowrange = len(df_supp_dir['Supplier Name'])
for i in range(0, rowrange):
    df_supp_dir['Supplier Name'][i] = str(df_supp_dir['Supplier Name'][i])
    df_supp_dir['Supplier Name'][i] = re.sub('[^A-Za-z0-9 ]+', '', df_supp_dir['Supplier Name'][i])

# SECTION 10: Convert Email Addresses to String
# Convert all email fields to string format for consistency
email_fields = ['Send_to_mail', 'CC1_mail', 'CC2_mail', 'CC3_mail', 'CC4_mail']
for field in email_fields:
    rowrange = len(df_supp_dir[field])
    for i in range(0, rowrange):
        df_supp_dir[field][i] = str(df_supp_dir[field][i])

# SECTION 11: Merge Supplier Lists
# Combine unique suppliers with their email information
df_supp_email_raw = pd.merge(df_supplier_unique, df_supp_dir)
# Remove any duplicates after merging
df_supp_email = df_supp_email_raw.drop_duplicates()
# Remove entries with no primary email address
df_supp_email = df_supp_email.dropna(subset=['Send_to_mail'])

# SECTION 12: Send Emails to Suppliers
# Iterate through each supplier and send customized email
for index, row in df_supp_email.iterrows():
    # Verify primary email is valid
    if re.match(r"[^@]+@[^@]+\.[^@]+", row['Send_to_mail']):
        # Create Outlook mail item
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # Build email address list (To and CC)
        to_address = row['Send_to_mail']
        
        # Add CC recipients if they have valid email addresses
        for cc_field in ['CC1_mail', 'CC2_mail', 'CC3_mail', 'CC4_mail']:
            if re.match(r"[^@]+@[^@]+\.[^@]+", row[cc_field]):
                to_address = to_address + ";" + row[cc_field]
        
        # Configure email properties
        mail.To = to_address
        mail.SentOnBehalfOfName = 'scm.team@yourcompany.com'
        mail.Subject = 'Open PO Report - ' + row['Supplier Name']
        
        # Create HTML body with formatted message
        mail.HTMLBody = f"""
        <p>Dear {row['Supplier Name']} Team,</p>
        <p>&nbsp;</p>
        <p>Please find attached a report containing your pending orders.</p>
        <p>Also, kindly prioritize based on the color coding in the sheet (Red - Critical, Orange - High).</p>
        <p>For any updates or queries, please contact the respective buyer.</p>
        <p>&nbsp;</p>
        <p>Thank you,</p>
        <p>Supply Chain Team,</p>
        <p>Your Company Name<br /><br /></p>
        <p><em><strong><u>Suppliers and Vendors:</u></strong></em><br />
        <em>Purchase commitments are made by Your Company Name only pursuant to written purchase orders. 
        Verbal or email discussions are subject to change at any time and are not binding commitments by Your Company Name.</em></p>
        """
        
        # Attach supplier-specific PO report
        direct = 'F:\File_location'
        Sheetname = 'Open_PO_' + row['Supplier Name'] + '.xlsx'
        attachment = direct + Sheetname
        mail.Attachments.Add(attachment)
        
        # Send the email
        mail.Send()
        to_address = ''  # Reset address list for next iteration
        print(row['Supplier Name'])  # Print confirmation of sent email

### End of Script ###
