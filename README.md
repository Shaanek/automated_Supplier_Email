# Automated_Supplier_Email Python BOT for Supply Chain Digitization
The code presented automates the process of sending open Purchase Order (PO) reports to suppliers via email. This is particularly useful in the supply chain industry where companies often have multiple suppliers and need to ensure timely communication. 
Here’s a sample GitHub README file for the described project:

```markdown
# Automated Supplier PO Report Email Bot

This Python-based bot automates the process of sending open Purchase Order (PO) reports to suppliers via email. It simplifies supply chain communication by reducing manual effort and minimizing human error, ensuring timely communication and efficient tracking of open orders.

## **Overview**

The bot performs the following steps:
1. Reads and cleans data from PO reports.
2. Identifies open POs (those still requiring action).
3. Sends customized email reports to suppliers, with attachments of their specific open POs.
4. Automates the entire process using Microsoft Outlook and Python.

## **Technologies Used**

- **Python**: The core language for automation.
- **pandas**: Used for data manipulation and analysis.
- **win32com.client**: Allows Python to interact with Microsoft Outlook for email automation.
- **re**: For regular expression operations (text cleaning).
- **pathlib** & **ntpath**: To handle file paths across different operating systems.

## **Step-by-Step Breakdown**

### **1. Import Required Libraries**

```python
import win32com.client as win32
import pandas as pd
import re
from pathlib import Path
import ntpath as os
```

These libraries help automate Outlook, clean data, and handle file paths.

### **2. Load and Prepare PO Report Data**

```python
df = pd.read_excel(r'F:\Python_Dhanda\Simulated_Auto_Email_Bot\Input Data\PO_Report_Simulated.xlsx')
df.info()
```

The data is loaded into a pandas DataFrame, which is then analyzed for further processing.

### **3. Clean Supplier Names**

```python
for i in range(0, rowrange):
    df['Supplier Name'][i] = re.sub('[^A-Za-z0-9 ]+', '', df['Supplier Name'][i])
```

Special characters are removed from supplier names to ensure consistency.

### **4. Clean Buyer Names**

```python
for i in range(0, rowrange_buyer):
    df['Buyer'][i] = str(df['Buyer'][i])
```

Buyer names are cleaned and standardized for uniformity.

### **5. Date Formatting**

```python
df['Po Creation Date'] = pd.to_datetime(df['Po Creation Date'])
```

Dates are converted to a consistent format for easy filtering.

### **6. Filter Open POs**

```python
df_filter_open = df['PO Qty Due'] > 0
df_openpo = df[df_filter_open]
```

Only open POs (those that require action) are retained.

### **7. Get Unique Suppliers**

```python
df_supplier = df_openpo[['Supplier Name']]
df_supplier_unique = df_supplier.drop_duplicates()
```

A list of unique suppliers is extracted from the filtered POs.

### **8. Load Supplier Email Directory**

```python
df_supp_dir = pd.read_excel(r'F:\Python_Dhanda\Simulated_Auto_Email_Bot\Input Data\Simulated Supplier Emails.xlsx')
```

The email directory for suppliers is loaded for reference.

### **9. Clean Supplier Directory Data**

```python
for i in range(0, rowrange):
    df_supp_dir['Supplier Name'][i] = re.sub('[^A-Za-z0-9 ]+', '', df_supp_dir['Supplier Name'][i])
```

We clean the email directory to remove any special characters from supplier names.

### **10. Convert Email Addresses to String**

```python
email_fields = ['Send_to_mail', 'CC1_mail', 'CC2_mail', 'CC3_mail', 'CC4_mail']
for field in email_fields:
    for i in range(0, rowrange):
        df_supp_dir[field][i] = str(df_supp_dir[field][i])
```

All email addresses are converted to string format for proper handling.

### **11. Merge Supplier Lists**

```python
df_supp_email_raw = pd.merge(df_supplier_unique, df_supp_dir)
df_supp_email = df_supp_email_raw.drop_duplicates()
df_supp_email = df_supp_email.dropna(subset=['Send_to_mail'])
```

The unique supplier names are merged with their corresponding email addresses.

### **12. Send Emails to Suppliers**

```python
for index, row in df_supp_email.iterrows():
    if re.match(r"[^@]+@[^@]+\.[^@]+", row['Send_to_mail']):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        to_address = row['Send_to_mail']
        for cc_field in ['CC1_mail', 'CC2_mail', 'CC3_mail', 'CC4_mail']:
            if re.match(r"[^@]+@[^@]+\.[^@]+", row[cc_field]):
                to_address = to_address + ";" + row[cc_field]
        
        mail.To = to_address
        mail.Subject = 'Open PO Report - ' + row['Supplier Name']
        
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
        <em>Purchase commitments are made by Your Company Name only pursuant to written purchase orders.</em></p>
        """
        
        attachment = 'F:\File_location/' + 'Open_PO_' + row['Supplier Name'] + '.xlsx'
        mail.Attachments.Add(attachment)
        
        mail.Send()
        print(row['Supplier Name'])
```

For each supplier, the bot sends a customized email with the attached open PO report, leveraging Outlook automation.

## **How to Run**

1. Clone the repository to your local machine.
2. Install the required libraries:
   ```bash
   pip install pandas pywin32
   ```
3. Place the necessary input files in the specified directories.
4. Run the script to send automated emails to suppliers.

## **Conclusion**

This bot automates the tedious task of sending open PO reports to suppliers. By using Python and Outlook automation, it streamlines the process, improves accuracy, and reduces the manual effort required in the supply chain.

## **License**

MIT License. See LICENSE file for details.

## **Contact**

For any questions or support, please contact [srekatpure@gmail.com].

```

You can add more details like configuration options, setup instructions, and more based on your exact use case!
