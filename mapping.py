import pandas as pd

def map_columns_from_excel(file_path, col1, col2):
    df = pd.read_excel(file_path)
    if col1 not in df.columns or col2 not in df.columns:
        raise ValueError(f"Columns {col1} or {col2} not found in the Excel file.")
    mapped_dict = dict(zip(df[col1], df[col2]))

    return mapped_dict

file_path_vendor= 'Vendor Mapping File.xlsx' 
col1_V = 'Vendor Name' 
col2_V = 'VENDOR_ID'

file_path_customer = 'Customer Mapping File.xlsx'
col1_C = 'QuickBooks Customer Name'
col2_C = 'CUSTOMER_ID'

file_path_account= 'Account Mapping File.xlsx'
col1_A = 'QB Account'
col2_A = 'Account'

vendor_map = map_columns_from_excel(file_path_vendor, col1_V, col2_V)
customer_map = map_columns_from_excel(file_path_customer, col1_C,col2_C)
account_map = map_columns_from_excel(file_path_account, col1_A, col2_A)

combine = pd.read_excel("Combine.xlsx")
data = pd.DataFrame()

data['DONOTIMPORT'] = ''
data['JOURNAL'] = ['OBJA'] * len(combine)
data['DATE'] = combine['Date']

combine['description_combined'] = combine.apply(lambda row: row['Transaction Type'] + " - " + row['Name'] if pd.notnull(row['Name']) and row['Name'] != '' else row['Transaction Type'], axis=1)
data['DESCRIPTION'] = combine['description_combined']
combine.drop(columns=['description_combined'], inplace = True) 

data['REFERENCE_NO'] = ''

combine['unique_id'] = combine['Date'].astype(str) + combine['Transaction Type']
combine['LINE_NO'] = combine.groupby('unique_id').cumcount() + 1
data['LINE_NO'] = combine['LINE_NO']

combine.drop(columns=['unique_id'], inplace=True)
data['ACCT_NO'] = combine['Account'].map(account_map)
data['LOCATION_ID'] = ''
data['DEPT_ID'] = ''
data['DOCUMENT'] = combine['Num']
data['MEMO'] = combine['Memo/Description']
data['DEBIT'] = combine['Debit'] - combine['Credit']
data['GLENTRY_CUSTOMERID'] = combine['Customer'].map(customer_map)
data['GLENTRY_VENDORID'] = combine['Vendor'].map(vendor_map)

missing_vendors = combine.loc[~combine['Vendor'].isin(vendor_map.keys()), 'Vendor'].drop_duplicates()
missing_customers = combine.loc[~combine['Customer'].isin(customer_map.keys()), 'Customer'].drop_duplicates()
missing_accounts = combine.loc[~combine['Account'].isin(account_map.keys()), 'Account'].drop_duplicates()

#missing_vendors.to_excel("missing_vendors.xlsx", index=False)
#missing_customers.to_excel("missing_customers.xlsx", index=False)
#missing_accounts.to_excel("missing_accounts.xlsx", index=False)

data.to_excel("new_excel.xlsx", index=False)