import numpy as np
import pandas as pd

# Hard coded path, should be input from gui later
database_path = r'C:\Users\310229518\Google Drive\Novozymes\Access Novozymes.xlsx'
labtek_path = r'C:\Users\310229518\Google Drive\Novozymes\2017 AUG LABTEK TOTAL.xlsx'
hk_path = r'C:\Users\310229518\Google Drive\Novozymes\Ba 2017 juni NovoZymes LAB.xlsx'
maillist_path = r'C:\Users\310229518\Google Drive\Novozymes\outlook aug 2017.txt'

# Read database excel file
database = pd.read_excel(database_path).dropna(axis=1, how='all')

# Read labktek excel file
labtek = pd.read_excel(labtek_path, skiprows=4).dropna(axis=1, how='all')

# Read hk excel file
hk = pd.read_excel(hk_path).dropna(axis=1, how='all')

# Read mailing list txt file
maillist_list = []
with open(maillist_path) as f:
    maillist_list.append([word for line in f for word in line.split(';')])
maillist_df = pd.DataFrame(np.squeeze(maillist_list))
maillist = maillist_df[0].str.extract(r'<(\w+@\w+.\w+)>', expand=False)

# Compare mailing list to database based on email
new_mails = set(maillist.dropna().str.lower()).difference(set(database['Mailadresse'].dropna().str.lower()))