import numpy as np
import pandas as pd

# Hard coded path, should be input from gui later
DataFolder = r'/media/anders/-Anders-3-/Google Drev/Novozymes/'
#DataFolder = r'C:\Users\310229518\Google Drive\Novozymes\'
database_path = DataFolder + 'Access Novozymes.xlsx'
vilkar_path = DataFolder + '2017 AUG LABTEK TOTAL.xlsx'
hk_path = DataFolder + 'Ba 2017 juni NovoZymes LAB.xlsx'
maillist_path = DataFolder + 'outlook aug 2017.txt'

# Read database excel file
database = pd.read_excel(database_path).dropna(axis=1, how='all')
medlem_bool = database['Medlem'].str.lower() == 'ja'

# Read labktek excel file
vilkar = pd.read_excel(vilkar_path, skiprows=4).dropna(axis=1, how='all')

# Read hk excel file
hk = pd.read_excel(hk_path).dropna(axis=1, how='all')

# Read mailing list txt file
maillist_list = []
with open(maillist_path, 'rb') as f:
    maillist_list.append([word for line in f for word in str(line).split(';')])

maillist_df = pd.DataFrame(np.squeeze(maillist_list))
maillist = maillist_df[0].str.extract(r'<(\w+@\w+.\w+)>', expand=False)

# Compare to vilkar
init_db = set(database['Initialer'].dropna().str.upper())
init_vilkar = set(vilkar['Initials'].dropna().str.upper())
init_onlydb = init_db.difference(init_vilkar)
init_onlyvilkar = init_vilkar.difference(init_db)
print('The following initials are only in NovoZymes database:\n' + '\n'.join(init_onlydb))
print('The following initials are only in Vilkaar database:\n' + '\n'.join(init_onlyvilkar))

# Compare to hk
cpr_db = set(database['CPR-nummer'].dropna().str.zfill(10).str.replace('-', ''))  # [medlem_bool] might be needed
cpr_hk = set(hk['CPR'].dropna().str.replace('-', ''))
cpr_onlyhk = cpr_hk.difference(cpr_db)
cpr_onlydb = cpr_db.difference(cpr_hk)
print('The following cpr are only in NovoZymes database:\n' + '\n'.join(cpr_onlydb))
print('The following cpr are only in HK database:\n' + '\n'.join(cpr_onlyhk))

# Compare mailing list to database based on email
mails_db = set(database['Mailadresse'][medlem_bool].dropna().str.lower())
mails_txt = set(maillist.dropna().str.lower())
mails_del = mails_txt.difference(mails_db)
mails_add = mails_db.difference(mails_txt)
print('The following mails can be deleted from outlook:\n' + '\n'.join(mails_del))
print('The following mails can be added to outlook:\n' + '\n'.join(mails_add))

