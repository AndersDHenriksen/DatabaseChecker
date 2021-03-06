import pandas as pd

# Hard coded path, should be input from gui later
# DataFolder = r'/media/anders/-Anders-3-/Google Drev/Novozymes/'
# #DataFolder = r'C:\Users\310229518\Google Drive\Novozymes\'
# database_path = DataFolder + 'Access Novozymes.xlsx'
# vilkar_path = DataFolder + '2017 AUG LABTEK TOTAL.xlsx'
# hk_path = DataFolder + 'Ba 2017 juni NovoZymes LAB.xlsx'
# maillist_path = DataFolder + 'outlook aug 2017.txt'


def read_db(database_path):
    """ Read database excel file """
    database = pd.read_excel(database_path).dropna(axis=1, how='all')
    medlem_bool = database['Medlem'].str.lower() == 'ja'
    return database, medlem_bool


def read_vilkar(vilkar_path):
    """ Read labktek excel file """
    vilkar = pd.read_excel(vilkar_path, skiprows=0).dropna(axis=1, how='all')
    return vilkar


def read_hk(hk_path):
    """ Read hk excel file """
    hk = pd.read_excel(hk_path).dropna(axis=1, how='all')
    return hk


def read_maillist(maillist_path):
    """ Read mailing list txt file """
    maillist_list = []
    with open(maillist_path, 'rb') as f:
        maillist_list.append([word for line in f for word in str(line).split(';')])

    maillist_df = pd.DataFrame(maillist_list[0])
    maillist = maillist_df[0].str.extract(r'<(\w+@\w+.\w+)>', expand=False)
    return maillist


def compare_vilkar(database, vilkar, vilkar_ignore, print_out=True):
    """ Compare main database to vilkar database """
    init_db = set(database['Initialer'].dropna().str.upper())
    init_vilkar = set(vilkar['Initials'].dropna().str.upper())
    init_vilkar = init_vilkar.difference(set(vilkar_ignore))
    init_onlydb = sorted(list(init_db.difference(init_vilkar)))
    init_onlyvilkar = sorted(list(init_vilkar.difference(init_db)))
    if print_out:
        print('The following initials are only in NovoZymes database:\n' + '\n'.join(init_onlydb))
        print('The following initials are only in Vilkaar database:\n' + '\n'.join(init_onlyvilkar))
    return init_onlydb, init_onlyvilkar


def compare_hk(database, hk, hk_ignore, print_out=True):
    """ Compare main database to HK database"""
    cpr_db = set(database['CPR-nummer'].dropna().str.zfill(10).str.replace('-', ''))
    try:
        cpr_hk = set(hk['CPR'].dropna().str.replace('-', ''))
    except KeyError:
        cpr_hk = set(hk['CPR (Person)'].dropna().str.replace('-', ''))
    cpr_hk = cpr_hk.difference(set(hk_ignore))
    cpr_onlyhk = sorted(list(cpr_hk.difference(cpr_db)))
    cpr_onlydb = sorted(list(cpr_db.difference(cpr_hk)))
    if print_out:
        print('The following cpr are only in NovoZymes database:\n' + '\n'.join(cpr_onlydb))
        print('The following cpr are only in HK database:\n' + '\n'.join(cpr_onlyhk))
    return cpr_onlydb, cpr_onlyhk


def compare_maillist(database, medlem_bool, maillist, print_out=True):
    """ Compare main database to email list """
    mails_db = set(database['Mailadresse'][medlem_bool].dropna().str.lower())
    mails_txt = set(maillist.dropna().str.lower())
    mails_del = sorted(list(mails_txt.difference(mails_db)))
    mails_add = sorted(list(mails_db.difference(mails_txt)))
    if print_out:
        print('The following mails can be deleted from outlook:\n' + '\n'.join(mails_del))
        print('The following mails can be added to outlook:\n' + '\n'.join(mails_add))
    return mails_del, mails_add
