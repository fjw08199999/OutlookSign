from main import *


class Create:

    """創建輸入視窗"""

    def get_info(self):
        DateFrm['chineseName_var'] = chineseNameEntry.get()
        DateFrm['englishName_var'] = englishNameEntry.get()
        DateFrm['emailAddress_var'] = emailAddressEntry.get()
        DateFrm['companyPhone_var'] = companyPhoneEntry.get()
        DateFrm['mobilePhone_var'] = mobilePhoneEntry.get()

    def clean_entry(self):
        chineseName_var.set('')
        englishName_var.set('')
        emailAddress_var.set('')
        companyPhone_var.set('')
        mobilePhone_var.set('')

