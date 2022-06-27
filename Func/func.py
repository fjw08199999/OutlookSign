from main import *


def clean_entry():

    """
    清除輸入視窗之數值
    """

    chineseName_var.set('')
    englishName_var.set('')
    emailAddress_var.set('')
    companyPhone_var.set('')
    mobilePhone_var.set('')


def create_label(name: str, row: int, column: int):
    tk.Label(root, text=name, font=('caliber', 10, 'bold')).grid(row=row, column=column)


def create_entry(var: str, row: int, column: int):
    tk.Entry(root, textvariable=var, font=('caliber', 10, 'bold')).grid(row=row, column=column)
