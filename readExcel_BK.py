# -*- coding: utf-8 -*-
"""
Read Excel
"""
import sys 
import pandas

def readSheetIn(book, sheet_name):
    sheet = book.parse("5782",converters={'G/L Account':str})
    d_ = {}
    current_state = ""
    for index, row in sheet.iterrows():
        state = row["Jurisdiction/Payee"]
        account = row["G/L Account"]
        amount_by_account = row["G/L Account Amount"]
        amount_total = row["Total Amount"]
        vendor_num = row["Vendor Number"]
        
        if pandas.notna(state):
            current_state = state
            d_[current_state] = {}
            d_[current_state]["account"] = {}
        if pandas.notna(amount_total):
            d_[current_state]["amount_total"] = amount_total
        if pandas.notna(vendor_num):
            d_[current_state]["vendor_num"] = vendor_num
        if pandas.notna(account):
            d_[current_state]["account"][account] = amount_by_account
    return d_

def readBookIn(filepath):
    book = pandas.ExcelFile(filepath)
    d_book = {}
    for sheet_name in book.sheet_names:
        if sheet_name not in ("Original data", "Sheet ready"):
            d_book[sheet_name] = readSheetIn(book, sheet_name)
    return d_book

def generateOneSheet(sheet, d_book):
    for index, row in sheet.iterrows():
        if row[0] == "Pay Code (Company Code)":
            sheet.iloc[index,1] = "5782"
        if row[0] == "Pmt Amount":
            sheet.iloc[index,1] = d_book["5782"]['GEORGIA STATE']["amount_total"]
        if row[0] == "Payment Reason":
            sheet.iloc[index,1] = " ".join(["GA"," SUT Tax PMT"])
        if row[0] == "Vendor Number":
            sheet.iloc[index,1] = d_book["5782"]['GEORGIA STATE']["vendor_num"]       
        if row[0] == "Vendor Name":
            sheet.iloc[index,1] = "GEORGIA DEPARTMENT OF REVENUE"       
        if row[0] == "GL Account (Cost Element)":
            d_account_index = {}
            for j in range(1,4):
                d_account_index[str(int(row[j]))] = j
        if row[0] == "Amount":
            for account in d_book["5782"]['GEORGIA STATE']["account"]:
                print(account, d_book["5782"]['GEORGIA STATE']["account"][account])
                if account in d_account_index:
                    sheet.iloc[index,d_account_index[account]] = \
                        d_book["5782"]['GEORGIA STATE']["account"][account]
        if row[0] == "Text":
            sheet.iloc[index,1] = " ".join(["GA"," SUT Tax PMT"])
    return sheet

def main():
    # filepath_in = "C:/Users/us16216/Documents/TZ/Work/SUT/2020/11.2020/Tax pmt JE 1120.xlsx"
    filepath_in = "C:/Users/us16216/Documents/TZ/Work/SUT/2020/11.2020/test_in.xlsx"
    filepath_template = "template.xlsx"    
    filepath_out = "test_out.xlsx"
    
    d_book = readBookIn(filepath_in)
    for sheet_name in d_book:
        print(sheet_name)
        print(d_book[sheet_name])
    
    book = pandas.ExcelFile(filepath_template)
    sheet = book.parse("GA")
    sheet = generateOneSheet(sheet, d_book)
    
    writer = pandas.ExcelWriter(filepath_out, engine="xlsxwriter")
    sheet.to_excel(writer, sheet_name="GA" + "-5782", header=["","","",""], index=False)
    writer.close()
  
if __name__=='__main__':
	main()
