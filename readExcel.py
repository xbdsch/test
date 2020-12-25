# -*- coding: utf-8 -*-
"""
Read Excel
"""
import pandas

class State():
    def __init__(self):
        self.d_state_abbr = {}
        self.d_state_vendor = {}
        self.d_book_in = {}
        self.d_center = {}
        self.d_center["5782"] = "ND16300010"
        self.d_center["3384"] = "NTD7330000"
        self.d_center["4290"] = "C0D7250000"
        self.d_center["4045"] = "U5D8590"  
        
    
    def readStateReference(self, filepath):
        book = pandas.ExcelFile(filepath)
        sheet = book.parse("states")
        for index, row in sheet.iterrows():
            state = row[0].upper()
            state_abbr = row[1].upper()
            vendor = row[2].upper()
            self.d_state_abbr[state] = state_abbr
            self.d_state_vendor[state_abbr] = vendor

    def readBookIn(self, filepath):
        book = pandas.ExcelFile(filepath)
        for sheet_name in book.sheet_names:
            if sheet_name not in ("Original data", "Sheet ready"):
                self.d_book_in[sheet_name] = {}
                sheet = book.parse(sheet_name, converters={'G/L Account':str})
                # current_state = ""
                for index, row in sheet.iterrows():
                    state_raw = row["Jurisdiction/Payee"]
                    account = row["G/L Account"]
                    amount_by_account = row["G/L Account Amount"]
                    amount_total = row["Total Amount"]
                    vendor_num = row["Vendor Number"]                    
                    if pandas.notna(state_raw):
                        state_name = state_raw.upper().split("STATE")[0].strip()
                        try:
                            state_abbr = self.d_state_abbr[" ".join([state_name, "STATE"])]                       
                            current_state = state_abbr
                        except KeyError as msg:
                            print(msg)
                            print("Error! Jurisdiction/Payee of '%s' in SHEET '%s' NOT found in the 'states.xlsx' reference" % (state_raw, sheet_name))
                            current_state = state_raw.upper()                                    
                        self.d_book_in[sheet_name][current_state] = {}
                        self.d_book_in[sheet_name][current_state]["account"] = {}
                    if pandas.notna(amount_total):
                        self.d_book_in[sheet_name][current_state]["amount_total"] = amount_total
                    if pandas.notna(vendor_num):
                        self.d_book_in[sheet_name][current_state]["vendor_num"] = vendor_num
                    if pandas.notna(account):
                        self.d_book_in[sheet_name][current_state]["account"][account] = amount_by_account
    
    def readTemplate(self, filepath):
        self.book_template = pandas.ExcelFile(filepath)         
        
    def generateSheets(self):
        for sheet_name_input in self.d_book_in:
            for state_abbr in self.d_book_in[sheet_name_input]:
                if state_abbr in self.d_state_vendor:
                    self.__generateOneSheet(sheet_name_input, state_abbr)

    def __generateOneSheet(self, sheet_name_input, state_abbr):
        sheet = self.book_template.parse(self.book_template.sheet_names[0]) #must refresh sheet every time to ensure last sheet's data don't save
        for index, row in sheet.iterrows():
            if str(row[0]).strip().lower() == "Pay Code (Company Code)".lower():
                sheet.iloc[index,1] = sheet_name_input
            if str(row[0]).strip().lower() == "Pmt Amount".lower():
                sheet.iloc[index,1] = self.d_book_in[sheet_name_input][state_abbr]["amount_total"]
            if str(row[0]).strip().lower() == "Payment Reason".lower():
                l_value = [state_abbr]
                for each in row[1].split()[1:]:
                    l_value.append(each)
                sheet.iloc[index,1] = " ".join(l_value)
            if str(row[0]).strip().lower() == "Vendor Number".lower():
                sheet.iloc[index,1] = self.d_book_in[sheet_name_input][state_abbr]["vendor_num"]       
            if str(row[0]).strip().lower() == "Vendor Name".lower():
                sheet.iloc[index,1] = self.d_state_vendor[state_abbr]       
            if str(row[0]).strip().lower() == "GL Account (Cost Element)".lower():
                d_account_index = {}
                for j in range(1,4):
                    d_account_index[str(int(row[j]))] = j
            if str(row[0]).strip().lower() == "Cost Center (Internal Order or WBS Element)".lower():
                if "6470000000" in d_account_index:
                    sheet.iloc[index,d_account_index["6470000000"]] = self.d_center[sheet_name_input]
            if str(row[0]).strip().lower() == "Amount".lower():
                for account in self.d_book_in[sheet_name_input][state_abbr]["account"]:
                    if account in d_account_index:
                        sheet.iloc[index,d_account_index[account]] = \
                            self.d_book_in[sheet_name_input][state_abbr]["account"][account]
            if str(row[0]).strip().lower() == "Text".lower():
                l_value = [state_abbr]
                for each in row[1].split()[1:]:
                    l_value.append(each)
                sheet.iloc[index,1] = " ".join(l_value)    
            if str(row[0]).strip().lower() == "Assignment (State Abreviation)".lower():
                sheet.iloc[index,1] = state_abbr
        sheet.to_excel(self.writer, sheet_name="-".join([state_abbr, sheet_name_input]), header=["","","",""], index=False)
                
    def writeFile(self, filepath):
        self.writer = pandas.ExcelWriter(filepath, engine="xlsxwriter")
    
    def closeFile(self):
        self.writer.close()
                
def main():
    filepath_states_reference = "states.xlsx"
    # filepath_in = "C:/Users/us16216/Documents/TZ/Work/SUT/2020/11.2020/Tax pmt JE 1120.xlsx"
    filepath_in = "input.xlsx"
    filepath_template = "template.xlsx"    
    filepath_out = "output.xlsx"
    s = State()
    s.readStateReference(filepath_states_reference)
    # print(s.d_state_abbr)
    s.readBookIn(filepath_in)
    # print(s.d_book_in.keys())
    # print(s.d_book_in["5782"])
    s.readTemplate(filepath_template)
    s.writeFile(filepath_out)
    s.generateSheets()
    s.closeFile()
    
if __name__=='__main__':
	main()
