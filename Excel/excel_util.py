
import pandas as pd
import webbrowser

class ease_excel:
    def __init__(self,file_loc,colnames):
        """
        file_loc: Location of the Excel file ../../your_file.xlsx
        colnames: The selection along with the order of the column names
                  you want your excel file to be in.
        """
        self.loc = file_loc
        self.df = pd.read_excel(file_loc)
        self.df = self.df[colnames]
        print("Excel File auto rearranged as:\n {}".format(' | '.join(colnames)))
        
    def export_excel(self):
        """
        Utility to export the rearranged excel file.
        """
        self.df.to_excel('data_new.xlsx',index=False)
        print("Successfully Exported as data_new.xlsx !")
        
    def open_transaction_link_of(self,fname,lname):
        """
        Utility to open any link from the row which satisfies a particular condition.
        My case: If fname and lname are equal, then it'll directly on the link in my browser.
        """
        for i in range(len(self.df)):
            if self.df.iloc[i,:]['First Name'].strip().lower() == fname.strip().lower() and self.df.iloc[i,:]['Last Name'].lower() == lname.lower():
                webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(self.df.iloc[i,:]['Upload Transaction Screenshot'])
                
    def open_id_proof_of(self,fname,lname):
        """
        Similar utility as open_transaction_link_of().
        """
        for i in range(len(self.df)):
            if self.df.iloc[i,:]['First Name'].strip().lower() == fname.strip().lower() and self.df.iloc[i,:]['Last Name'].strip().lower() == lname.strip().lower():
                webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(self.df.iloc[i,:]['Upload ID Proof'])
                
    

 
filename = input("Enter File Name: ")
colnames = input("Enter order of column names seperated with commas").strip().split(',')
for i in range(len(colnames)):
    colnames[i] = colnames[i].strip()

a = ease_excel(filename,colnames)

choice = int(input("1. Export Rearranged Excel File \n2. Open Transaction Screenshot\n3. Open ID Proof\nEnter choice: "))

while choice!= -1:
    if choice>1:
        fname = input("Enter First Name: ")
        lname = input("Enter Last Name: ")
        print("\n\n")  
        if choice == 2:
            a.open_transaction_link_of(fname,lname)
        else:
            a.open_id_proof_of(fname,lname)
    else:
    	a.export_excel()
            
    choice = int(input("1. Export Rearranged Excel File \n2. Open Transaction Screenshot\n3. Open ID Proof\nEnter choice: "))
    

        