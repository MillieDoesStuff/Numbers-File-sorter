#CODE MADE 100% BY MILLIE

#IMPORTS -------------------------------------------------------------------------------------------------------------------------------------------------------------------

from numbers_parser import Document
import datetime
import tkinter
from tkinter import filedialog
import customtkinter as ctk
import shutil

#SETUP VAIRABLES! -----------------------------------------------------------------------------------------------------------------------------------------------------------

doc = None
doc2 = None
doc3 = None
SaveToName = None

#FUNCTIONS AND CLASSES ------------------------------------------------------------------------------------------------------------------------------------------------------------------

ImportCodeInput = input(str("Type the column letter of the Import codes:")).upper()
ExportCodeInput = input(str("Type the column letter of the Export codes:")).upper()
PriceColumnInput = input(str("Type the column letter of the Prices:")).upper()

print("The app has now opened, this tab is useless now.")

def GetCell(Table,Row,ColumnLetter):
    ColumnLetter = ColumnLetter.lower()
    ColumnNumber = ord(ColumnLetter) - 97
    Result = None
    try:
        Result = Table.cell(Row,ColumnNumber).value
    except:
        Result = None
    
    if type(Result) == float and Result.is_integer():
        return int(Result)
    elif type(Result) == datetime.datetime:
        return Result.strftime('%m/%d/%Y')
    else:
        return Result

class SortedRowData:
    def __init__(self, Table,Row):
        self.A = GetCell(Table,Row,"A")
        self.B = GetCell(Table,Row,"B")
        self.C = GetCell(Table,Row,"C")
        self.D = GetCell(Table,Row,"D")
        self.E = GetCell(Table,Row,"E")
        self.F = GetCell(Table,Row,"F")
        self.G = GetCell(Table,Row,"G")
        self.H = GetCell(Table,Row,"H")
        self.I = GetCell(Table,Row,"I")
        self.J = GetCell(Table,Row,"J")
        self.K = GetCell(Table,Row,"K")
        self.L = GetCell(Table,Row,"L")
        self.M = GetCell(Table,Row,"M")
        self.N = GetCell(Table,Row,"N")
        self.O = GetCell(Table,Row,"O")
        self.P = GetCell(Table,Row,"P")
        self.Q = GetCell(Table,Row,"Q")
        self.R = GetCell(Table,Row,"R")
        self.S = GetCell(Table,Row,"S")

def FindDataFromVairable(VAIRABLE,Table,Row):
    if VAIRABLE.upper() == "A":
        return SortedRowData(Table,Row).A
    elif VAIRABLE.upper() == "B":
        return SortedRowData(Table,Row).B
    elif VAIRABLE.upper() == "C":
        return SortedRowData(Table,Row).C
    elif VAIRABLE.upper() == "D":
        return SortedRowData(Table,Row).D
    elif VAIRABLE.upper() == "E":
        return SortedRowData(Table,Row).E
    elif VAIRABLE.upper() == "F":
        return SortedRowData(Table,Row).F
    elif VAIRABLE.upper() == "G":
        return SortedRowData(Table,Row).G
    elif VAIRABLE.upper() == "H":
        return SortedRowData(Table,Row).H
    elif VAIRABLE.upper() == "I":
        return SortedRowData(Table,Row).I
    elif VAIRABLE.upper() == "J":
        return SortedRowData(Table,Row).J
    elif VAIRABLE.upper() == "K":
        return SortedRowData(Table,Row).K
    elif VAIRABLE.upper() == "L":
        return SortedRowData(Table,Row).L
    elif VAIRABLE.upper() == "M":
        return SortedRowData(Table,Row).M
    elif VAIRABLE.upper() == "N":
        return SortedRowData(Table,Row).N
    elif VAIRABLE.upper() == "O":
        return SortedRowData(Table,Row).O
    elif VAIRABLE.upper() == "P":
        return SortedRowData(Table,Row).P
    elif VAIRABLE.upper() == "Q":
        return SortedRowData(Table,Row).Q
    elif VAIRABLE.upper() == "R":
        return SortedRowData(Table,Row).R
    elif VAIRABLE.upper() == "S":
        return SortedRowData(Table,Row).S

#APP CODE -------------------------------------------------------------------

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")



class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Spreadsheet sorter")
        self.geometry("720x480")
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(0,weight=1)

        #sidebar!
        sidebar = ctk.CTkFrame(self)
        sidebar.grid(row=0,column=0,sticky="ns")
        
        #SidebarTitle = ctk.CTkLabel(sidebar,text="Some tasks").grid(pady=10)
        #ctk.CTkButton(sidebar,text="Customs Speadsheet Sorter").grid(pady=10,padx=10)
        #ctk.CTkButton(sidebar,text="ECAS price scraper").grid(pady=10,padx=10)

        frame = ctk.CTkFrame(self)
        frame.grid(row=0,column=1,sticky="nesw",padx=10,pady=10)

        SorterTitle = ctk.CTkLabel(frame,text="Customs Spreadsheet Sorter")
        SorterTitle.pack(pady=10)
        
        self.ReadFromFileButton = ctk.CTkButton(master=frame,text="Select A file to read from",command=self.FindReadFromFile)
        self.ReadFromFileButton.pack(pady=10)

        self.ImportToExportButton = ctk.CTkButton(master=frame,text="Select the file with import and export numbers",command=self.FindImExportFile)
        self.ImportToExportButton.pack(pady=10)

        self.LaunchButton = ctk.CTkButton(master=frame,text="Run task",command=self.CustomsSorter)
        self.LaunchButton.pack(pady=10)

    #BUTTON FUNCTIONS ----------------------------------------------------------------------------------------------
        
    def FindReadFromFile(self):
        File = filedialog.askopenfilename(title="Select A file")
        filename = File.split('/')[len(File.split('/'))-1]
        self.ReadFromFileButton.configure(text="Reading From: "+filename)
        global doc
        doc = Document(File)

    def FindImExportFile(self):
        File = filedialog.askopenfilename(title="Select A file")
        filename = File.split('/')[len(File.split('/'))-1]
        self.ImportToExportButton.configure(text="Import & Export File: "+filename)
        global doc3
        doc3 = Document(File)

    #MAIN FUNCTIONS -------------------------------------------------------------------------------------------------------------------------
        
    def CustomsSorter(self):
        DuplicatedFile = shutil.copyfile("Template.numbers","Sorted File.numbers") #Copies the Template
        SaveToName = DuplicatedFile
        doc2 = Document(DuplicatedFile)
        
        Table = doc.sheets[0].tables["Table 1"]
        Table2 = doc2.sheets[0].tables["Table 1"]
        Table3 = doc3.sheets[0].tables["Table 1"]
        Previous_Invoice_Number = 0
        Counter = 0
        TotalLeadingToInvoice = 0
        ImExportCodes = {}


        for i,v in enumerate(Table3.rows()):
            if i!= 0:
                ExportCode = GetCell(Table3,i,"A")
                ImportCode = GetCell(Table3,i,"B")
                ImExportCodes[ExportCode] = ImportCode

        for i,v in enumerate(Table.rows()):
            if i != 0:
                CurrentRow = SortedRowData(Table,i)
                print(getattr(CurrentRow, ExportCodeInput))
                if CurrentRow.A != None: #Checking if there is data in the row
                    Counter += 1

                    ImportCode = None
                    if ImExportCodes[FindDataFromVairable(ExportCodeInput,Table,i)]:
                        ImportCode = ImExportCodes[FindDataFromVairable(ExportCodeInput,Table,i)]
                    
                    if Previous_Invoice_Number != CurrentRow.A and Previous_Invoice_Number != 0:
                        Table2.write(Counter,8,TotalLeadingToInvoice)
                        TotalLeadingToInvoice = 0
                        Counter+=1

                    if CurrentRow.A != None:
                        Table2.write(Counter,0,CurrentRow.A)
                    if CurrentRow.B != None:
                        Table2.write(Counter,1,CurrentRow.B)
                    if CurrentRow.C != None:
                        Table2.write(Counter,2,CurrentRow.C)
                    if CurrentRow.D != None:
                        Table2.write(Counter,3,CurrentRow.D)
                    if CurrentRow.E != None:
                        Table2.write(Counter,4,CurrentRow.E)
                    if CurrentRow.F != None:
                        Table2.write(Counter,5,CurrentRow.F)
                    if CurrentRow.G != None:
                        Table2.write(Counter,6,CurrentRow.G)
                    if CurrentRow.H != None:
                        Table2.write(Counter,7,CurrentRow.H)
                    if CurrentRow.I != None:
                        Table2.write(Counter,8,CurrentRow.I)
                    if CurrentRow.J != None:
                        Table2.write(Counter,9,CurrentRow.J)
                    if CurrentRow.K != None:
                        Table2.write(Counter,10,CurrentRow.K)
                    if CurrentRow.L != None:
                        Table2.write(Counter,11,CurrentRow.L)
                    if CurrentRow.M != None:
                        Table2.write(Counter,12,CurrentRow.M)
                    if CurrentRow.N != None:
                        Table2.write(Counter,13,CurrentRow.N)
                    if CurrentRow.O != None:
                        Table2.write(Counter,14,CurrentRow.O)
                    if CurrentRow.P != None:
                        Table2.write(Counter,15,CurrentRow.P)
                    if CurrentRow.Q != None:
                        Table2.write(Counter,16,CurrentRow.Q)
                    if CurrentRow.R != None:
                        Table2.write(Counter,17,CurrentRow.R)
                    if CurrentRow.S != None:
                        Table2.write(Counter,18,CurrentRow.S)

                    Previous_Invoice_Number = CurrentRow.A
                    TotalLeadingToInvoice += FindDataFromVairable(PriceColumnInput,Table,i)
            elif i == 0:
                if GetCell(Table,i,"A") != None:
                    Table2.write(Counter,0,GetCell(Table,i,"A"))
                if GetCell(Table,i,"B") != None:
                    Table2.write(Counter,1,GetCell(Table,i,"B"))
                if GetCell(Table,i,"C") != None:
                    Table2.write(Counter,2,GetCell(Table,i,"C"))
                if GetCell(Table,i,"D") != None:
                    Table2.write(Counter,3,GetCell(Table,i,"D"))
                if GetCell(Table,i,"E") != None:
                    Table2.write(Counter,4,GetCell(Table,i,"E"))
                if GetCell(Table,i,"F") != None:
                    Table2.write(Counter,5,GetCell(Table,i,"F"))
                if GetCell(Table,i,"G") != None:
                    Table2.write(Counter,6,GetCell(Table,i,"G"))
                if GetCell(Table,i,"H") != None:
                    Table2.write(Counter,7,GetCell(Table,i,"H"))
                if GetCell(Table,i,"I") != None:
                    Table2.write(Counter,8,GetCell(Table,i,"I"))
                if GetCell(Table,i,"J") != None:
                    Table2.write(Counter,9,GetCell(Table,i,"J"))
                if GetCell(Table,i,"K") != None:
                    Table2.write(Counter,10,GetCell(Table,i,"K"))
                if GetCell(Table,i,"L") != None:
                    Table2.write(Counter,11,GetCell(Table,i,"L"))
                if GetCell(Table,i,"M") != None:
                    Table2.write(Counter,12,GetCell(Table,i,"M"))
                if GetCell(Table,i,"N") != None:
                    Table2.write(Counter,13,GetCell(Table,i,"N"))
                if GetCell(Table,i,"O") != None:
                    Table2.write(Counter,14,GetCell(Table,i,"O"))
                if GetCell(Table,i,"P") != None:
                    Table2.write(Counter,15,GetCell(Table,i,"P"))
                if GetCell(Table,i,"Q") != None:
                    Table2.write(Counter,16,GetCell(Table,i,"Q"))
                if GetCell(Table,i,"R") != None:
                    Table2.write(Counter,17,GetCell(Table,i,"R"))
                if GetCell(Table,i,"S") != None:
                    Table2.write(Counter,18,GetCell(Table,i,"S"))
        Counter+=1
        Table2.write(Counter,8,TotalLeadingToInvoice)

        doc2.save(SaveToName)
        self.LaunchButton.configure(text="Task completed!")
        

if __name__ == "__main__":
    app = App()
    app.mainloop()
# For anyone reading this YES the code is very sloppy, messy and inefficient however this had 4 hours to be recoded from the ground up with little knowledge on Classes and objects so yeah it's a big mess
    
        
