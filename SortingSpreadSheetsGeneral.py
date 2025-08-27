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

def GetCell(Table,Row,ColumnLetter):
    ColumnLetter = ColumnLetter.lower()
    ColumnNumber = ord(ColumnLetter) - 97
    Result = Table.cell(Row,ColumnNumber).value
    if type(Result) == float and Result.is_integer():
        return int(Result)
    elif type(Result) == datetime.datetime:
        return Result.strftime('%m/%d/%Y')
    elif ColumnLetter == "m":
        return int(Result)
    else:
        return Result

class SortedRowData:
    def __init__(self, Table,Row):
        self.Tiers = GetCell(Table,Row,"A")
        self.Invoice_Number = GetCell(Table,Row,"B")
        self.Date = GetCell(Table,Row,"C")
        self.Reference = GetCell(Table,Row,"D")
        self.Name = GetCell(Table,Row,"E")
        self.Quantity = GetCell(Table,Row,"F")
        self.Num_Ig = GetCell(Table,Row,"G")
        self.Num_sig = GetCell(Table,Row,"H")
        self.Price = GetCell(Table,Row,"I")
        self.HT_Product = GetCell(Table,Row,"J")
        self.Weight_Per_Unit = GetCell(Table,Row,"K")
        self.Total_Weight = GetCell(Table,Row,"L")
        self.Export_Code = GetCell(Table,Row,"M")
        self.Pays = GetCell(Table,Row,"N")

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

        self.WriteToFileButton = ctk.CTkButton(master=frame,text="Select A file to write onto",command=self.FindWriteToFile)
        self.WriteToFileButton.pack(pady=10)

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

    def FindWriteToFile(self):
        File = filedialog.askopenfilename(title="Select A file")
        filename = File.split('/')[len(File.split('/'))-1]
        self.WriteToFileButton.configure(text="Writing to: "+filename)
        global doc2
        global SaveToName
        DuplicatedFile = shutil.copyfile(File,"Sorted File.numbers")
        SaveToName = DuplicatedFile
        doc2 = Document(DuplicatedFile)

    #MAIN FUNCTIONS -------------------------------------------------------------------------------------------------------------------------
        
    def CustomsSorter(self):
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
                if CurrentRow.Quantity != None and CurrentRow.Weight_Per_Unit != None:
                    CurrentRow.Total_Weight = CurrentRow.Weight_Per_Unit * CurrentRow.Quantity
                if CurrentRow.Tiers != None: #Checking if there is data in the row
                    Counter += 1

                    ImportCode = None
                    if ImExportCodes[CurrentRow.Export_Code]:
                        ImportCode = ImExportCodes[CurrentRow.Export_Code]
                    
                    if Previous_Invoice_Number != CurrentRow.Invoice_Number and Previous_Invoice_Number != 0:
                        Table2.write(Counter,8,TotalLeadingToInvoice)
                        TotalLeadingToInvoice = 0
                        Counter+=1
                    
                    Table2.write(Counter,0,CurrentRow.Tiers)
                    Table2.write(Counter,1,CurrentRow.Invoice_Number)
                    Table2.write(Counter,2,CurrentRow.Date)
                    Table2.write(Counter,3,CurrentRow.Reference)
                    Table2.write(Counter,4,CurrentRow.Name)
                    Table2.write(Counter,5,CurrentRow.Quantity)
                    Table2.write(Counter,6,CurrentRow.Num_Ig)
                    Table2.write(Counter,7,CurrentRow.Num_sig)
                    Table2.write(Counter,8,CurrentRow.Price)
                    Table2.write(Counter,9,CurrentRow.HT_Product)
                    Table2.write(Counter,10,CurrentRow.Weight_Per_Unit)
                    Table2.write(Counter,11,CurrentRow.Total_Weight)
                    Table2.write(Counter,12,CurrentRow.Export_Code)
                    Table2.write(Counter,13,ImportCode)
                    Table2.write(Counter,14,CurrentRow.Pays)

                    Previous_Invoice_Number = CurrentRow.Invoice_Number
                    TotalLeadingToInvoice += CurrentRow.Price

        Counter+=1
        Table2.write(Counter,8,TotalLeadingToInvoice)

        doc2.save(SaveToName)
        self.LaunchButton.configure(text="Task completed!")
        

if __name__ == "__main__":
    app = App()
    app.mainloop()

    
        
