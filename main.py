import tkinter as tk
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox

class ExcelProcessor: 
    def __init__(self):
        self.df = None

    def openFile(self): 
        filePath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filePath: 
            try: 
                self.df = pd.read_excel(filePath)
                self.filePath = filePath
            except pd.errors.ParserError: 
                messagebox.showerror("Error", "Please open a valid Excel File")
    
    def initHomeMenu(self): 
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate buttons
        openButton = tk.Button(root, text="Open Excel file", command=ep.openFile)
        openButton.pack()
        insertButton = tk.Button(root, text="Insert Data", command=ep.insertData)
        insertButton.pack()

    def handleInsertData(self, colNames, inputEntries): 
        data = {}

        for colName in colNames:
            entryValue = inputEntries[colName].get()
            data[colName] = entryValue

        # insert data to the DataFrame
        newDf = pd.DataFrame (data, index=[0]) 
        self.df = pd.concat([self.df, newDf], ignore_index=True) 
        print("Data inserted successfully!")
        print(self.df)

        # save updated DataFrame to Excel File
        self.df.to_excel(self.filePath, index=False)
        print("Excel file updated!")
        
        # clear form
        for entry in inputEntries.values():
            entry.delete(0, tk.END)

    def insertData(self): 
        if self.df is None: 
            print("Please open an excel file")
            return 
        
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate form
        colNames = self.df.columns.tolist()
        inputEntries = {}
        for colName in colNames:
            label = tk.Label(self.root, text=colName)
            entry = tk.Entry(self.root)
            label.pack()
            entry.pack()    
            
            inputEntries[colName] = entry

        # generate buttons
        submitBtn = tk.Button(self.root, text="Submit", command=lambda: self.handleInsertData(colNames, inputEntries))
        cancelBtn = tk.Button(self.root, text="Cancel", command=self.initHomeMenu)
        submitBtn.pack()
        cancelBtn.pack()

root = tk.Tk()
root.minsize(width=400, height=300)

ep = ExcelProcessor()
ep.root = root
ep.initHomeMenu()

root.mainloop()