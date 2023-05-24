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
                print("Openning file...")
                self.df = pd.read_excel(filePath)
                self.filePath = filePath
                print(self.df)
            except pd.errors.ParserError: 
                messagebox.showerror("Error", "Please open a valid Excel File")

    def handleFuzzySearch(self, entry, screen):
        screen.delete('1.0', tk.END)

        res = pd.DataFrame([])
        inputData = entry.get()
        print("Searching " + inputData + "...")
        if (inputData == ""): 
            print("No specific condition, printing all data...")
            inputData = "all rows"
            res = self.df
        else: 
            print("Searching " + inputData)
            colNames = self.df.columns.tolist()
            # search each columns
            for colName in colNames: 
                matchingRows = self.df[self.df[colName].astype(str).str.contains(inputData, case=False)]
                print(colName)
                print(matchingRows)
                res = pd.concat([res, matchingRows], ignore_index=True)

        dfStr = res.to_string(index=False)
        screen.insert(tk.END, dfStr)
    
    def accurateSearch(self): 
        # impl accurate search
        return None

    def search(self):
        if self.df is None: 
            print("Please open an excel file")
            return 
        print("Search data")
        
        # clear GUI
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate layout
        entry = tk.Entry(root)
        scrollbar = tk.Scrollbar(root)
        text = tk.Text(root, yscrollcommand=scrollbar.set)
        fuzzySerachButton = tk.Button(root, text="Serach", command=lambda: ep.handleFuzzySearch(entry, text))
        accurateSearchButton = tk.Button(root, text="Accuarte Serach", command=ep.accurateSearch)
        cancelButton = tk.Button(self.root, text="Cancel", command=self.initHomeMenu)

        entry.pack()
        fuzzySerachButton.pack()
        accurateSearchButton.pack()
        cancelButton.pack()
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(side=tk.LEFT, fill=tk.BOTH)
        scrollbar.config(command=text.yview)
    
    def initHomeMenu(self): 
        # clear GUI
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate buttons
        openButton = tk.Button(root, text="Open Excel file", command=ep.openFile)
        openButton.pack()
        insertButton = tk.Button(root, text="Insert Data", command=ep.insertData)
        insertButton.pack()
        searchButton = tk.Button(root, text="Search Data", command=ep.search)
        searchButton.pack()

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
        print("Insert data")
        
        # clear layout
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