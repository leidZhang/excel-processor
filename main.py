import tkinter as tk
import pandas as pd
import numpy as np
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk


class ExcelProcessor: 
    def __init__(self):
        self.df = None
        self.tree = None
        self.get = None
        self.selected = None
        self.selectedID = None
        self.selectedItem = None

    def itemSelected(self, event): 
        self.selectedItem = int(self.tree.focus())
        self.selected = self.get.iloc[self.selectedItem-1]
        self.selectedID = self.selected.name
        print(self.selected)
        print(self.selected.name)

    def genForm(self, mode): 
        if self.df is None: 
            print("Please open an excel file")
            return 
        print("Generating form...")

        # clear layout
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate form
        initX = 10
        initY = 10
        cnt = 0
        colNames = self.df.columns.tolist()
        inputEntries = {}
        for colName in colNames:
            label = tk.Label(self.root, text=colName)
            entry = tk.Entry(self.root, name=str(cnt))
            label.place(x=initX, y=initY + cnt * 30)
            entry.place(x=initX + 100, y=initY + cnt * 30)   
            cnt += 1 
            
            inputEntries[colName] = entry

        # generate other components
        resLablel = tk.Label(self.root, text="")        
        cancelButton = tk.Button(self.root, text="Cancel", command=self.initHomeMenu)
        cancelButton.place(x=initX + 160, y=initY + cnt * 30)
        resLablel.place(x=initX, y=initY + (cnt + 1) * 30)

        if mode == "Insert": 
            submitButton = tk.Button(self.root, text=mode, command=lambda: self.handleInsertData(colNames, inputEntries, resLablel))
            submitButton.place(x=initX + 100, y=initY + cnt * 30)
        else: 
            updateButton = tk.Button(self.root, text=mode, command=lambda: self.handleUpdateData(inputEntries))
            updateButton.place(x=initX + 100, y=initY + cnt * 30)
        
        print("Generationg complete")
    
    def updateData(self): 
        # generate form
        self.genForm("Update")

        # fill the form
        cnt = 0        
        colNames = self.df.columns.to_list()
        for colName in colNames: 
            entry = root.nametowidget(str(cnt))
            value = self.selected[colName]
            entry.insert(0, value)
            cnt += 1
    
    def handleUpdateData(self, inputEntries):
        print("Handling update...")

        data = {}
        for key, value in inputEntries.items(): 
            data[key] = value.get()        
        newDf = pd.DataFrame(data, index=[0]) 
        self.df.loc[self.selectedID] = newDf.loc[0]
        
        print(newDf)
        self.selected = None
        self.selectedID = None

        # save updated DataFrame to Excel File
        self.df.to_excel(self.filePath, index=False)
        print("Update successful")
    
    def handleDeleteData(self):
        print("Executing deletion...")
        
        # execute deletion
        try: 
            self.df.drop(index=self.selectedID, inplace=True)
            self.df.to_excel(self.filePath, index=False)
            self.tree.delete(self.tree.get_children()[self.selectedItem-1])
            
            self.selected = None
            self.selectedID = None
            self.selectedItem = None
            print("Deletion successful")
        except: 
            print("Deletion failed")


    def initTree(self):
        colNames = self.df.columns.tolist()
        cols = tuple(colNames)
        
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        self.tree.bind("<<TreeviewSelect>>", self.itemSelected)
        for col in cols: 
            self.tree.heading(col, text=col)    
        cnt = 1    
        for row in self.df.itertuples(index=False): 
            self.tree.insert("", "end", iid=cnt, values=row) 
            cnt += 1

        self.tree.place(x=10, y=50, height=500)

        updateButton = tk.Button(self.root, text="Update", command=self.updateData)
        deleteButton = tk.Button(self.root, text="Delete", command=self.handleDeleteData)
        updateButton.place(x=10, y=560)
        deleteButton.place(x=70, y=560)

    def openFile(self): 
        filePath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

        if filePath: 
            try: 
                print("Openning file...")
                self.df = pd.read_excel(filePath)
                self.get = self.df
                self.filePath = filePath
                self.initTree()
                print("File Opened")
            except pd.errors.ParserError: 
                messagebox.showerror("Error", "Please open a valid Excel File")

    def handleFuzzySearch(self, entry):
        self.tree.delete(*self.tree.get_children())

        res = pd.DataFrame([])
        inputData = entry.get()
        print("Searching " + inputData + "...")
        
        colNames = self.df.columns.tolist()
        if (inputData == ""): 
            print("No specific condition, printing all data...")
            inputData = "all rows"
            res = self.df
        else: 
            print("Searching " + inputData)
            # search each columns
            for colName in colNames: 
                matchingRows = self.df[self.df[colName].astype(str).str.contains(inputData, case=False)]
                res = pd.concat([res, matchingRows], ignore_index=False)

        self.get = res 
        # self.initTree()
        cnt = 1
        print("____________")
        for row in res.itertuples(index=False): 
            self.tree.insert("", "end", iid=cnt, values=row) 
            cnt += 1

    def getMatchingRows(self, inputData, mode):
        # execute search
        matchingRows = []
        for i, row in self.df.iterrows():
            match = True
            for colName, value in inputData.items(): 
                print("Searching: " + colName + ", " + value)
                print("Getting: " + colName + ", " + str(row[colName]))
                if str(inputData[colName]) == "":
                    print("The column " + colName + "is empty") 
                    continue
                if mode == "adv":
                    if value not in str(row[colName]): 
                        match = False
                        break
                else: 
                    if str(row[colName]) != value:
                        match = False
                        break
            if match: 
                matchingRows.append(row)

        return pd.DataFrame(matchingRows)

    def handleAdvSearch(self, colNames, inputEntries, initX, initY): 
        print("Searching...")
        # get user input 
        inputData = {}
        for colName in colNames:
            entryValue = inputEntries[colName].get()
            inputData[colName] = entryValue
        # get matching rows
        matchingDf = self.getMatchingRows(inputData, "adv")
        self.get = matchingDf

        cols = tuple(colNames)
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        self.tree.bind("<<TreeviewSelect>>", self.itemSelected)
        for col in cols: 
            self.tree.heading(col, text=col) 
        cnt = 1       
        for row in matchingDf.itertuples(index=False): 
            self.tree.insert("", "end", iid=cnt, values=row) 
            cnt += 1

        self.tree.place(x=initX + 250, y=initY, height=500)
        deleteButton = tk.Button(self.root, text="Delete", command=self.handleDeleteData)
        updateButton = tk.Button(self.root, text="Update", command=self.updateData)
        updateButton.place(x=initX + 250, y=515)
        deleteButton.place(x=initX + 310, y=515)

    def advSearch(self): 
        print("Switch to advanced search")
        if (self.df is None): 
            print("Please open a valid Excel File")
            return 
        
        # clear GUI
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate form
        initX = 10
        initY = 10
        cnt = 0
        colNames = self.df.columns.tolist()
        inputEntries = {}
        for colName in colNames:
            label = tk.Label(self.root, text=colName)
            entry = tk.Entry(self.root)
            label.place(x=initX, y=initY + cnt * 30)
            entry.place(x=initX + 100, y=initY + cnt * 30)     
            cnt += 1
            
            inputEntries[colName] = entry

        # generate other components
        searchButton = tk.Button(self.root, text="Search", command=lambda: self.handleAdvSearch(colNames, inputEntries, initX, initY))
        cancelButton = tk.Button(self.root, text="Cancel", command=self.initHomeMenu)
        searchButton.place(x=initX + 100, y=initY + cnt * 30)
        cancelButton.place(x=initX + 160, y=initY + cnt * 30)

    def initHomeMenu(self): 
        print("Main menu")
        # clear GUI
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate buttons
        entry = tk.Entry(root, text="")
        openButton = tk.Button(root, text="Open Excel file", command=self.openFile)
        searchButton = tk.Button(root, text="Search", command=lambda: self.handleFuzzySearch(entry))
        insertButton = tk.Button(root, text="Insert Data", command=lambda: self.genForm("Insert"))
        AdvSearchButton = tk.Button(root, text="Advanced Search", command=self.advSearch)

        openButton.place(x=10, y=10) 
        entry.place(x=115, y=15)
        searchButton.place(x=265, y=10)
        insertButton.place(x=320, y=10) 
        AdvSearchButton.place(x=400, y=10) 

        if self.df is not None:
            self.initTree()

    def handleInsertData(self, colNames, inputEntries, resLabel): 
        data = {}
        for colName in colNames:
            entryValue = inputEntries[colName].get()
            data[colName] = entryValue

        try: 
            # insert data to the DataFrame
            newDf = pd.DataFrame(data, index=[0]) 
            self.df = pd.concat([self.df, newDf], ignore_index=True) 
            self.get = self.df
            print("Data inserted successfully!")

            # save updated DataFrame to Excel File
            self.df.to_excel(self.filePath, index=False)
            resLabel.configure(text="Excel file updated!")
        except: 
            resLabel.configure(text="Update failed")
        
        # clear form
        for entry in inputEntries.values():
            entry.delete(0, tk.END)

root = tk.Tk()
root.title("Simple Excel Processor")
root.minsize(width=1500, height=650)

ep = ExcelProcessor()
ep.root = root
ep.initHomeMenu()

root.mainloop()