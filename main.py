import tkinter as tk
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

class ExcelProcessor: 
    def __init__(self):
        self.df = None
        self.tree = None

    def itemSelected(self, event): 
        selectedItem = self.tree.focus()
        values = self.tree.item(selectedItem, "values")
        print(values)

    def initTree(self):
        colNames = self.df.columns.tolist()
        cols = tuple(colNames)
        
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        self.tree.bind("<<TreeviewSelect>>", self.itemSelected)
        for col in cols: 
            self.tree.heading(col, text=col)        
        for row in self.df.itertuples(index=False): 
            self.tree.insert("", "end", values=row) 

        self.tree.place(x=10, y=50, height=500)

    def openFile(self): 
        filePath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

        if filePath: 
            try: 
                print("Openning file...")
                self.df = pd.read_excel(filePath)
                self.filePath = filePath
                self.initTree()
            except pd.errors.ParserError: 
                messagebox.showerror("Error", "Please open a valid Excel File")

    def initHomeMenu(self): 
        print("Main menu")
        # clear GUI
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate buttons
        openButton = tk.Button(root, text="Open Excel file", command=self.openFile)
        insertButton = tk.Button(root, text="Insert Data", command=self.insertData)
        searchButton = tk.Button(root, text="Search Data", command=self.initSearch)

        openButton.place(x=10, y=10) 
        insertButton.place(x=120, y=10) 
        searchButton.place(x=210, y=10) 

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
            print("Data inserted successfully!")
            print(self.df)

            # save updated DataFrame to Excel File
            self.df.to_excel(self.filePath, index=False)
            resLabel.configure(text="Excel file updated!")
        except: 
            resLabel.configure(text="Update failed")
        
        # clear form
        for entry in inputEntries.values():
            entry.delete(0, tk.END)


    def insertData(self): 
        print("Insert data")
        if self.df is None: 
            print("Please open an excel file")
            return 
        
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
            entry = tk.Entry(self.root)
            label.place(x=initX, y=initY + cnt * 30)
            entry.place(x=initX + 100, y=initY + cnt * 30)   
            cnt += 1 
            
            inputEntries[colName] = entry

        # generate other components
        resLablel = tk.Label(self.root, text="")        
        submitButton = tk.Button(self.root, text="Submit", command=lambda: self.handleInsertData(colNames, inputEntries, resLablel))
        cancelButton = tk.Button(self.root, text="Cancel", command=self.initHomeMenu)
        submitButton.place(x=initX + 100, y=initY + cnt * 30)
        cancelButton.place(x=initX + 160, y=initY + cnt * 30)
        resLablel.place(x=initX, y=initY + (cnt + 1) * 30)

    def handleAccurateSearch(self, colNames, inputEntries, initX, initY, cnt): 
        print("Searching...")
        # get user input 
        inputData = {}
        for colName in colNames:
            entryValue = inputEntries[colName].get()
            inputData[colName] = entryValue

        # accurate search 
        matchingRows = []
        for i, row in self.df.iterrows():
            match = True
            for colName, value in inputData.items(): 
                print("Searching: " + colName + ", " + value)
                print("Getting: " + colName + ", " + str(row[colName]))
                if str(inputData[colName]) == "":
                    print("The column " + colName + "is empty") 
                    continue
                if str(row[colName]) != value: 
                    match = False
                    break
            if match: 
                matchingRows.append(row)
        matchingDf = pd.DataFrame(matchingRows)

        cols = tuple(colNames)
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        self.tree.bind("<<TreeviewSelect>>", self.itemSelected)
        for col in cols: 
            self.tree.heading(col, text=col)        
        for row in matchingDf.itertuples(index=False): 
            self.tree.insert("", "end", values=row) 

        self.tree.place(x=initX, y=initY + (cnt + 1) * 30 + 10, height=500)

    def accurateSearch(self): 
        print("Switch to accurate search")
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
        searchButton = tk.Button(self.root, text="Search", command=lambda: self.handleAccurateSearch(colNames, inputEntries, initX, initY, cnt))
        cancelButton = tk.Button(self.root, text="Cancel", command=self.initSearch)
        searchButton.place(x=initX + 100, y=initY + cnt * 30)
        cancelButton.place(x=initX + 160, y=initY + cnt * 30)

    def handleFuzzySearch(self, entry):
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
                print(colName)
                print(matchingRows)
                res = pd.concat([res, matchingRows], ignore_index=True)
        
        cols = tuple(colNames)
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        self.tree.bind("<<TreeviewSelect>>", self.itemSelected)
        for col in cols: 
            self.tree.heading(col, text=col)        
        for row in res.itertuples(index=False): 
            self.tree.insert("", "end", values=row) 

        self.tree.place(x=10, y=50, height=500)

    def initSearch(self):
        print("Search data")
        if self.df is None: 
            print("Please open an excel file")
            return 
        
        # clear GUI
        for widget in self.root.winfo_children():
            widget.destroy()

        # generate layout
        entry = tk.Entry(root)
        label = tk.Label(root, text="Search:")
        fuzzySerachButton = tk.Button(root, text="Serach", command=lambda: self.handleFuzzySearch(entry))
        accurateSearchButton = tk.Button(root, text="Accuarte Serach", command=self.accurateSearch)
        cancelButton = tk.Button(self.root, text="Cancel", command=self.initHomeMenu)

        label.place(x=10, y=15)
        entry.place(x=80, y=15)
        fuzzySerachButton.place(x=240, y=10)
        accurateSearchButton.place(x=300, y=10)
        cancelButton.place(x=415, y=10)

root = tk.Tk()
root.title("Simple Excel Processor")
root.minsize(width=470, height=300)

ep = ExcelProcessor()
ep.root = root
ep.initHomeMenu()

root.mainloop()