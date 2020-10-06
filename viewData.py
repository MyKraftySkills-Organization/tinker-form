from tkinter import *
from pandastable import Table, TableModel
from pandas import read_excel

class TestApp(Frame):
    """Basic test frame for the table"""
    def __init__(self, parent=None):
        file = read_excel('excel.xlsx')
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('600x400+200+100')
        self.main.title('Table app')
        f = Frame(self.main)
        f.pack(fill=BOTH,expand=1)
        # df = TableModel.getSampleData()
        df = file
        self.table = pt = Table(f, dataframe=df,
                                showtoolbar=True, showstatusbar=True)
        pt.show()
        return


app = TestApp()
#launch the app
app.mainloop()