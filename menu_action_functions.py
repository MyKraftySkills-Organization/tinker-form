import pandas as pd
from tkinter.filedialog import askopenfilename


# Menu Actions
def new_file():
    print("New File!")


def open_file():
    name = askopenfilename(filetypes=[('CSV', '*.csv',), ('Excel', ('*.xls', '*.xlsx'))])
    load_data(name)


def about():
    print("This is a simple example of a menu")


def load_data(name):
    if name:
        if name.endswith('.csv'):
            df = pd.read_csv(name)
        else:
            df = pd.read_excel(name)