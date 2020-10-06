from tkinter import *
from tkinter import filedialog
from create_menu_function import *
from pandastable import Table
import openpyxl
import pandas as pd


# creates excel file with a given file name
def create_excel(filename):
    wb = openpyxl.Workbook()
    wb.save(filename)


if __name__ == "__main__":
    # create a GUI window
    root = Tk()

    # creating menubar
    menubar = Menu(root)
    # setting menubar to window
    root.config(menu=menubar)
    # adding controls to menubar
    create_menu(menubar)

    # insert table
    table = Table(root, dataframe=df, showtoolbar=False, showstatusbar=False)

    # set the background colour of GUI window
    root.configure(background='light blue')

    # set the title of GUI window
    root.title("Window")

    # set the configuration of GUI window
    root.geometry("720x400")

    # start the GUI
    root.mainloop()