from tkinter import *
from tkinter import filedialog
import openpyxl
import pandas as pd


def new_file():
    print("New File!")


def open_file():
    name = filedialog.askopenfilename()
    print(name)


def about():
    print("This is a simple example of a menu")


# creates excel file with a given file name
def create_excel(filename):
    wb = openpyxl.Workbook()
    wb.save(filename)


# creates a menu and assign them action functions
def create_menu(menubar):
    file_menu = Menu(menubar)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="New", command=new_file)
    file_menu.add_command(label="Open...", command=open_file)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)
    help_menu = Menu(menubar)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="About...", command=about)


if __name__ == "__main__":
    # create a GUI window
    root = Tk()

    # creating menubar
    menubar = Menu(root)
    # setting menubar to window
    root.config(menu=menubar)
    # adding controls to menubar
    create_menu(menubar)

    # set the background colour of GUI window
    root.configure(background='light blue')

    # set the title of GUI window
    root.title("Window")

    # set the configuration of GUI window
    root.geometry("720x400")

    # start the GUI
    root.mainloop()