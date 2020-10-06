from menu_action_functions import *
from tkinter import *

# creates a menu and assign them action functions
def create_menu(menubar):
    file_menu = Menu(menubar)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="New", command=new_file)
    file_menu.add_command(label="Open...", command=open_file)
    file_menu.add_separator()
    file_menu.add_command(label="Exit")
    help_menu = Menu(menubar)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="About...", command=about)
