import tkinter as tk
from tkinter import ttk
import SecondaryGUI as gui

from BankProcesses import FirstBank as frst
from BankProcesses import SecondBank as scnd
from BankProcesses import ThirdBank as thrd

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.create_window()
        self.create_buttons()
    
    def create_window(self):
        self.title("History of Expenses (Main)")
        self.geometry('500x250')
        self.mainframe = ttk.Frame(self, borderwidth=2, relief='solid', padding='3 3 12 12')
        self.mainframe.grid(column=0, row=0, sticky=tk.N)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
    
    def create_buttons(self):
        self.empty_label = ttk.Label(
            self.mainframe, 
            width=10,
        )
        self.info_label = ttk.Label(
            self.mainframe, 
            text="Welcome to your History of Expenses program!\nWhich file do you want to process?", 
            justify=tk.CENTER,
        )
        self.process_frst = ttk.Button(
            self.mainframe, text="First Bank", 
            command=self.open_frst_window,
        )
        self.process_scnd = ttk.Button(
            self.mainframe, 
            text="Second Bank", 
            command=self.open_scnd_window,
        )
        self.process_thrd = ttk.Button(
            self.mainframe, 
            text="Third Bank", 
            command=self.open_thrd_window,
        )
        self.button_abort = ttk.Button(
            self.mainframe,
            text="Abort",
            command=self.quit,
        )
        
        self.empty_label.grid(column=0, row=0, sticky=tk.N)
        self.info_label.grid(column=2, row=0, sticky=tk.N)
        self.process_frst.grid(column=2, row=1, sticky=tk.N)
        self.process_scnd.grid(column=2, row=2, sticky=tk.N)
        self.process_thrd.grid(column=2, row=3, sticky=tk.N)
        self.button_abort.grid(column=3, row=3, sticky=tk.N)
        
        for child in self.mainframe.winfo_children(): 
            child.grid_configure(padx=1, pady=5)
    
    
    def open_frst_window(self):
        self.frst_window = gui.FRSTWindow("First Bank", frst)

    def open_scnd_window(self):
        self.scnd_window = gui.SCNDWindow("Second Bank", scnd)
        
    def open_thrd_window(self):
        self.thrd_window = gui.THRDWindow("Third Bank", thrd)


if __name__ == "__main__":
    new_window = MainWindow()
    new_window.mainloop()