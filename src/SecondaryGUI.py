import threading
import tkinter as tk
from tkinter import ttk

import BankProcesses.THRDProcesses.GmailHandler as gmail

class BaseWindow(tk.Toplevel):
    def __init__(
        self, bank, bank_module, 
        info="Info", 
        error="Error",
        specific_error: bool = True, 
    ):
        super().__init__()
        self.bank = bank
        self.bank_module = bank_module
        self.info = info
        self.error = error
        self.specific_error = specific_error

    def create_window(self):
        self.title("History of Expenses GUI")
        self.geometry('600x250')
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
            text=f"You chose {self.bank}.\nClick the button below to process your bank", 
            justify=tk.CENTER,
            wraplength=250,
        )   
        self.process_button = ttk.Button(
            self.mainframe, 
            text=f"Process {self.bank}",
            command=self.process_file,
        )
        self.button_close = ttk.Button(
            self.mainframe, 
            text="Choose another bank", 
            command=self.destroy,
        )
        self.button_abort = ttk.Button(
            self.mainframe,
            text="Abort",
            command=self.quit,
            width=21,
        )
        
        self.empty_label.grid(column=0, row=0, sticky=tk.N)
        self.info_label.grid(column=1, row=0, sticky=tk.N)
        self.process_button.grid(column=1, row=2, sticky=tk.N)
        self.button_close.grid(column=0, row=3, sticky=tk.N)
        self.button_abort.grid(column=2, row=3, sticky=tk.N)
        
        # Sets the focus and prevents the user from using the main window
        self.focus_set()
        self.grab_set()
        
        for child in self.mainframe.winfo_children(): 
            child.grid_configure(padx=1, pady=5)
            
            
    # The following functions work together to handle the "Process bank" button functionality
    # This section utilizes threading to manage button states and display progress messages    
    
    # Executes the main function in the bank module
    def run_bank(self):
        self.bank_module.main()
        
    # Schedules a periodic check to monitor task completion
    def schedule_check(self, t):
        self.after(100, self.check_if_done, t)
        
    # Checks if the thread has finished; re-enables the button and displays a message upon completion
    # If not finished, continues to check after one second
    def check_if_done(self, t):
        if not t.is_alive():
            self.handle_specific_errors()
            
            if not self.specific_error:
                self.handle_errors()
                self.update_button_states('normal')
            
            self.update_button_states('normal')
            
        else:
            self.schedule_check(t)
    
    # Initiates the file processing procedure
    def process_file(self):
        self.update_info_label("Processing file(s)...")
        
        # Disables the buttons during file processing
        self.update_button_states('disabled')
        
        # Initiates file processing in a separate thread
        t = threading.Thread(target=self.run_bank)
        t.start()
        
        # Periodically checks the status of the processing thread
        self.schedule_check(t)        

    def handle_specific_errors(self):  # will be inherited in the subclass
        self.specific_error = False
        return self.specific_error

    def handle_errors(self):
        if self.bank_module.file_error:
            tk.messagebox.showinfo(parent=self, title=self.info, message=self.bank_module.file_error)
            self.update_info_label(self.bank_module.file_error)
            
        elif not self.bank_module.procedures_ended:
            msg = "Something went wrong!\nCheck the statement and try again."
            tk.messagebox.showerror(parent=self, title=self.error, message=msg)
            self.update_info_label(msg)
            
        else:
            self.update_info_label("File(s) successfully processed!")


    def update_info_label(self, message, wraplength=250):
        self.info_label['text'] = message
        self.info_label['wraplength'] = wraplength
        
    def update_button_states(self, state):
        self.process_button['state'] = state
        self.button_close['state'] = state
        self.button_abort['state'] = state
    
    
class FRSTWindow(BaseWindow):
    def __init__(self, bank, bank_module):
        super().__init__(bank, bank_module)
        self.create_window()
        self.create_buttons()

class SCNDWindow(BaseWindow):
    def __init__(self, bank, bank_module):
        super().__init__(bank, bank_module)
        self.create_window()
        self.create_buttons()
        
class THRDWindow(BaseWindow):
    def __init__(self, bank, bank_module):
        super().__init__(bank, bank_module)
        self.create_window()
        self.create_buttons()
          
    def handle_specific_errors(self):
        if gmail.gmail_error:
            tk.messagebox.showerror(self.error, gmail.gmail_error, parent=self)
            self.geometry('900x250')
            self.update_info_label(gmail.gmail_error, wraplength=550)
            
        elif gmail.no_email:
            tk.messagebox.showinfo(self.info, gmail.no_email, parent=self)
            self.update_info_label(gmail.no_email) 
            
        elif gmail.duplicate_statement:
            tk.messagebox.showerror(self.error, gmail.duplicate_statement, parent=self)
            self.geometry('700x250')
            self.update_info_label(gmail.duplicate_statement, wraplength=350)
                
        else:
            self.specific_error = False
            return self.specific_error
        
