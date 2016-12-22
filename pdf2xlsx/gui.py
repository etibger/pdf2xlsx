"""
Simple tkinter based gui around the pdf2xlsx.do_it function.
"""
from tkinter import *
from tkinter import ttk, filedialog, messagebox
# -*- coding: utf-8 -*-
from .pdf2xlsx import do_it
from .config import config

__version__ = '0.1.0'

class ConfOption():
    def __init__(self, root, name, text, row): 
        self.name = name
        ttk.Label(root, text=text).grid(row=row, column=0, sticky = 'w')    
        self.sv = StringVar()
        if isinstance(config[self.name],list):
            self.sv.set(", ".join(map(str,config[self.name])))
        else:        
            self.sv.set(config[self.name])
        self.entry = ttk.Entry(root, textvariable=self.sv)
        self.entry.grid(row = row, column = 1, sticky = 'e')

    def update_config(self):
        if isinstance(config[self.name],list):
            if isinstance(config[self.name][0],int):
                config[self.name] = list(map(int,self.sv.get().split(', ')))
            else:
                config[self.name] = self.sv.get().split(', ')
        else:
            config[self.name] = self.sv.get()

class ConfigWindow():
    def __init__(self, master):
        self.master = master
        self.window = Toplevel(self.master)
        self.window.withdraw()
        self.window.title('Settings...')
        self.conf_list = []
        
        self.main_frame = ttk.Frame(self.window)
        self.main_frame.pack(padx = 5, pady = 5)

        ttk.Label(self.main_frame, text = 'Configuration:').grid(row = 0, column = 0, columnspan=2, sticky = 'w') 

        self.conf_list.append(
            ConfOption(root=self.main_frame, name='tmp_dir', text='tmp dir:', row=1))
        
        self.conf_list.append(
            ConfOption(root=self.main_frame, name='file_extension', text='file ext:', row=2))
        
        self.conf_list.append(
            ConfOption(root=self.main_frame, name='xlsx_name', text='xlsx name:', row=3))
        
        self.conf_list.append(
            ConfOption(root=self.main_frame, name='invo_header_ident', text='invo header pos:', row=4))
        
        self.conf_list.append(
            ConfOption(root=self.main_frame, name='ME', text='ME category:', row=5))

        
        ttk.Button(self.main_frame, text = 'Save',
                   command = self.save_callback).grid(row = 6, column = 0, sticky = 'e')

        ttk.Button(self.main_frame, text = 'Accept',
                   command = self.accept_callback).grid(row = 6, column = 1, sticky = 'w')

    def save_callback(self):
        self.window.withdraw()
        self.accept_callback()

    def accept_callback(self):
        for conf in self.conf_list:
            conf.update_config()
        config.store()
        

class PdfXlsxGui():
    """
    Simple GUI which lets the user select the source file zip and the desitination directory
    for the xlsx file.

    :param master: Tk parent class
    """
    
    def __init__(self, master):        
        self.master = master      
        self.master.title('Convert Zip -> Pdf -> Xlsx')
        self.master.resizable(False, False)
        
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(padx = 5, pady = 5)
        
        ttk.Label(self.main_frame, text = 'Source File:').grid(row = 0, column = 0, sticky = 'w')        
        self.src_entry = ttk.Entry(self.main_frame, width = 54)
        self.src_entry.grid(row = 1, column = 0, sticky = 'e')
        self.src_entry.insert(0, '.\\src.zip')            
        ttk.Button(self.main_frame, text = 'Browse...',
                   command = self.browse_src_callback).grid(row = 1, column = 1, sticky = 'w')
        
        ttk.Label(self.main_frame, text = 'Destination Directory:').grid(row = 2, column = 0, sticky = 'w')   
        self.dest_entry = ttk.Entry(self.main_frame, width = 54)
        self.dest_entry.grid(row = 3, column = 0, sticky = 'e')
        self.dest_entry.insert(0, '.\\')        
        ttk.Button(self.main_frame, text = 'Browse...',
                   command = self.browse_dest_callback).grid(row = 3, column = 1, sticky = 'w')
              
        ttk.Button(self.main_frame, text = 'Convert to Xlsx',
                   command = self.process_pdf).grid(row = 5, column = 0, columnspan = 2)


        ttk.Button(self.main_frame, text = 'Settings',
                   command = self.config_callback).grid(row = 6, column = 1, columnspan = 1, sticky='e')

        self.config_window = ConfigWindow(self.master)

    def config_callback(self):
        self.config_window.window.state('normal')
            
    def browse_src_callback(self):
        """
        Asks for the source zip file, the opened dialog filters for zip files by default
        The src_entry attribute is updated based on selection
        """
        path = filedialog.askopenfilename(initialdir='.\\',
                                          title="Choose the Zip file...",
                                          filetypes=(("zip files","*.zip"),("all files","*.*")))
        self.src_entry.delete(0, END)
        self.src_entry.insert(0, path)
        
    def browse_dest_callback(self):
        """
        Asks for the destination directory to generate the xlsx file.
        the dest_entry attribute is updeted.
        """
        path = filedialog.askdirectory(initialdir = self.dest_entry.get())
        self.dest_entry.delete(0, END)
        self.dest_entry.insert(0, path)

    def process_pdf(self):
        """
        Faxade for the do_it function. Only the src file and destination dir is updated
        the other parameters are left for defaults.
        """
        try:
            logger = do_it(self.src_entry.get(),self.dest_entry.get(),
                           xlsx_name=config['xlsx_name'], tmp_dir=config['tmp_dir'],
                           file_extension=config['file_extension'])
            
            messagebox.showinfo(title = 'Conversion Completed',
                            message = 'The following Invoices/Entries were found:\n{0!s}'.format(logger))
        except PermissionError as e:
            messagebox.showerror('Exception', e)
        
        
        
def main():    
    root = Tk()
    gui = PdfXlsxGui(root)
    root.mainloop()

if __name__ == '__main__' : main()
