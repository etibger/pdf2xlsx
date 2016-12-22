"""
Simple tkinter based gui around the pdf2xlsx.do_it function.
"""
from tkinter import *
from tkinter import ttk, filedialog, messagebox
# -*- coding: utf-8 -*-
from .pdf2xlsx import do_it
from .config import config

__version__ = '0.1.0'

class ConfigWindow():
    def __init__(self, master):
        self.master = master
        self.window = Toplevel(self.master)
        self.window.withdraw()
        self.window.title('Settings...')
        
        self.main_frame = ttk.Frame(self.window)
        self.main_frame.pack(padx = 5, pady = 5)

        ttk.Label(self.main_frame, text = 'Configuration:').grid(row = 0, column = 0, columnspan=2, sticky = 'w') 

        ttk.Label(self.main_frame, text = 'tmp dir:').grid(row = 1, column = 0, sticky = 'w')    
        self.tmp_dir_sv = StringVar()
        self.tmp_dir_sv.set(config['tmp_dir'])
        self.tmp_dir_entry = ttk.Entry(self.main_frame, textvariable=self.tmp_dir_sv)
        self.tmp_dir_entry.grid(row = 1, column = 1, sticky = 'e')

        ttk.Label(self.main_frame, text = 'file ext:').grid(row = 2, column = 0, sticky = 'w')    
        self.file_ext_sv = StringVar()
        self.file_ext_sv.set(config['file_extension'])
        self.file_ext_entry = ttk.Entry(self.main_frame, textvariable=self.file_ext_sv)
        self.file_ext_entry.grid(row = 2, column = 1, sticky = 'e')

        ttk.Label(self.main_frame, text = 'xlsx name:').grid(row = 3, column = 0, sticky = 'w')    
        self.xlsx_name_sv = StringVar()
        self.xlsx_name_sv.set(config['xlsx_name'])
        self.xlsx_name_entry = ttk.Entry(self.main_frame, textvariable=self.xlsx_name_sv)
        self.xlsx_name_entry.grid(row = 3, column = 1, sticky = 'e')

        ttk.Label(self.main_frame, text = 'invo header pos:').grid(row = 4, column = 0, sticky = 'w')    
        self.invo_head_pos_sv = StringVar()
        self.invo_head_pos_sv.set(", ".join(map(str,config['invo_header_ident'])))
        self.invo_head_pos_entry = ttk.Entry(self.main_frame, textvariable=self.invo_head_pos_sv)
        self.invo_head_pos_entry.grid(row = 4, column = 1, sticky = 'e')

        ttk.Label(self.main_frame, text = 'ME category:').grid(row = 5, column = 0, sticky = 'w')    
        self.me_cat_sv = StringVar()
        self.me_cat_sv.set(", ".join(map(str,config['ME'])))
        self.me_cat_entry = ttk.Entry(self.main_frame, textvariable=self.me_cat_sv)
        self.me_cat_entry.grid(row = 5, column = 1, sticky = 'e')

        #ttk.Style().configure("TButton", padding=6)
        
        self.bok = ttk.Button(self.main_frame, text = 'Save',
                   command = self.save_callback).grid(row = 6, column = 0, sticky = 'e')

        ttk.Button(self.main_frame, text = 'Accept',
                   command = self.accept_callback).grid(row = 6, column = 1, sticky = 'w')

    def save_callback(self):
        self.window.withdraw()
        self.accept_callback()

    def accept_callback(self):
        config['tmp_dir'] = self.tmp_dir_sv.get()
        config['file_extension'] = self.file_ext_sv.get()
        config['xlsx_name'] = self.xlsx_name_sv.get()
        config['invo_header_ident'] = list(map(int,self.invo_head_pos_sv.get().split(', ')))
        config['ME'] = self.me_cat_entry.get().split(', ')
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
        print("button pushed")
            
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
