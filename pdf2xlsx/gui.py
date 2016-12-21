"""
Simple tkinter based gui around the pdf2xlsx.do_it function.
"""
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from .pdf2xlsx import do_it

__version__ = '0.1.0'

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
            logger = do_it(self.src_entry.get(),self.dest_entry.get())
            messagebox.showinfo(title = 'Conversion Completed',
                            message = 'The following Invoices/Entries were found:\n{0!s}'.format(logger))
        except PermissionError as e:
            messagebox.showerror('Exception', e)
        
        
        
def main():    
    root = Tk()
    gui = PdfXlsxGui(root)
    root.mainloop()

if __name__ == '__main__' : main()
