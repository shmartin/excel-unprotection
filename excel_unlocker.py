# application version 0.0.1
# programmed by juan

import tkinter as tk
from tkinter import filedialog as fd
from openpyxl import load_workbook as lw

class UI_Main:

    # 

    def __init__(self):

        self.font_size = 20
        self.font_face = 'system-default'

        self.root_window = tk.Tk()
        self.root_window.geometry('800x200')
        self.root_window.title('Excel Unlocker v0.0.1')

        self.label_filedirectory = tk.Label(self.root_window, text = '', font = (self.font_face, self.font_size))
        self.label_filedirectory.pack(fill = 'both', expand = True)

        self.button_openfile = tk.Button(self.root_window, text = 'OPEN FILE', font = (self.font_face, self.font_size), command = self.file_dialog_open) 
        self.button_openfile.pack(fill = 'both', expand = True)

        self.button_unprotect = tk.Button(self.root_window, text = 'UNPROTECT FILE', font = (self.font_face, self.font_size), command = self.unlocker) 
        self.button_unprotect.pack(fill = 'both', expand = True)

        self.root_window.mainloop()

    def unlocker(self):
        try:
            self.wb = lw(fr'{self.file_dialog}')
            for s in self.wb.sheetnames:
                self.wb[s].protection.sheet = False
            self.wb.save(fr'{self.file_dialog}')
            self.label_filedirectory.config(text = fr'{self.file_dialog}', fg = 'green')
        except:
            self.label_filedirectory.config(text = 'Invalid File Type. Excel files only', fg = 'red')

    def file_dialog_open(self):
        self.file_dialog = fd.askopenfilename(title = 'juan', filetypes = (('All file types', '*.*'),('Excel Workbook', '*.xlsx'),('Excel Macro-Enabled Workbook (code)', '*.xlsm'),('Template','*.xltx'),('Template (code)','*.xltm')))
        self.label_filedirectory.config(text = fr'{self.file_dialog}', fg = 'black')
    
def main():
    UI_Main()

if __name__ == '__main__':
    main()
