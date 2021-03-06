#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import os
import time
import Tkinter as tk
from tkFileDialog import askopenfilename
import tkMessageBox
from ttk import Frame, Style, Combobox
from win32com.shell import shell, shellcon
import pythoncom

#custom imports
import csvwriter as cw
import formatxls as xls

#imports if I go back to python3
#import tkinter as tk
#from tkinter.filedialog import askopenfilename as askopenfilename
#from tkinter import messagebox as tkMessageBox
#from tkinter.ttk import Frame, Style, Combobox
# remember if baack to python3 change xrange to range 

class gui(Frame):
    def __init__(self,parent):
        Frame.__init__(self, parent)   
         
        self.parent         = parent
        self.filevar        = tk.StringVar()
        self.radiovar       = tk.StringVar()
        self.encodingvar    = tk.StringVar()
        self.entry          = tk.Entry(self, bd = 5, width=4)
        self.filename       = " "
        self.headercell     = "A1"
        self.abformat       = tk.BooleanVar()
        self.one2one        = tk.BooleanVar()
        self.oneheader      = tk.BooleanVar()
        self.filevar.set("C:")
        self.entry.insert(0, "A1")
        
        # define options for opening or saving a file
        desktop = shell.SHGetFolderPath (0, shellcon.CSIDL_DESKTOP, 0, 0) # only windows
        self.file_opt = options = {}
        options['defaultextension'] = '.txt'
        options['filetypes'] = [('Excel files', '*.xls;*.xlsx'),('all files', '*.*')]
        options['initialdir'] = desktop
        options['parent'] = self.parent
        options['title'] = 'Choose file'
            
        self._initUI()
        
        
    def _initUI(self):
        """
        Setup the gui
        """
        self.parent.title("xls2csv")
        self.pack(fill=tk.BOTH, expand=1)
        
        style = Style()
        style.configure("TFrame")#, background="#333")  
        
        
        ############################################################
        # definitions of elements                                  #
        ############################################################
        
        #labels
        L1 = tk.Label(self, text="Date format")
        L2 = tk.Label(self, text="Header cell")
        L3 = tk.Label(self, textvariable=self.filevar)
        L4 = tk.Label(self, text="Encoding")
        L5 = tk.Label(self, text="Other options")
        L6 = tk.Label(self, text="Delimeter")
        
        #entry
        #entrybutton = tk.Button(self, text="Get", command= self.on_button)
          
        # radio buttons
        R1 = tk.Radiobutton(self, text="31/12/2099", variable= self.radiovar, value="%d/%m/%Y")
        R2 = tk.Radiobutton(self, text="20991231", variable= self.radiovar, value="%Y%m%d")
        R3 = tk.Radiobutton(self, text="31-Dec-2099", variable= self.radiovar, value="%d%b%Y")    
        R4 = tk.Radiobutton(self, text="UTF-8 (Unicode)", variable= self.encodingvar, value="UTF-8")
        R5 = tk.Radiobutton(self, text="cp1252 (Windows)", variable= self.encodingvar, value="cp1252")
        #clear all radio buttons and select R1
        R1.deselect()
        R2.deselect()
        R3.deselect()
        R4.deselect()
        R5.deselect()
        
        R1.select()
        R5.select()
    
        #button
        runBTN = tk.Button(self, text ="Run", command =  self._xls2csv, bg = 'white' )
        chooseBTN = tk.Button(self, text ="Choose file", command =  self.open_file, bg = 'white' )
        closeButton = tk.Button(self, text =" Close ", command = lambda self=self: self.close_top() , bg = 'white' )
        
        
        #checkbox
        C1 = tk.Checkbutton( self, text="Only header in first sheet", variable=self.oneheader, onvalue=True, offvalue=False)
        C2 = tk.Checkbutton( self, text="One csv pr. one sheet", variable=self.one2one, onvalue=True, offvalue=False)
        C3 = tk.Checkbutton( self, text="Ascending numbers in Col A&B", variable=self.abformat, onvalue=True, offvalue=False)
        
        
        #combobox
        self.box_value = tk.StringVar()
        self.box = Combobox(self.parent, textvariable=self.box_value, state='readonly', width = 4)
        self.box['values'] = (';', ',', 'TAB')
        self.box.current(0)
        #value= box['values'][0]
        
        
        ############################################################
        # placement of elements                                  #
        ############################################################        
        
               
        #left side
        L1.place(x = 0 , y = 10 )
        R1.place(x = 0 , y = 40 )
        R2.place(x = 0 , y = 70 )
        R3.place(x = 0 , y = 100 )
        L2.place(x = 0 , y = 140 )
        self.entry.place(x = 100 , y = 140 )
        
        chooseBTN.place(x=5 , y = 185)
        L3.place(x = 100 , y = 190 )
        runBTN.place(x=5 , y = 230)
        
        
        
        # Center
        L4.place(x = 160 , y = 10 )
        R4.place(x = 160 , y = 40 )
        R5.place(x = 160 , y = 70 )
        L6.place(x = 160 , y = 140 )
        self.box.place(x = 240 , y = 140 )
        
        # Right side
        L5.place(x = 320 , y = 10 )
        C1.place(x = 320 , y = 40 ) 
        C2.place(x = 320 , y = 70 )
        C3.place(x = 320 , y = 100 )
    
        closeButton.place(x=500, y=230)
        

    def open_file(self):
        ret = 1
        self.filename = askopenfilename(**self.file_opt)     
        self.filevar.set(self.filename)


    def close_top(self):
        self.parent.destroy()    
        
    def _xls2csv(self):
        # get delimiter
        if self.box_value.get() == 'TAB':
            delimiter = '\t'
        else:
            delimiter = self.box_value.get()
        
        #manipulate xls file
        self.headercell = self.entry.get()
        try:
            excel = xls.formatxls(self.filename, self.headercell, self.one2one.get(),self.oneheader.get())
            excel.process_workbook()
        except pythoncom.com_error as excelerror: #, (hr, msg, exc, arg):
            hr,msg,exc,arg = excelerror.args
            err_msg =  "The Excel call failed with code %d: %s" % (hr, msg)
            err_msg += '\n'
            if exc is None:
                err_msg += "There is no extended error information"
            else:
                    wcode, source, text, helpFile, helpId, scode = exc
                    err_msg += "The source of the error is", source
                    err_msg += '\n'
                    err_msg += "The error message is", text
                    err_msg += '\n'
                    err_msg += "More info can be found in %s (id=%d)" % (helpFile, helpId)
            tkMessageBox.showinfo( "",err_msg)
      
      #write to csv
        csv_wr = cw.csvwriter(self.filename, str(self.radiovar.get()),
                              str(self.encodingvar.get()), self.abformat.get(), 
                              self.one2one.get(), delimiter )
        ret = csv_wr.xlsallsheet2onecsv()
        if ret == 0 :
            tkMessageBox.showinfo( "","CSV file Written")
        else :
            tkMessageBox.showinfo( "","Error writing csv file")
        
def main():
    root = tk.Tk()
    root.geometry("560x280+300+300")
    app = gui(root)
    root.mainloop()  
   

if __name__ == "__main__":

    main()  