#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import os
import time
import Tkinter as tk
from tkFileDialog import askopenfilename
import tkMessageBox
from ttk import Frame, Style

#custom imports
import csvwriter as cw
import formatxls as xls



class gui(Frame):
    def __init__(self,parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.filevar = tk.StringVar()
        self.filevar.set("C:")
        self.radiovar = tk.StringVar()
        self.encodingvar = tk.StringVar()
        self.entry = tk.Entry(self, bd = 5)
        self.entry.insert(0, "A1")
        self.filename = "tmp"
        self.headercell = "tmp"
        self.cvar = tk.BooleanVar()
        
            # define options for opening or saving a file
        self.file_opt = options = {}
        options['defaultextension'] = '.txt'
        options['filetypes'] = [('Excel files', '.xls;.xlsx'),('all files', '.*')]
        options['initialdir'] = 'C:\\Desktop'
        options['parent'] = self.parent
        options['title'] = 'Choose file'
        
        
        
        self._initUI()
        
        
    def _initUI(self):
        
        self.parent.title("xls2csv")
        self.pack(fill=tk.BOTH, expand=1)
        
        style = Style()
        style.configure("TFrame")#, background="#333")  
        
        
        L3 = tk.Label(self, textvariable=self.filevar)
        
        #entry
        L1 = tk.Label(self, text="Header cell")
        entrybutton = tk.Button(self, text="Get", command= self.on_button)
          
        # radio buttons
        L2 = tk.Label(self, text="Date format")
        R1 = tk.Radiobutton(self, text="31/12/2099", variable= self.radiovar, value="%d/%m/%Y")
        R2 = tk.Radiobutton(self, text="20991231", variable= self.radiovar, value="%Y%m%d")
        R3 = tk.Radiobutton(self, text="31-Dec-2099", variable= self.radiovar, value="%d%b%Y")
        
        L4 = tk.Label(self, text="Encoding")
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
        chooseBTN = tk.Button(self, text ="Choose file", command =  self.open_file, bg = 'white' )
        runBTN = tk.Button(self, text ="Run", command =  self._xls2csv, bg = 'white' )
        closeButton = tk.Button(self, text =" Close ", command = lambda self=self: self.close_top() , bg = 'white' )
        
        #checkbox
        C1 = tk.Checkbutton( self, text="Ascending numbers in Col A&B", variable=self.cvar,onvalue=True, offvalue=False)
        
        
        L3.place(x = 100 , y = 190 )
        L2.place(x = 0 , y = 10 )
        R1.place(x = 0 , y = 40 )
        R2.place(x = 0 , y = 70 )
        R3.place(x = 0 , y = 100 )
        L1.place(x = 0 , y = 140 )
        self.entry.place(x = 100 , y = 140 )
        
        L4.place(x = 200 , y = 10 )
        R4.place(x = 200 , y = 40 )
        R5.place(x = 200 , y = 70 )
        C1.place(x = 200 , y = 100 ) 
       
        chooseBTN.place(x=5 , y = 185)
        runBTN.place(x=5 , y = 230)
        closeButton.place(x=440, y=240)
        

    def sel(self):
        selection = "You selected the option " + str( self.radiovar.get())



    def open_file(self):
        ret = 1
        self.filename = askopenfilename(**self.file_opt)     
        self.filevar.set(self.filename)


    def close_top(self):
        self.parent.destroy()
    
    def on_button(self):
        print self.entry.get()
        
        
    def _xls2csv(self):
        #manipulate xls file
        self.headercell = self.entry.get()
        excel = xls.formatxls(self.filename, self.headercell)
        excel.process_workbook()
        time.sleep(0.5)
        #write to csv
        csv_wr = cw.csvwriter(self.filename,str(self.radiovar.get()),str(self.encodingvar.get()),self.cvar.get() )
        csv_wr.xlsallsheet2onecsv()
        
def main():
    root = tk.Tk()
    root.geometry("500x280+300+300")
    app = gui(root)
    root.mainloop()  
   

if __name__ == "__main__":

    main()  