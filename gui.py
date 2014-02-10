#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import os
import Tkinter as tk
from tkFileDialog import askopenfilename
import tkMessageBox
from ttk import Frame, Style

#custom imports
#import csvwriter as cw
#import formatxls as xls



class gui(Frame):
    def __init__(self,parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.filevar = tk.StringVar()
        self.filevar.set("C:")
        self.radiovar = tk.StringVar()
        self.radiovar = None
        self.entry = tk.Entry(self, bd = 5)
        self.entry.insert(0, "A1")
        
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
        R1 = tk.Radiobutton(self, text="dd/mm/yyyy", variable= self.radiovar, value="%d/%m/%Y")
        R2 = tk.Radiobutton(self, text="yyyymmdd", variable= self.radiovar, value="%Y%m%d")
    
        #button
        chooseBTN = tk.Button(self, text ="Choose file", command =  self.open_file, bg = 'white' )
        runBTN = tk.Button(self, text ="Run", command =  self._xls2csv, bg = 'white' )
        closeButton = tk.Button(self, text =" Close ", command = lambda self=self: self.close_top() , bg = 'white' )
        
        
        L3.place(x = 100 , y = 175 )
        L2.place(x = 0 , y = 10 )
        R1.place(x = 0 , y = 40 )
        R2.place(x = 0 , y = 70 )
        L1.place(x = 0 , y = 110 )
        self.entry.place(x = 100 , y = 110 )
       
        chooseBTN.place(x=5 , y = 170)
        runBTN.place(x=5 , y = 220)
        closeButton.place(x=440, y=240)


    def sel(self):
        selection = "You selected the option " + str( self.radiovar.get())



    def open_file(self):
        ret = 1
        filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
        #ret = csv_from_excel(filename)
        #if ret == 0 :
        #    tkMessageBox.showinfo( "","Fil processeret")
        #else :
        #    tkMessageBox.showinfo( "","Fil fejlet")
        self.filevar.set(filename)


    def close_top(self):
        self.parent.destroy()
    
    def on_button(self):
        print self.entry.get()
        
        
    def _xls2csv(self):
        print 'hej'
        
def main():
    root = tk.Tk()
    root.geometry("500x280+300+300")
    app = gui(root)
    root.mainloop()  
   

if __name__ == "__main__":

    main()  