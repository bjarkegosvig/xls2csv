#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import os
import Tkinter as tk
from tkFileDialog import askopenfilename
import tkMessageBox

#custom imports
import csvwriter as cw
import formatxls as xls



class gui(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        
        
        self.var = tk.StringVar()
        self.label = tk.Label( self, textvariable=self.var, relief=tk.RAISED )
        self.var.set("Forsvaret Vest Excel2csv")

        #entry
        self.entry = tk.Entry(self)
        self.entrybutton = tk.Button(self, text="Get", command=self.on_button)
    
    
        # radio buttons
        self.radiovar = tk.IntVar()
        self.R1 = tk.Radiobutton(self, text="dd/mm/yyyy", variable=self.radiovar, value=1,command=self.sel)
        self.R2 = tk.Radiobutton(self, text="yyyymmdd", variable=self.radiovar, value=2,command=self.sel)
        self.radiolabel = tk.Label(self)
    
        #button
        self.B = tk.Button(self, text ="Vælg excel fil", command = self.open_file, bg = 'white' )
        self.Btn = tk.Button(self, text ="Luk", command = lambda self=self:self.close_top() , bg = 'white' )
        
        
        
        
        self.label.pack()
        self.R1.pack( anchor = tk.W )
        self.R2.pack( anchor = tk.W )
        self.radiolabel.pack(anchor = tk.W)
        self.B.pack()
        self.Btn.pack()
        self.entrybutton.pack()
        self.entry.pack()


    def sel(self):
        selection = "You selected the option " + str(self.radiovar.get())
        self.radiolabel.config(text = selection)


    def open_file(self):
        ret = 1
        filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
        #ret = csv_from_excel(filename)
        #if ret == 0 :
        #    tkMessageBox.showinfo( "","Fil processeret")
        #else :
        #    tkMessageBox.showinfo( "","Fil fejlet")


    def close_top(self):
        self.destroy()
    
    def on_button(self):
        print self.entry.get()
        
    

if __name__ == "__main__":

    app = gui()
    app.mainloop()