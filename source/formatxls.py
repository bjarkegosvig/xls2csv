#!/usr/bin/env python
# -*- encoding: utf-8 -*-
import win32com.client 
from os import sys
import os
from datetime import datetime
import time

class formatxls :
    """
    excel_file: is the file path to work on
    headerline_number: is the line number which the header starts on 
    del_columns: is the columns to delete must be a list
    """
    #def __init__(self, excel_file, headerline_number, del_columns):
    def __init__(self, excel_file, header_start_cell):
        self.filename       = excel_file
        dir = os.path.realpath('.')
        self.tmpfilename    = os.path.join(dir, 'tmp.xls')   
        self.header_cell    = header_start_cell
        self.headerline     = 0
        self.del_columns    = 'A'
        self.excel          = win32com.client .gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible  = 0
        self.count          = 0

    def _find_header_pos(self):
        length = len(self.header_cell)
        if length < 2:
            return 1
        else :
          
            self.headerline = int(self.header_cell[1:])
            self.del_columns = self.header_cell[0].upper() # uppercase
            print self.headerline
            print self.del_columns
            return 0

        
        
    def _open_workbook(self, file):
        try : 
            workbook=self.excel.Workbooks.Open(file) 
            self.count = workbook.Sheets.Count
            return workbook
        except:
            print "Failed to open spreadsheet " + str(file)
            sys.exit(1)
    
    def _close_workbook(self,workbook):
        workbook.Save()
        workbook.Close(SaveChanges=True) 
        self.excel.Application.Quit()
    
    def _copy_workbook(self,workbook):
        workbook.SaveAs(self.tmpfilename)
        
    
    def _del_title(self, workbook):
    # work on each sheet
        sheet_num = 0
        for sheet in workbook.Sheets :
            sheet.Select()
            #keep only headerline in first sheet
            sheet_num += 1
            if sheet_num == 1:
                offset = 1
            else:
                offset = 0
            
            # delete title line and empty lines before header line
            for n in range(0,self.headerline - offset ): 
                self.excel.Rows(1).Select() #delete line 1 n times
                self.excel.Selection.Delete() 


    def _del_columns(self,workbook):
        
        if self.del_columns == 'A':
            return
        else :
            col = []
            beginNum = ord('A')
            endNum = ord(self.del_columns)
            for number in xrange(beginNum, endNum):
                 col.append( chr(number) )
            for sheet in workbook.Sheets :
                sheet.Select()
                for column in col: 
                   
                    self.excel.Columns(column).Select() 
                    self.excel.Selection.Delete() 
    
    def process_workbook(self):
        # copy the excelfile to a tmp file
        workbook = self._open_workbook(self.filename)
        self._copy_workbook(workbook)
        self._close_workbook(workbook)
        # delete row and cols in the tmp file
        a = self._find_header_pos()
        print a
        tmp_workbook = self._open_workbook(self.tmpfilename)
        #if self.header_cell != 'A1':
        self._del_title(tmp_workbook)
        self._del_columns(tmp_workbook)
        self._close_workbook(tmp_workbook)
    

        
#a = formatxls(r'C:\Users\bjagos\Desktop\New folder\new.xls','b4')
#a.print_me()
#a.process_workbook()





