#!/usr/bin/env python
# -*- encoding: utf-8 -*-
import win32com.client 
from os import sys
import os
from datetime import datetime

class excel2csv :
    """
    excel_file: is the file path to work on
    headerline_number: is the line number which the header starts on 
    del_columns: is the columns to delete must be a list
    """
    def __init__(self, excel_file, headerline_number, del_columns):
        self.filename       = excel_file
        self.tmpfilename    = 'tmp.xls'
        self.headerline     = headerline_number
        self.del_columns    = del_columns
        self.excel          = win32com.client .gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible  = 0
        self.count          = 0

        #self.date_format   = "%Y%m%d"

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
        for sheets in workbook.Sheets :
            # delete title line and empty lines before header line
            for n in range(self.headerline): 
                print n
                self.excel.Rows(n).Select() 
                self.excel.Selection.Delete() 
    
    def _del_columns(self,workbook):
        for sheets in workbook.Sheets :
            for column_name in self.del_columns :
                self.excel.Columns(column_name).Select() 
                self.excel.Selection.Delete() 
    
    def process_workbook(self):
        # copy the excelfile to a tmp file
        workbook = self._open_workbook(self.filename)
        self._copy_workbook(workbook)
        self._close_workbook(workbook)
        # delete row and cols in the tmp file
        tmp_workbook = self._open_workbook(self.tmpfilename)
        self._del_title(tmp_workbook)
        self._del_columns(tmp_workbook)
        self._close_workbook(tmp_workbook)
   
             

        
a = excel2csv(r'C:\Users\bjagos\Desktop\excel_test\new.xls',3,['A'])






