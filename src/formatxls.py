#!/usr/bin/env python
# -*- encoding: utf-8 -*-
import win32com.client 
from os import sys
import os
from datetime import datetime
import time

class formatxls :
    """
    excel_file:         string which contains the file path to work on
    header_start_cell:  string which contains the cell where the headerline starts
    one2one:            bool which indicates if each sheet must be one csv file
    oneheader:          bool which indicates if there is only a headerline on the first 
                        sheet and data starts in A1 for the rest of the sheets
    """
    def __init__(self, excel_file, header_start_cell = 'A1', one2one = 0, oneheader = 0):
        self.filename       = excel_file
        dir = os.path.realpath('.')
        self.tmpfilename    = os.path.join(dir, 'tmp.xls')   
        self.header_cell    = header_start_cell
        self.one2one        = one2one
        self.oneheader      = oneheader
        self.headerline     = 0
        self.del_columns    = 'A'
        self.excel          = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible  = 0
        self.count          = 0

    def _find_header_pos(self):
        length = len(self.header_cell)
        if length < 2:
            return 1
        else :
          
            self.headerline = int(self.header_cell[1:])
            self.del_columns = self.header_cell[0].upper() # uppercase
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
            sheet_num += 1
            
            # keep only headerline in first sheet except if we want one sheet to be one csv, 
            # then we want header on all csv files.
            # the last thing is if we only have header on the first sheet then there
            # is no header line to delete
            if sheet_num == 1 or self.one2one == 1:
                offset = 1
            elif sheet_num > 1 and self.oneheader == 1:
                return
            else:
                offset = 0
            
            #select the current sheet to work on
            sheet.Select()           
            # delete title line and empty lines before header line
            for n in range(0,self.headerline - offset ): 
                self.excel.Rows(1).Select() #delete line 1 n times
                self.excel.Selection.Delete() 


    def _del_columns(self,workbook):
        
        sheet_num = 0
        if self.del_columns == 'A':
            return
        else :
            # find all columns to delete
            col = []
            beginNum = ord('A')
            endNum = ord(self.del_columns)
            for number in xrange(beginNum, endNum):
                 col.append( chr(number) )
            
            # work on each sheet
            for sheet in workbook.Sheets :
                sheet_num += 1
            
                # don't delete anything after sheet 1 if we only have a header in sheet one
                # and data starts in A1 for the rest of the sheets
                if sheet_num >= 2 and self.oneheader == 1:
                    return
             
                # select the current sheet to work on, and delete the necessary columns
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
        self._find_header_pos()
        tmp_workbook = self._open_workbook(self.tmpfilename)
        self._del_title(tmp_workbook)
        self._del_columns(tmp_workbook)
        self._close_workbook(tmp_workbook)
 