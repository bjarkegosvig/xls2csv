#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import xlrd
import csv
import os
from decimal import Decimal
from datetime import datetime


class csvwriter :
    """
    excel_file: string which contains the file path to work on
    dateformat: string containing the dateformat e.g. "%d/%m/%Y"
    encoding:   string containg the encoding format e.g. utf-8 or cp1252 etc.
    abformat:   bool which indicates if row 0 or 1 should be filled with an asending number    
    one2one:    bool which indicates if each sheet must be one csv file
    delimiter:  string containg the delimiter e.g. ';' or ',' or '\t' (tab)    

    """
    def __init__(self, excel_file, dateformat= "%d/%m/%Y", encoding = "cp1252", 
                 abformat = 0 , one2one = 0, delimiter = ';'):
        dir = os.path.realpath('.')
        self.filename    = os.path.join(dir, 'tmp.xls')        
        self.date_format = dateformat
        self.encoding    = encoding
        self.abformat    = abformat
        self.one2one     = one2one
        self.delimeter   = delimiter
        # test for xls or xlsx file
        if 'xlsx' in excel_file:
            self.csv_name    = ''.join([excel_file[:-5],'.csv'])
        elif 'xls' in excel_file:
            self.csv_name    = ''.join([excel_file[:-4],'.csv'])
        else:
            self.csv_name    = ''.join([excel_file,'.csv'])
     
    def xlsallsheet2onecsv(self):
        
        workbook = xlrd.open_workbook(self.filename)
        all_worksheets = workbook.sheet_names()
        your_csv_file = open(self.csv_name, 'w', newline='', encoding = self.encoding  )
        wr = csv.writer(your_csv_file, delimiter=self.delimeter,  quoting=csv.QUOTE_NONE)
        i = 1
        n = 1
        row = []
        row_num = 0
        for worksheet_name in all_worksheets:
            worksheet = workbook.sheet_by_name(worksheet_name)
            for rownum in range(worksheet.nrows):
                del row[:]
                for entry in worksheet.row(rownum):
            
                    # encode dates 
                    if entry.ctype == xlrd.XL_CELL_DATE:
                        a1_tuple = xlrd.xldate_as_tuple(entry.value , workbook.datemode)
                        a1_datetime = datetime(*a1_tuple)
                        tmp = a1_datetime.strftime(self.date_format)
                    # find integers
                    elif entry.ctype == xlrd.XL_CELL_NUMBER:
                        if float(entry.value).is_integer():
                            tmp = int(entry.value)
                        else:
                            # float with two decimals
                            tmp = "{0:.2f}".format(float(entry.value))
                    # the rest is text
                    else:
                        
                        tmp = entry.value
                        tmp = tmp.replace('\n','').replace(self.delimeter,'').replace('"','')                     
                    row.append(tmp)
                   
                    #row.append(tmp.replace('\n','').replace(';','').replace('"',''))
                    
                if any(row): # don't write empty rows
                   # do we want column A and B to be filled with asending numbers from 1 to xxx
                    if self.abformat :
                        #don't add number on headerline
                        if row_num == 0:
                            wr.writerow(row)
                            row_num = 1
                        else :
                            row[0] = int(i)
                            row[1] = int(i)              
                            wr.writerow(row)
                            i += 1
                    else:
                        wr.writerow(row)
            # if we want one csv file pr sheet close current file and open a new one 
            # but only if we havn't reached the last sheet
            if self.one2one and (int(worksheet.number) != ( int(workbook.nsheets) - 1 )):
                first = 0
                your_csv_file.close()
                if first == 0:
                    self.csv_name    = ''.join([self.csv_name[:-4],'_' ,str(n) ,'.csv'])
                    first == 1
                else:
                    self.csv_name    = ''.join([self.csv_name[:-6],'_' ,str(n) ,'.csv'])
                your_csv_file = open(self.csv_name, 'w', newline='', encoding = self.encoding  )
                wr = csv.writer(your_csv_file, delimiter=self.delimeter, quoting=csv.QUOTE_NONE)
                n += 1
                    
        your_csv_file.close()
        os.remove(self.filename)
        if your_csv_file.closed:
            return 0
        else :
            return 1 # file didn't close proper   