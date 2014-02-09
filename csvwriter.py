#!/usr/bin/env python
# -*- encoding: utf-8 -*-


import xlrd
import csv

class csvwriter :
    """
    excel_file: is the file path to work on
    """
    def __init__(self, excel_file):
        self.filename       = excel_file
        self.csv_name       = ''.join([excel_file[:-4],'.csv'])
        self.date_format    = "%d/%m/%Y"
        #self.date_format   = "%Y%m%d"
    
    def _xlsallsheet2onecsv(self):
        workbook = xlrd.open_workbook(self.filename)
        all_worksheets = workbook.sheet_names()
        your_csv_file = open(csv_name, 'wb')
        wr = csv.writer(your_csv_file, dialect='excel', delimiter=';',  quoting=csv.QUOTE_NONE)

        for worksheet_name in all_worksheets:
            worksheet = workbook.sheet_by_name(worksheet_name)
            for rownum in xrange(worksheet.nrows):
                row = []
                for entry in worksheet.row(rownum):
            
                    # encode dates 
                    if entry.ctype == xlrd.XL_CELL_DATE:
                        a1_tuple = xlrd.xldate_as_tuple(entry.value , workbook.datemode)
                        a1_datetime = datetime(*a1_tuple)
                        tmp = unicode(a1_datetime.strftime(self.date_format)).encode(encoding)
                    # encode integers
                    elif entry.ctype == xlrd.XL_CELL_NUMBER:
                        if Decimal(entry.value)._isinteger():
                            tmp = unicode(int(entry.value)).encode(encoding)
                        # encode float
                        else:
                            tmp = unicode(entry.value).encode(encoding)
                    # encode text
                    else:
                        tmp = unicode(entry.value).encode(encoding)
                    row.append(tmp.replace('\n','').replace(';','').replace('"',''))
                    
                if any(row): # don't write empty rows
                    wr.writerow(row)
        your_csv_file.close()
        if your_csv_file.closed:
            return 0
        else :
            return 1 # file didn't close proper   