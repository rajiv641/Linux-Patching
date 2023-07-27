#!/usr/bin/env python3

import pandas as pd
import xlsxwriter
import datetime
import re

book = xlsxwriter.Workbook('EOL' + ".xlsx")   
worksheet = book.add_worksheet('EOL')
date_format = book.add_format({'num_format': 'dd/mm/yyyy'})


def worksheet_entry(row,osname='',eol=''):
    name = None
    ver = None
    if osname:
        temp = osname.split()
        if (temp[0] != 'Microsoft' and temp[0] != 'Windows'):
            osname = osname.rstrip('*')
            if (temp[0] == 'Red'):
                name = 'RedHat'
            elif(temp[0] == 'Oracle'):
                name = 'OracleLinux'
            elif (temp[0] == 'Rocky'):
                name = 'RockyLinux'
            else:
                name = temp[0]

            if (temp[0] == 'Ubuntu'):
                ver = re.search('\d+\.\d+',osname)
            else:
                ver = re.search('[0-9]+', osname)

            worksheet.write(row,0,name)
            worksheet.write(row,1,ver.group())
            
            if (eol):
                date_line = eol.split(',')
                year = date_line[1]
                month_day = date_line[0]
                mon = month_day.split(" ")
                mname = mon[0]
                mname = mname.lstrip('~')
                day = mon[1]
                mnum = datetime.datetime.strptime(mname, '%B').month
                datetime_object = datetime.datetime(int(year),int(mnum),int(day))
                eol = str(day) + "/" + str(mnum)+ "/" + str(year).strip()
                date_string = datetime_object.strptime(eol,"%d/%m/%Y").date()
         
                worksheet.write(row,2,date_string,date_format)


 
url = "https://www.rackspace.com/information/legal/eolterms"
dfs = pd.read_html(url)
mydict = {}
osname = ''
eol = ''
icount = 0


for rows in dfs[1].iterrows():
    for x in rows:
        if (not isinstance(x,int)):
            mydict = x.to_dict()
            for key,value in mydict.items():
                if (key == 'Unnamed: 0'and icount > 0):
                    osname = value
                elif (key == 'End of Life' and icount > 0 and pd.isna(value) != True):
                    eol = value
                    worksheet_entry(icount - 1,osname,value)
            icount = icount + 1

book.close()
