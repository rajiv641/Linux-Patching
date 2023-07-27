import xlrd,xlwt
import json
from datetime import date
import time
import re
import pandas as pd
from datetime import datetime
import os
import sys
from matplotlib import pyplot as plt
import matplotlib
import seaborn as sns
import openpyxl
from xlutils.copy import copy as xl_copy
from xls2xlsx import XLS2XLSX
import numpy as np


colour_code = {'green': 0x0B, 'yellow': 0x0D, 'red': 0x0A}

file_json = None
accountno = None
msrfile = None
dict_plat = {}
dict_device = {}
dict_plat_temp = {}
dc_device = {}


os.system("clear")
accountno = input ("Enter the Account Number : ")
if (accountno):
    ret = os.system("curl --insecure -H \"x-auth-token: $(ht credentials)\" https://stepladder.rax.io/api/patching_portal/v1/report/linux/"+ accountno +"  > "+ accountno + ".json 2>/dev/null")
    if (ret != 0):
        print ("Couldn't download the JSON File, Exiting.......")
        exit()
    else:
        file_json = accountno + ".json"
        print (file_json)


msrfile = input ("Pls Enter the MSR file in xlsx format : ")
print (msrfile)



loc = None
ret = os.system("python3 url.py")
if (ret == 0):
    loc = 'EOL.xlsx'
else:
    print ("couldn't create EOL.xlsx, Exiting.............")
    exit()


xl = pd.ExcelFile(loc)

if (msrfile):
    xl2 = pd.ExcelFile(msrfile)
    df = xl2.parse("linuxServersall",header=None)
    df1 = df.tail(-1)
    os_plat = df1[14].unique() ## column 'O'
    dict_plat = df1.to_dict()
    ###########  mapping of devices and the Platform they belong ###########
    for x,v in dict_plat.items():
        if (x == 0):
            dict_device.update(v)
        if (x == 14):
            dict_plat_temp.update(v)
        if (x ==2):
            dc_device.update(v)
else:
    print ("MSR File not entered, Exiting..........")
    exit(-1)


def style_pattern(colour):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = True
    style.font = font
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # yellow 0x0D ##0x0B (bright green)##
    pattern.pattern_fore_colour = colour_code[colour]
    style.pattern = pattern
    return style


book_linux = xlwt.Workbook(encoding="utf-8")


def workbook_linux(name):
    style = style_pattern('yellow')
    sheet_linux = book_linux.add_sheet(name)
    sheet_linux.write(0, 0, "Device Number", style=style)
    sheet_linux.write(0, 1, "Patching Group", style=style)
    sheet_linux.write(0, 2, "Kernel Packages", style=style)
    sheet_linux.write(0, 3, "Count_Kernel Pkg", style=style)
    sheet_linux.write(0, 4, "Available Packages", style=style)
    sheet_linux.write(0, 5, "Count Pkg", style=style)
    sheet_linux.write(0, 6, "Last Pkg install date", style=style)
    sheet_linux.write(0, 7, "Distribution", style=style)
    sheet_linux.write(0, 8, "Kernel", style=style)
    sheet_linux.write(0, 9, "Last_Boot", style=style)
    sheet_linux.write(0, 10, "Uptime", style=style)
    sheet_linux.write(0, 11, "Version", style=style)
    sheet_linux.write(0, 12, "Needs Reboot", style=style)
    sheet_linux.write(0, 13, "EOL", style=style)
    sheet_linux.write(0, 14, "EOL Date", style=style)
    sheet_linux.write(0,15,"Platform",style=style)
    sheet_linux.write(0,16,"Uptime in Days",style=style)
    sheet_linux.write(0,17,"DataCenter",style=style)
    return sheet_linux




row = 1
name = 'IPSOS'
date_eol = ''

with open(file_json, 'r') as f:
    json_data = f.read()

sheet = workbook_linux(name)
date_format = xlwt.XFStyle()
date_format.num_format_str = 'dd/mm/yyyy'

def eol(distribution,version):
    global date_eol
    if ( distribution != 'Ubuntu'):
        x = version.split('.')
        ver_ipsos = x[0]
    else:
        ver_ipsos = version
    
    dt_today = datetime.today()  
    sec_today = dt_today.timestamp()

    for idx, name in enumerate(xl.sheet_names):
        sheet = xl.parse(name)
        for row in sheet.itertuples():
            if (pd.isna(row[1]) != True):
                dist = row[1]
                if (dist == distribution and row[2] == float(ver_ipsos)):
                    datetime_str = row[3]
                    date_eol = datetime_str.date()
                    sec_eol= datetime_str.timestamp()
                    if (sec_eol > sec_today):
                        return True
                elif ( distribution == 'RedHat' and int(ver_ipsos) <= 3):
                    return False
    return False

dictionary = json.loads(json_data)
for device in dictionary:
    val = dictionary[device]
    patching_group = 'NA'
    updates = []
    updates_kernel = []
    count = 0
    count_kernel = 0
    last_pkg_install_date = 'NA'
    distribution = 'NA'
    kernel = 'NA'
    lastb = 'NA'
    uptme = 'NA'
    version = 'NA'
    need_reboot = 'False'
    uptime_days = 'NA'
    platform = 'NA'
    dc = 'NA'

    for key1, val1 in val.items():
        if (key1 == 'patching_group'):
            patching_group = val1
        elif (key1 == 'updates'):
            if val1 is not None:
                for x in val1:
                    if (re.search(r"^(kernel)",x)):
                        updates_kernel.append(x)
                        updates_kernel.append(',')
                        count_kernel = count_kernel + 1
                    else:
                        updates.append(x)
                        updates.append(',')
                        count = count + 1
        elif (key1 == 'last_pkg_install_date'):
            last_pkg_install_date = val1
        elif (key1 == 'distribution'):
            distribution = val1
        elif (key1 == 'kernel'):
            kernel = val1
        elif (key1 == 'last_boot'):
            lastb = val1
        elif (key1 == 'version'):
            version = val1
        elif ( key1 == 'Datacenter'):
            dc = val1
        elif (key1 == 'needs_reboot'):
            if val1 is not None:
                need_reboot = val1
            else:
                need_reboot = 'Null'
        elif (key1 == 'uptime'):
            uptme = val1
            uptime_days = uptme.split(',')[0]

    trap = None
    ############ Find the Platform the Device belongs & Update the Excel Sheet #################
    for x,v in dict_device.items():
        if (str(v) == str(device)):
            trap = str(x)
            break

    for x2,v2 in dict_plat_temp.items():
        if ( str(x2) == str(trap)):
            platform = v2
            break

    for x,v in dc_device.items():
        if (str(x) == str(trap)):
            dc = v
            break
    
    sheet.write(row, 0, device)
    sheet.write(row, 1, patching_group)
    sheet.write(row, 2, updates_kernel)
    sheet.write(row, 3, count_kernel)
    sheet.write(row, 4, updates)
    sheet.write(row, 5, count)
    sheet.write(row, 6, last_pkg_install_date)
    sheet.write(row, 7, distribution)
    sheet.write(row, 8, kernel)
    sheet.write(row, 9, lastb)
    sheet.write(row, 10, uptme)
    sheet.write(row, 11, version)
    sheet.write(row, 15, platform)
    sheet.write(row, 16, uptime_days)
    sheet.write(row, 17, dc)

    if ( need_reboot == True ):
        style_green = style_pattern('green')
        sheet.write(row, 12, need_reboot,style = style_green)
    elif ( need_reboot == False):
         sheet.write(row, 12, need_reboot)
    else:
         sheet.write(row, 12, need_reboot)
    if (eol(distribution,version)):
        style_green = style_pattern('green')
        sheet.write(row,13,'Supported',style = style_green)
    else:
        style_red = style_pattern('red')
        sheet.write(row,13,'Expired',style=style_red)
    sheet.write(row, 14, date_eol,date_format)
    row = row + 1


today = date.today()
seconds = time.time()
local_time = time.ctime(seconds)
str1 =  "IPSOS.xls" # + "--" + format(today) + "--" + format(local_time) + ".xlsx"
book_linux.save(str1)

from xls2xlsx import XLS2XLSX
x2x = XLS2XLSX("IPSOS.xls")
x2x.to_xlsx("IPSOS.xlsx")

print ("""
        1. Device Number 
        2. Patching Group 
        3. Kernel Packages
        4. Count_Kernel Pkg 
        5. Available Packages 
        6. Count Pkg 
        7. Last Pkg install Date 
        8. Distribution  
        9. Kernel 
        10. Last_Boot 
        11. Uptime 
        12. Version
        13. Needs Reboot 
        14. EOL 
        15. EOL Date  
        16. Platform   
        17. Uptime in Days 
        18. DataCenter
        """)

xaxis = int(input ( "please enter the Field number for X-Axis ( default Platform 16) : "))
if (type(xaxis) != int or (xaxis < 1) or (xaxis > 18)):
    xaxis = 16

yaxis = int(input("please enter the Field number for Y-Axis ( default Distribution 8) : "))
if (type(yaxis) != int or (yaxis < 1) or (yaxis > 18)):
    yaxis = 8

choice = { 1:'Device Number',
        2 :'Patching Group',
        3: 'Kernel Packages',
        4: 'Count_Kernel Pkg',
        5: 'Available Packages',
        6: 'Count Pkg',
        7: 'Last Pkg install Date',
        8: 'Distribution',
        9: 'Kernel',
        10: 'Last_Boot',
        11: 'Uptime',
        12: 'Version',
        13: 'Needs Reboot',
        14: 'EOL',
        15: 'EOL Date',
        16: 'Platform',
        17: 'Uptime in Days',
        18: 'DataCenter'
        }

x=[]
y=[]
x.append('df1.' + choice[xaxis])
y.append('df1.'+ choice[yaxis])

xl2 = pd.ExcelFile(str1)
df = xl2.parse("IPSOS") #,header=None)
df1 = df.tail(-1)

pivot = pd.crosstab(index=eval(x[0]), columns=eval(y[0]), margins=True,margins_name='Grand Total')
gr = pivot.T.plot(kind='bar', xlabel = choice[xaxis], ylabel=choice[yaxis],figsize=(25,10),rot=30,fontsize=18)
fig = gr.get_figure()
fig.savefig("pivot.png")

rb = openpyxl.load_workbook('IPSOS.xlsx')
writer = pd.ExcelWriter('all_IPSOS.xlsx', engine = 'openpyxl')
writer.workbook = rb
pivot.to_excel(writer,sheet_name='Pivot-Table')
df.to_excel(writer,sheet_name='IPSOS')
writer.close()




os.system("python3 graph.py  "+ msrfile)

