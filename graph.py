import xlrd
import xlwt
import xlsxwriter
import re
import pandas as pd
import importlib
from matplotlib import pyplot as plt
import matplotlib
import sys
import numpy as np
import seaborn as sns
import plotly.express as px
import plotly.io as pio
import os

loc =''
os.system("clear")

n = len(sys.argv)
if (n < 2):
    print("please provide the File name, exiting............")
    exit()
else:
    loc = sys.argv[1]

if (os.path.exists(loc) == False):
    print (format(loc) + " : File doesn't exist, Exiting....")
    exit()

xl = pd.ExcelFile(loc)
df = pd.DataFrame()
os = []
countos = []
ostype = []
platform = []


book = xlsxwriter.Workbook("temp.xlsx")
worksheet = book.add_worksheet('stack')
worksheet.write(0,0,"platform")
worksheet.write(0,1,"ostype")
worksheet.write(0,2,"count")


df = xl.parse("linuxServersall",header=None)
df1 = df.tail(-1)
answer = df1[9].value_counts()
stack = df1[14].value_counts()
os_plat = df1[14].unique() ## column 'O' #
os_version = df1[9].unique() ## column 'J' Base channel ##
temp_ostype = []
temp_count = []

stack_len = 1


raj = df1[[14,9]].value_counts().reset_index(name='count')
hey = pd.melt(raj,id_vars=[14,9],value_vars=['count'],value_name='occurances')#var_name='9', value_name='value')
print (format(hey))
fig = px.bar(hey, x=14, color=9,y='occurances',barmode='overlay',title="Linux Version split across Platforms",height=600,labels={"14":'Platform',"9":'Linux Version'})
config = {
    'toImageButtonOptions': {
        'format': 'png', # one of png, svg, jpeg, webp
        'filename': 'custom_image',
        'height': 500,
        'width': 700,
        'scale':8 # Multiply title/legend/axis/canvas sizes by this factor
        }
    }

fig.show(config=config)
pio.write_html(fig,"Linux_Stacked.html")

print ("="*80)

for i,v in answer.items():
    os.append(i)
    countos.append(v)



data_os = pd.DataFrame({"os":os, "countos":countos})
plt.figure(figsize=(20, 10))
splot=sns.barplot(x="os",y="countos",data=data_os)
plt.xlabel("Linux Version", size=16)
plt.ylabel("Number of Instances", size=16)
plt.bar_label(splot.containers[0],size=16)
plt.xticks(rotation=10, horizontalalignment="center")


#Two  lines to make our compiler able to draw:
plt.savefig("os.png")
sys.stdout.flush()
book.close()
