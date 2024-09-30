import  os
import pandas as pd
from openpyxl import Workbook
import xlrd
os.system('cls')
COLORS = {\
"Black": "\u001b[30m",
"Red": "\u001b[31m",
"Green": "\u001b[32m",
"Yellow": "\u001b[33m",
"Blue": "\u001b[34m",
"Magenta": "\u001b[35m",
"Cyan": "\u001b[36m",
"White": "\u001b[37m",
"Cyan-background":"\u001b[46m",
"Black-background":"\u001b[40m",
}
def colorText (text):
    for color in COLORS:
        text = text.replace("[["+ color+"]]",COLORS[color])
    return text
f = open ("logo.txt","r")
ascii = "".join(f.readlines())
print(colorText(ascii))
data_file_folder= '.\PLIKI'
print ("Enter a name for the sheets to be joined:")
nazwa_arkusza=input()
print ("Enter the name of the workbook to be created:")
nazwa_skoroszytu=input()
roz=(".xlsx",".xls")
df = []
dp=[]
for file in os.listdir(data_file_folder):
    spr=pd.ExcelFile(os.path.join(data_file_folder,file))
    if nazwa_arkusza in spr.sheet_names:
        print('loading a file {0}...'.format(file)+"   OK")
        df3=(pd.read_excel(os.path.join(data_file_folder, file), sheet_name=nazwa_arkusza))
        df2 = df3.assign(HARRNESS=file,FOLDER=data_file_folder)
        df.append(df2)
    else:
        print('loading a file {0}...'.format(file)+"   NOK")

if len(df)>0:
    df_master = pd.concat(df, axis=0)
    df_master.to_excel(nazwa_skoroszytu+'.xlsx', index=False)
    print ("Workbook " +nazwa_skoroszytu+" created")
else:
    print("Not found sheet "+ nazwa_arkusza)
