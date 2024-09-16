import pandas as pd
import numpy as np
import os
import win32com.client
import win32con
import win32gui
import glob
from pathlib import Path

files = glob.glob("*.rpy")
olds = []
news = []
dfs = []
place = []

languages = input("Enter the name of language:")

for file in files:
    ids = []
    sourcetxt = []
    trans = []
    chr = []

    stringsflag = False
    f = open(file,'r',encoding='utf_8',errors='ignore')

    sfile = f.read()
    sfile2 = sfile.splitlines()
    for i in range(len(sfile2)):
        if stringsflag == False:
            if f"translate {languages}" in sfile2[i]:
                if "strings:" in sfile2[i]:
                    stringsflag = True
                    continue
                a = sfile2[i].replace("translate Japanese ","")
                ids.append(a)
                a = sfile2[i+2].replace("    # ", "")
                sourcetxt.append(a)
                a = sfile2[i+3].replace("    ", "")
                if len(a.split('"')) > 1:
                    b = a.split('"')[1]
                    c = a.split('"')[0]
                    trans.append(b)
                    chr.append(c)
                else:
                    chr.append("")
                    trans.append(a.replace('"',""))
        else:
            if "   old " in sfile2[i]:
                try:
                    place.append(sfile2[i-1].replace("   ",""))
                except:
                    place.append("unknown")
                a = sfile2[i].replace("    old ", "")
                a = a.replace("\"", "")
                olds.append(a)
            elif "   new " in sfile2[i]:
                a = sfile2[i].replace("    new ", "")
                a = a.replace("\"","")
                news.append(a)

    b = file.replace(f"./{languages}/", "")
    b = b.replace(".rpy", "")
    dfs.append(pd.DataFrame(np.vstack([[ids], [sourcetxt],[chr], [trans]]).T))

df = pd.DataFrame(np.vstack([[place],[olds], [news]]).T)
df.to_excel(f'{languages}.xlsx',sheet_name="Strings",index=False, header=["place","Original","Translated"])
with pd.ExcelWriter(f'{languages}.xlsx',mode='a') as writer:
    for i in range(len(dfs)):
        print(f"Processing {files[i]} ({i+1}/{len(dfs)})")
        dfs[i].to_excel(writer,sheet_name=files[i].replace(".rpy",""),index=False,header=["ID","Original","Chara","Translated"])

print("Starting Excel...")

xlApp = win32com.client.Dispatch("Excel.Application")
xlApp.Visible = True
wb = xlApp.Workbooks.Open(os.path.abspath(f'{languages}.xlsx'))

print("Setting the Format...")
try:
    for i in range(len(dfs)):
        ws = wb.Worksheets(i+1)
        ws.Activate()
        ws.Cells.WrapText = True
        if i == 0:
            ws.Range("A1").ColumnWidth = 20
            ws.Range("B1").ColumnWidth = 40
            ws.Range("C1").ColumnWidth = 40
        else:
            ws.Range("A1").ColumnWidth = 10
            ws.Range("B1").ColumnWidth = 40
            ws.Range("D1").ColumnWidth = 40
    wb.Save()
    wb.Close()
    xlApp.Quit()
except Exception as e:
    print(str(e))

print(f"Created Spreadsheet as {languages}.xlsx")
