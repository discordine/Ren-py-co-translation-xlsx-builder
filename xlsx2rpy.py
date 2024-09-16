import pandas as pd
import numpy as np
import sys
import glob

file = glob.glob('*.xlsx')

language = input("Input the name of the language:")

print("Initializing...")

bk = pd.ExcelFile(file[0])

for i in range(len(bk.sheet_names)):
    dfs = pd.read_excel(file[0], sheet_name=i, index_col=None)
    print(f"Processing {bk.sheet_names[i]}.rpy ({i+1}/{len(bk.sheet_names)})")
    if i == 0:
        with open(f"{bk.sheet_names[i]}.rpy", 'w',encoding="utf8") as f:
            f.write(f"translate {language} strings:\n")

        with open(f"{bk.sheet_names[i]}.rpy", 'a',encoding="utf8") as f:
            for j in range(len(dfs)):
                olds = dfs.iloc[j, 1]
                news = dfs.iloc[j, 2]
                f.write('   old"' + olds + '"\n')
                f.write('   new "' + news + '"\n\n')
    else:
        with open(f"{bk.sheet_names[i]}.rpy", 'w',encoding="utf8") as f:
            f.write(f"# {bk.sheet_names[i]}.rpy\n")
        with open(f"{bk.sheet_names[i]}.rpy", 'a',encoding="utf8") as f:
            for j in range(len(dfs)):
                ids = dfs.iloc[j, 0]
                ortext = dfs.iloc[j, 1]
                chr = dfs.iloc[j, 2]
                transs = dfs.iloc[j, 3]
                postfunctions = ""
                if dfs.shape[1] == 5:
                    if type(dfs.iloc[j, 4]) is str:
                        postfunctions = dfs.iloc[j, 4]
                f.write(f'translate {language} ' + ids + '\n')
                if pd.isnull(chr):
                    f.write(f"   # {ortext}\n")
                    f.write(f'   "' + transs + '" ')
                else:
                    f.write(f"   # {ortext}\n")
                    f.write(f'   {chr} "' + transs + '" ')
                if not postfunctions == "":
                    print(postfunctions)
                    f.write(postfunctions)
                f.write("\n\n")



