#%% Basic file create, open and save
f = open("sample.txt", "w") # mode => r, w, a
read_data = f.read()
print(read_data)
f.close()

#%% 2
with open("sample.txt", "w") as f:
    read_data = f.read()
    print(read_data)


#%% openpyxl file open
from openpyxl import load_workbook
wb = load_workbook('sample.xlsx')

#%% pandas file open
import pandas as pd
dfs = pd.read_excel('sample.xlsx')

