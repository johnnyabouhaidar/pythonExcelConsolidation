import os
import pandas as pd
from openpyxl import load_workbook


list = [
    [1,2,3,4,5],
    [3,5,6,3,6],
    [2,3,45,6,7],
]

df =pd.DataFrame(list)

products_list = df.values.tolist()

print(products_list)

wb = load_workbook("tmptmp22.xlsx")
ws_write = wb["Sheet1"]

ws_write["C6":"G8"]=products_list

wb.save()