import os
import pandas as pd
from openpyxl import load_workbook


referenceExcelFile = "valid\\newBM.xlsx"
inputFolder = "validation_folder_input"

sheetname="Sheet1"
bm_sheetname="Sheet1"

cells_to_check=["E9","F9","G9","H9",
                "E12","F12","G12","H12",
                "E13","F13","G13","H13",
                "E29","F29","G29","H29",
                "E31","F31","G31","H31",]

final_cells_to_check =  ["E11","F11","G11","H11",
                         "E15","F15","G15","H15",
                         "E33","F33","G33","H33",]      

#for fileObj in os.listdir(inputFolder):
#    print(fileObj)
#    print(referenceExcelFile)
#    wb1 = load_workbook(os.path.join(inputFolder,fileObj))
#    ws1 = wb1[sheetname]
#    for cell in cells_to_check:
#        print(ws1[cell].value)

def load_bm_file():
    df = pd.read_excel(referenceExcelFile)
    return df

def check_file_if_valid(filename,reference_excel_file):
    wb_filetoVerify = load_workbook(filename)
    ws_toverify = wb_filetoVerify[sheetname]
    for cell in cells_to_check:
        print(ws_toverify[cell].value)

    bmDF = load_bm_file()


if __name__ == '__main__':
    check_file_if_valid("validation_folder_input\M1_IBS_ItemNames.xlsx",referenceExcelFile)


