import pandas as pd
from openpyxl import load_workbook



inputFile = "validation_folder_input\M1_IBS_ItemNames.xlsx"
sheetname="Sheet1"

referenceExcelFile = "valid\\newBM.xlsx"
bm_sheetname="Sheet1"

cells_to_check=["E9","F9","G9","H9",
                "E12","F12","G12","H12",
                "E13","F13","G13","H13",
                "E29","F29","G29","H29",
                "E31","F31","G31","H31",]

final_cells_to_check =  ["E11","F11","G11","H11",
                         "E15","F15","G15","H15",
                         "E33","F33","G33","H33",]      



def load_bm_file():
    df = pd.read_excel(referenceExcelFile)
    return df


def get_cell_value(initial_cell_val,data_frame):
    returned_value = "empty"
    for index, vall in data_frame.iterrows():
        if initial_cell_val == vall['item']:
            returned_value=vall['ValueNumeric']
            break

    return returned_value    

def check_file_if_valid(filename):
    wb_filetoVerify = load_workbook(filename)#,data_only=True)
    ws_toverify = wb_filetoVerify[sheetname]
    bmDF = load_bm_file()
    
    for cell in cells_to_check:
        value_toWrite = get_cell_value(ws_toverify[cell].value,bmDF)
        if value_toWrite !="empty":
            ws_toverify[cell] = value_toWrite
        
    wb_filetoVerify.save("tmptmptmp.xlsx")
     

    #wb_modified = load_workbook("tmptmptmp.xlsx",data_only=True)
    #ws_modified = wb_modified[sheetname]
    #for finalcells in final_cells_to_check:
    #    print(ws_modified[finalcells].value)



if __name__ == '__main__':
    check_file_if_valid(inputFile)


