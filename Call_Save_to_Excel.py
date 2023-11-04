# SAVES THE TRACKS FROM A PLAYLIST OR THE WHOLE LIBRARY INTO AN EXCEL FILE
# MORE THAN ONE PLAYLIST CAN BE SELECTED
# WHEN CREATING AN EXCEL OUTPUT FOR THE LIBRARY, A NUMBER OF TRACKS CAN BE SPECIFIED.
# IF NOT SPECIFIED, IT'S ALL THE TRACKS.
from os.path import exists
import Read_PL
import WMP_Read_PL
import pandas as pd
# import Files

def file_w_ext(path):
    pos = path.rfind("\\")+1
    file_w_ext = path[pos:]
    return file_w_ext

# MAIN CODE
def Save_Excel(PL_name=None,PL_nbr=None,Do_lib=False,rows=None,iTunes=True):
    # CALLS Read_PL FUNCTION ,Do_lib=True,rows=10
    # col_names =  ["Arq","Art","Title","AA","Album","Genre","Covers","Year"]
    col_names =  ["Arq","Art","Title","Year","Covers"]
    
    if iTunes:
       # EXCLUDE INVALID TAGS 
       col_names = [x for x in col_names if x in Read_PL.order_list_itunes]
       dict = Read_PL.Read_PL(col_names,PL_name=PL_name,PL_no=PL_nbr,Do_lib=Do_lib,rows=rows)
    else:
        # EXCLUDE INVALID TAGS 
        col_names = [x for x in col_names if x in WMP_Read_PL.order_list_wmp]
        dict = WMP_Read_PL.Read_WMP_PL(col_names,PL_name=PL_name,PL_no=PL_nbr,Do_lib=Do_lib,rows=rows)   
    # ASSIGNS
    # App = dict["App"]
    # playlists = dict["PLs"]
    df = dict["DF"]

    # KEEP ONLY SELECTED COLS.
    df = df.loc[:, col_names]
    
    # RENAME THE COLUMNS HEADERS
    if "Arq" in col_names:
        df.loc[:, "File"] = df["Arq"].apply(file_w_ext)
        df = df.rename(columns={"Arq": "Location" })

    # SAVE TO EXCEL FILE:
    user_inp = input("\nOutput name (file will be saved to D:\iTunes\Excel\\all.xls): ")
    if user_inp == "":
       file_nm = "D:\\iTunes\\Excel\\all.xlsx"
    else:
        file_nm = "D:\\iTunes\\Excel\\" + user_inp + ".xlsx"
    
    # NAMES SHEET
    if iTunes:
       sheet = "iTunes"
    else:
        sheet = "WMP"     
    # save the dataframe to an Excel file
    df.to_excel(file_nm, sheet_name=sheet, index=False)
    # print("Hello, " + file_nm + "!")

# CHAMA PROGRAM PL_name="ALL",Fave-iPhone
Save_Excel(PL_name="Fave-Tags",Do_lib=0,rows=300,iTunes=0)