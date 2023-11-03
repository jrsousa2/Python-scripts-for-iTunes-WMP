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
def Save_Excel(PL_name_vl=None,PL_nbr=None,Do_lib_vl=False,rows_vl=None,iTunes=True):
    # CALLS Read_PL FUNCTION ,Do_lib=True,rows=10
    # col_names =  ["Arq","Art","Title","AA","Album","Genre","Covers","Year"]
    col_names =  ["Arq","Art","Title","Year"]
    if iTunes:
       dict = Read_PL.Read_PL(col_names,PL_name=PL_name_vl,PL_no=PL_nbr,Do_lib=Do_lib_vl,rows=rows_vl)
    else:
        dict = WMP_Read_PL.Read_WMP_PL(col_names,PL_name=PL_name_vl,PL_no=PL_nbr,Do_lib=Do_lib_vl,rows=rows_vl)   
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
Save_Excel(PL_name_vl="ZZZ",Do_lib_vl=1,rows_vl=100,iTunes=0)