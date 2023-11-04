# SYNCS PLAY COUNTS BETWEEN ITUNES AND WMP
from os.path import exists
import Read_PL
import WMP_Read_PL
import pandas as pd
# import Files

def file_w_ext(path):
    pos = path.rfind("\\")+1
    file_w_ext = path[pos:]
    return file_w_ext

# CALLS Read_PL FUNCTION ,Do_lib=True,rows=10
col_names =  ["Arq","Plays", "ID"]

# ONLY USED TO DEBUG, NOT BEING USED ANYMORE
def Save(df):
    # CREATES A NEW COLUMN WITH THE FILE NAME WO/ THE PATH
    # RENAME THE COLUMNS HEADERS
    if "Arq" in col_names:
        df.loc[:, "File"] = df["Location"].apply(file_w_ext)
        # df = df.rename(columns={"Arq": "Location" })

    # SAVE TO EXCEL FILE:
    file_nm = "D:\\iTunes\\Excel\\iTunes_vs_WMP.xlsx"
    # save the dataframe to an Excel file
    df.to_excel(file_nm, index=False)

def df_dedupe(source,df):
    # KEEP ONLY SELECTED COLS.
    # df = df.loc[:, col_names]
    
    # LOWERCASE THE FILE NAME
    df["Arq"] = df["Arq"].str.lower()

    start_rows = df.shape[0]
    print("\nThe",source," df has",df.shape[0],"tracks before deduping")

    # ADDS KEY COL. TO DF
    df["max_Plays_" + source] = df.groupby("Arq")["Plays"].transform("max")

    # Eliminate duplicate records based on "Arq" column
    df = df.drop_duplicates(subset="Arq", keep="first")

    end_rows = df.shape[0]
    print("\nThe",source,"df has",df.shape[0],"tracks after deduping")

    dict = {}
    dict["start_rows"] = start_rows
    dict["end_rows"] = end_rows
    dict["df"] = df
    return dict

# MAIN CODE
def Sync_plays(PL_name=None,PL_nbr=None,Do_lib=False,rows=None):
    # CALLS Read_PL FUNCTION ,Do_lib=True,rows=10

    # ITUNES DF
    print("\nReading the iTunes playlist")
    dict_iTu = Read_PL.Read_PL(col_names,PL_name=PL_name,PL_no=PL_nbr,Do_lib=Do_lib,rows=rows,Modify_cols=False)
    iTunes_App = dict_iTu['App']
    iTu_df = dict_iTu["DF"]
    # MAKE A COPY OF THE ORIGINAL PATH
    iTu_df["Location"] = iTu_df['Arq'].copy()
    # CHANGE DF-ELIMINATE DUPES
    dict = df_dedupe("iTunes",iTu_df)
    iTu_df = dict["df"]
    iTu_start_rows = dict["start_rows"]
    iTu_end_rows = dict["end_rows"]
    
    # WMP
    print("\nReading the WMP playlist")
    dict_wmp = WMP_Read_PL.Read_WMP_PL(col_names,PL_name=PL_name,PL_no=PL_nbr,Do_lib=Do_lib,rows=rows,Modify_cols=False) 
    wmp_df = dict_wmp["DF"]
    # WMP_media_coll = dict_wmp["media"]  
    WMP_PL = dict_wmp["PL"]
    
    # CHANGE DF-ELIMINATE DUPES
    dict = df_dedupe("WMP",wmp_df)
    wmp_df = dict["df"]
    wmp_start_rows = dict["start_rows"]
    wmp_end_rows = dict["end_rows"]
    
   
    # JOIN THE DATAFRAMES
    df = iTu_df[["Arq", "max_Plays_iTunes", "ID", "Location"]].merge(wmp_df[["Arq", "max_Plays_WMP", "Pos"]], on="Arq", how="inner")

    merged_rows = df.shape[0]
    print("\nThe merged df has",df.shape[0],"tracks")

    # SELECT ONLY RELEVANT ROWS
    df = df[df["max_Plays_iTunes"] != df["max_Plays_WMP"]]

    diff_plays = df.shape[0]
    print("\nThe merged df has",df.shape[0],"tracks where the WMP and iTunes play counts diff")

    # POPULATES LISTS
    Arq = [x for x in df["Location"]]
    ID = [x for x in df["ID"]]
    Pos = [x for x in df["Pos"]]
    iTunes_plays = [x for x in df["max_Plays_iTunes"]]
    WMP_plays = [x for x in df["max_Plays_WMP"]]
    nbr_files = len(Arq)

    print()
    wmp_cnt = 0
    iTu_cnt = 0
    # COMPARE AND CHANGE THE FILES
    for i in range(nbr_files):
        print("\nChecking file",i+1,"of",nbr_files,":",Arq[i])
        print("WMP plays:",WMP_plays[i],"// iTunes plays:",iTunes_plays[i])
        if 0 <= iTunes_plays[i] < WMP_plays[i]:
           # CHANGE ITUNES TRACK PLAY COUNT
           m = ID[i]
           track = iTunes_App.GetITObjectByID(*m) 
           print("iTunes plays:",track.PlayedCount)
           track.PlayedCount = WMP_plays[i]
           iTu_cnt = iTu_cnt+1
        else:
            # CHANGE WMP TRACK PLAY COUNT INSTEAD
            wmp_cnt = wmp_cnt+1
            track = WMP_PL.Item(Pos[i])
            plays = track.getiteminfo("UserPlayCount")
            # ENSURES THAT THE FILE IS THE SAME
            if track.getiteminfo("SourceURL").lower() == Arq[i].lower():
               print("WMP plays:",plays,"// Changing...")
               track.setItemInfo("UserPlayCount", str(iTunes_plays[i]))
               print("Doublecheck WMP count:",track.getiteminfo("UserPlayCount"))
            # track.setItemInfo("WM/Genre", "iPhone\Favorite\XXX")

    print("\n\nUpdated",wmp_cnt,"WMP plays")
    print("Updated",iTu_cnt,"iTunes plays")

    print("\nThe iTunes df has",iTu_start_rows,"tracks before deduping")
    print("The iTunes df has",iTu_end_rows,"tracks after deduping")

    print("\nThe WMP df has",wmp_start_rows,"tracks before deduping")
    print("The WMP df has",wmp_end_rows,"tracks after deduping")

    print("\nThe merged df has",merged_rows,"tracks")
    print("The merged df has",diff_plays,"tracks where the WMP and iTunes play counts diff\n")
    print()

# CHAMA PROGRAM PL_name="ALL",Fave-iPhone
Sync_plays(PL_name="XXX",Do_lib=0,rows=None)