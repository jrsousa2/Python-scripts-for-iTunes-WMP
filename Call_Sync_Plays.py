# SYNCS PLAY COUNTS BETWEEN ITUNES AND WMP
# USING THE LIBRARY XML FILE FOR ITUNES
# IF AN ITUNES PL IS PROVIDED, WILL SEARCH WMP ONLY FOR THOSE SAME FILES (NOT ENTIRE LIBRARY)
# ADDING A LOGIC TO LOCATE MISSING FILES IN PLAYLIST "DEAD"
# IF Srch_dead IS TRUE, WILL TRY TO LOCATE DEAD TRACKS BASED ON A FUZZY MATCH
# THE PURPOSE IS TO UPDATE THE PLAY COUNT FOR TRACKS THAT MATCH A DEAD TRACK
# THE LAST OPTION SHOULD NO LONGER BE NEEDED SINCE I CAN USE THE ITUNES XML FILE
# HERE THE ID THE PID (PERSISTENT ID) FOR ITUNES, SOMETHING ELSE FOR WMP

from os.path import exists
import Read_PL
# import pandas as pd
import Files

import WMP_Read_PL as WMP 

# CALLS Read_PL FUNCTION 
col_names =  ["Arq", "Plays", "ID"]

# ONLY USED TO DEBUG, NOT BEING USED ANYMORE
def Save(df,output="iTunes_vs_WMP.xlsx"):
    # CREATES A NEW COLUMN WITH THE FILE NAME WO/ THE PATH
    # RENAME THE COLUMNS HEADERS
    if "Arq" in df.columns:
        df.loc[:, "File"] = df["Arq"].apply(Files.file_w_ext)
        # df = df.rename(columns={"Arq": "Location" })

    # SAVE TO EXCEL FILE:
    file_nm = "D:\\Python\\Excel\\" + output
    # save the dataframe to an Excel file
    df.to_excel(file_nm, index=False)


# THIS MODULE SETS ARQ TO LOWER CASE FOR THE MERGER
# IT ALSO DEDUPES THE DUPLICATE RECORDS
def df_dedupe(source,df):
    # LOWERCASE THE FILE NAME
    #df["Arq"] = df["Arq"].str.lower()
    df.loc[:, "Arq"] = df["Arq"].str.lower()

    # ADDS KEY COL. TO DF
    # df["max_Plays_" + source]
    df.loc[:, "max_Plays_" + source] = df.groupby("Arq")["Plays"].transform("max")

    start_rows = df.shape[0]
    # Eliminate duplicate records based on "Arq" (subset=df["Arq"].str.lower() also works)
    df_dedupe = df.drop_duplicates(subset="Arq", keep="first")
    end_rows = df_dedupe.shape[0]
    print("\nThe",source,"df has",start_rows,"tracks before deduping (",end_rows,"after)")

    dict = {}
    dict["start_rows"] = start_rows
    dict["end_rows"] = end_rows
    dict["DF"] = df_dedupe
    return dict

# MAIN CODE
def Sync_plays(PL_name=None,PL_nbr=None,Do_lib=False,rows=None):
    # CALLS Read_PL FUNCTION ,Do_lib=True,rows=10
    if Do_lib:
       # ONLY IN THIS CASE WE NEED TO ADD THE LENGTH OF THE TRACK TO VERIFY ACCURACY
       iTu_col_names = col_names[:]
       iTu_col_names.append("PID")
       iTu_dict = Read_PL.Read_xml(iTu_col_names,rows=rows)
    else:
        iTu_dict = Read_PL.Read_PL(col_names,PL_name=PL_name,PL_nbr=PL_nbr,Do_lib=Do_lib,rows=rows,Modify_cols=False)
    
    # ASSIGNS VARS
    iTu_App = iTu_dict["App"]
    iTu_df = iTu_dict["DF"]
    # MAKES A COPY OF THE ORIGINAL PATH LIST ("ARQ")
    iTu_df["Location"] = iTu_df["Arq"].copy()
    # LOWERCASE THE FILE NAME
    iTu_df["Arq"] = iTu_df["Arq"].str.lower()

    plus_miss_rows = iTu_df.shape[0]
    
    Found = [exists(x) for x in iTu_df["Location"]]
    iTu_df["Found"] = Found

    print("\nThe iTunes df has",Found.count(False),"missing tracks")

    # SEL ONLY FOUND FILES
    iTu_df = iTu_df[iTu_df["Found"] == True]
    
    # DROPS DUPES
    print("\nDeduping iTunes df, this may take a while...")
    dict = df_dedupe("iTunes", iTu_df)
    iTu_df = dict["DF"]
    iTu_start_rows = dict["start_rows"]
    iTu_end_rows = dict["end_rows"]

    # XLSX REPORT
    # Save(iTu_df,output="Miss_files_merged2.xlsx")

    # WMP
    miss_files = []
    if Do_lib:
       print("\nReading the WMP library...")
       wmp_dict = WMP.Read_WMP_PL(col_names,Do_lib=Do_lib,rows=rows,Modify_cols=False) 
    else:
        Arq = [x for x in iTu_df["Location"] if exists(x)]
        print("\nReading the tracks from the iTunes playlist in WMP")  
        wmp_dict = WMP.Read_WMP_MC(col_names,Arq,Modify_cols=True)
        miss_files = wmp_dict["Missing"]
    
    # WMP ASSIGNING
    wmp_df = wmp_dict["DF"]
    # USED IF INPUT IS THE WHOLE LIBRARY
    WMP_lib = wmp_dict["Lib"]
    # USED IF INPUT IS A PLAYLIST (BASED ON iTunes)
    WMP_player = wmp_dict["WMP"]
    
    # CHANGE DF-ELIMINATE DUPES
    print("\nDeduping WMP df, this may take a while...")
    dict = df_dedupe("WMP",wmp_df)
    wmp_df = dict["DF"]
    wmp_start_rows = dict["start_rows"]
    wmp_end_rows = dict["end_rows"]

    wmp_df = wmp_df.rename(columns={"Pos": "WMP_Pos"})
   
    # JOIN THE DATAFRAMES [["Arq", "max_Plays_iTunes", "ID", "Location"]]
    df = iTu_df.merge(wmp_df[["Arq", "max_Plays_WMP", "WMP_Pos"]], on="Arq", how="inner")

    # XLSX REPORT
    #Save(df,output="Miss_files_merged.xlsx")

    merged_rows = df.shape[0]
    print("\nThe merged df has",df.shape[0],"tracks")

    # SELECT ONLY RELEVANT ROWS (IN CASE IT"S NOT THE DEAD TRACKS PL)
    df = df[df["max_Plays_iTunes"] != df["max_Plays_WMP"]]

    diff_plays = df.shape[0]
    print("\nThe merged df has",df.shape[0],"tracks where the WMP and iTunes play counts differ")

    # POPULATES LISTS
    Arq = [x for x in df["Location"]]
    if Do_lib:
       ID = [x for x in df["PID2"]]
    else:    
        ID = [x for x in df["ID"]]
    WMP_Pos = [x for x in df["WMP_Pos"]]
    iTunes_plays = [int(x) for x in df["max_Plays_iTunes"]]
    WMP_plays = [x for x in df["max_Plays_WMP"]]
    nbr_files = len(Arq)

    # Get the attribute name from the dictionary
    Plays_attr_name = WMP.tag_dict["Plays"]

    print()
    wmp_cnt = 0
    iTu_cnt = 0
    # TRACK METADATA
    cols = ["Art","Title","Genre"]
    # COMPARE AND CHANGE THE FILES
    for i in range(nbr_files):
        if Do_lib:
           iTu_track = iTu_App.LibraryPlaylist.Tracks.ItemByPersistentID(*ID[i])
        else:    
            iTu_track = iTu_App.GetITObjectByID(*ID[i])
        track_dict = Read_PL.iTunes_tag_dict(iTu_track, cols)
        track_meta = track_dict["Art"] +" - "+ track_dict["Title"]
        
        # CHANGE TAGS, BOTH ITUNES AND WMP CAN BE UPDATED AT THE SAME TIME NOW (NO LONGER XOR)
        # MAXIMUM OF THE 3 POSSIBLE SOURCES
        max_plays = max(iTunes_plays[i], WMP_plays[i])

        # MESSAGES
        print("\nChecking file",i+1,"of",nbr_files)
        print("Current:",track_meta)
        print("WMP plays:",WMP_plays[i],"// iTunes plays:",iTunes_plays[i])

        # UPDATE OR NOT?
        iTu_updt_cnt = iTunes_plays[i]<max_plays
        WMP_updt_cnt = WMP_plays[i]<max_plays
        if iTu_updt_cnt or WMP_updt_cnt:
           print("Updating counts:","iTunes" if iTu_updt_cnt else "","WMP" if WMP_updt_cnt else "")
        else:
           print("Not updating counts")   

        # SETS COUNTS
        # ITUNES COUNTS
        if 0 <= iTunes_plays[i] < max_plays:
           # CHANGE ITUNES TRACK PLAY COUNT
           print("Doublechecking iTunes plays:",iTu_track.PlayedCount,"// Changing...")
           iTu_track.PlayedCount = max_plays
           iTu_cnt = iTu_cnt+1
        
        # WMP COUNTS
        if 0 <= WMP_plays[i] < max_plays:
           # CHANGE WMP TRACK PLAY COUNT INSTEAD
           wmp_cnt = wmp_cnt+1
           if Do_lib:
              WMP_track = WMP_lib.Item(WMP_Pos[i])
           else:
               WMP_track = WMP_player.mediaCollection.getByAttribute("SourceURL", Arq[i]).Item(0)
           # TRIES TO RETRIEVE PLAYS
           plays = WMP_track.getiteminfo("UserPlayCount")
           # ENSURES THAT THE FILE IS THE SAME
           if WMP_track.getiteminfo("SourceURL").lower() == Arq[i].lower():
              print("Doublechecking WMP plays:",plays,"// Changing...", end=" ")
              WMP_track.setItemInfo("UserPlayCount", max_plays)
              #WMP_track.setItemInfo("UserPlayCount", str(max_plays))
              print("(doublecheck WMP count:",WMP_track.getiteminfo("UserPlayCount"),")")
    
    print("\nUpdated",wmp_cnt,"WMP plays")
    print("Updated",iTu_cnt,"iTunes plays")

    print("\nThe iTunes df has",iTu_start_rows,"tracks before deduping (",iTu_end_rows,"after)")

    print("\nThe WMP df has",wmp_start_rows,"tracks before deduping (",wmp_end_rows,"after)")

    print("\nThe merged df has",merged_rows,"tracks")
    print("\nThe merged df has",diff_plays,"tracks where the WMP and iTunes play counts differ")

    # MISSING
    if not Do_lib:
       miss_nbr = len(miss_files)
       print("\nFiles not found in WMP:",miss_nbr)
       for i in range(miss_nbr):   
           print("Missing",i+1,"of",miss_nbr,":",miss_files[i])
       print()

# CALLS PROGRAM 
Sync_plays(Do_lib=True)