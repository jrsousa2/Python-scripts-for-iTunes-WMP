# SYNCS PLAY COUNTS BETWEEN ITUNES AND WMP
# IF AN ITUNES PL IS PROVIDED, WILL SEARCH WMP ONLY FOR THOSE SAME FILES (NOT ENTIRE LIBRARY)
# ADDING A LOGIC TO LOCATE MISSING FILES IN PLAYLIST "DEAD"
# IF Srch_dead IS TRUE, WILL TRY TO LOCATE DEAD TRACKS BASED ON A FUZZY MATCH
# THE PURPOSE IS TO UPDATE THE PLAY COUNT FOR TRACKS THAT MATCH A DEAD TRACK
# THE LAST OPTION SHOULD NO LONGER BE NEEDED SINCE I CAN USE THE ITUNES XML FILE

from os.path import exists
from unidecode import unidecode
from Tags import similar_ratio
from Tags import Genre_is_live
import Read_PL
import WMP_Read_PL as WMP
import pandas as pd
import Files


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
    file_nm = "D:\\iTunes\\Excel\\" + output
    # save the dataframe to an Excel file
    df.to_excel(file_nm, index=False)

# SEARCHES FOR FILES IN THE D:\MP3 FOLDER USING ART+TITLE
# RATIO IS THE MATCH RATIO
def find_files(file_list, substring):
    match_files = [file for file in file_list if substring in file]
    return match_files

# GIVEN A DF, SEARCHES MISSING FILES IN THE MP3 DIRECTORY
# THE OUTPUT IS iTUNES PL TO BE READ
def Search_missing(App, PLs, df, thres):
    Art = [x for x in df["Miss_Art"]]
    Title = [x for x in df["Miss_Title"]]
    Miss_ID = [x for x in df["Miss_ID"]]
    nbr_files = len(Art)

    # LISTA DE GREATES HITS
    dir_path = "D:\\MP3"
    print("\nBuilding list of mp3 files in",dir_path,"(this may take a while...)")
    filelist = Files.get_Win_files(dir_path, ".mp3")

    # CREATES A PLAYLIST WITH THE RESULTS
    PL_nm = "Find_dead"
    Tag_PL = Read_PL.Cria_PL(PL_nm,recria="Y")

    res = {}
    # LOCATION IS USED SINCE "ARQ" IS LOWER CASE
    res["Arq"] = []
    res["Ratio"] = []
    res["Miss_Art"] = []
    res["Miss_Title"] = []
    print()
    for i in range(nbr_files):
        # MISS TRACK METADATA
        miss_track = App.GetITObjectByID(*Miss_ID[i])
        try:
            max_ratio_tag = float(miss_track.comment)
        except ValueError:
            max_ratio_tag = 0
        if max_ratio_tag > thres or max_ratio_tag == 0:    
            # CHANGE ITUNES TRACK PLAY COUNT
            # New_loc = Move(Arq[i],Art[i],Genre[i])
            file_to_srch = unidecode(Art[i].lower()+" - "+Title[i].lower())
            # Split the string by spaces and count the number of words
            print("\nSearching",i+1,"of ",nbr_files,":",file_to_srch)
            # Create a set with elements longer than 1 character
            file_to_srch_set = set(word for word in file_to_srch.split(" ") if len(word)>1)
            word_count = len(file_to_srch_set)
            # Create a subset where some word in srch_str matches some word in the 2nd element
            sublist = [tup for tup in filelist if sum(word in file_to_srch_set for word in set(tup[1].split(" ")))>=word_count/2]
            sublist_len = len(sublist)
            print("List len:",sublist_len)
            found_file = "@"
            cnt = 0
            max_ratio = 0
            # MATCH DOESN"T NEED TO BE EXACT
            while (found_file):
                found_file = False
                for file, normal_file in sublist:
                    dict = similar_ratio(file_to_srch,normal_file,thres = thres)
                    ratio = dict["ratio"]
                    match = dict["match"]
                    max_ratio = max(max_ratio, ratio)
                    if match and file not in res["Arq"]:
                        cnt = cnt + 1
                        found_file = file
                        print("\tFound:",found_file,"-- Ratio",ratio)
                        res["Arq"].append(found_file)
                        res["Ratio"].append(ratio)
                        res["Miss_Art"].append(Art[i])
                        res["Miss_Title"].append(Title[i])
                        break
            
            # UPDT MAX_RATIO TAG
            miss_track.comment = max_ratio

    # Convert the list to a DataFrame
    df_res = pd.DataFrame(res)

    # Sort the DataFrame by columns A, B, and X in descending order
    
    df_res = df_res.sort_values(by=['Miss_Art', 'Miss_Title', 'Ratio'], ascending=[True, True, False])

    # MANAGE DUPES
    start_rows = df_res.shape[0]
    # Drop duplicates based on columns A and B, keeping the first occurrence with the greatest X
    df_res = df_res.drop_duplicates(subset=['Miss_Art', 'Miss_Title'], keep='first')
    end_rows = df_res.shape[0]
    print("\nThe missing match df has",start_rows,"tracks before deduping (",end_rows,"after)")

    # CREATE RESULTING PL ONLY WITH DEDUPED FILES
    Arq = [x for x in df_res["Arq"]]
    # ADDS FOUND FILES TO PL
    for i in range(len(Arq)):
        Read_PL.Add_file_to_PL(PLs,PL_nm,Arq[i]) 

    # MAKES ARQ LOWER CASE
    if len(Arq)>0:
       df_res["Arq"] = df_res["Arq"].str.lower()

    # RETURNS A DF
    return df_res
    

# THIS MODULE SETS ARQ TO LOWER CASE FOR THE MERGER
# IT ALSO DEDUPES THE DUPLICATE RECORDS
def df_dedupe(source,df):
    # LOWERCASE THE FILE NAME
    df["Arq"] = df["Arq"].str.lower()

    # ADDS KEY COL. TO DF
    df["max_Plays_" + source] = df.groupby("Arq")["Plays"].transform("max")

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

# DIFFERENCE BT 2 LENTGHS
def Diff_secs(t1,t2):
    diff = Read_PL.time_to_sec(t1)-Read_PL.time_to_sec(t2)
    return abs(diff)

# MAIN CODE
# Srch_dead IS USED FOR DEAD TRACKS (NEEDS TO SCAN WHOLE LIBRARY TO BE ABLE TO DELETE TRACKS)
def Sync_plays(PL_name=None,PL_nbr=None,Do_lib=False,rows=None,Srch_dead=0,thres=0.75):
    # CALLS Read_PL FUNCTION ,Do_lib=True,rows=10

    if Srch_dead:
       Do_lib = True
       # CALLS Read_PL FUNCTION 
       iTu_lib_dict = Read_PL.Read_lib_miss(rows=rows)
       iTu_App = iTu_lib_dict["App"]
       iTu_lib_df = iTu_lib_dict["DF"]
       PLs = iTu_lib_dict["PLs"]
       print("\nNumber of missing found:",iTu_lib_df.shape[0])

       # IF TO RUN ON PL "DEAD", WILL SEARCH ACTUAL LOCATION OF THE MISSING FILES FIRST
       miss_df = Search_missing(iTu_App, PLs, iTu_lib_df, thres)
       # ONLY IN THIS CASE WE NEED TO ADD THE LENGTH OF THE TRACK TO VERIFY ACCURACY
       iTu_col_names = col_names[:]
       iTu_col_names.append("Len")
       iTu_dict = Read_PL.Read_PL(iTu_col_names,PL_name="Find_missing",rows=None,Modify_cols=False,Len_type="char")
    else:
        iTu_dict = Read_PL.Read_PL(col_names,PL_name=PL_name,PL_nbr=PL_nbr,Do_lib=Do_lib,rows=rows,Modify_cols=False)
    
    # ASSIGNS VARS
    iTu_App = iTu_dict["App"]
    # PLs = iTu_dict["PLs"]
    iTu_df = iTu_dict["DF"]
    # MAKES A COPY OF THE ORIGINAL PATH LIST ("ARQ")
    iTu_df["Location"] = iTu_df["Arq"].copy()
    # LOWERCASE THE FILE NAME
    iTu_df["Arq"] = iTu_df["Arq"].str.lower()
    
    
    # LEFT JOINS RESULTS WITH DEAD TRACKS (IT SHOULD NOT MISS ANY RECORDS)
    if Srch_dead and iTu_df.shape[0]>0:
       # Merge 3 DataFrames
       iTu_df = iTu_df.merge(miss_df, on="Arq", how="left").merge(iTu_lib_df, on=["Miss_Art", "Miss_Title"], how="left")

    # DROPS DUPES
    dict = df_dedupe("iTunes", iTu_df)
    iTu_df = dict["DF"]
    iTu_start_rows = dict["start_rows"]
    iTu_end_rows = dict["end_rows"]

    # XLSX REPORT
    # Save(iTu_df,output="Miss_files_merged2.xlsx")

    # WMP
    miss_files = []
    if Do_lib:
       print("\nReading the WMP playlist")
       wmp_dict = WMP.Read_WMP_PL(col_names,PL_name=PL_name,PL_nbr=PL_nbr,Do_lib=Do_lib,rows=rows,Modify_cols=False) 
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
    dict = df_dedupe("WMP",wmp_df)
    wmp_df = dict["DF"]
    wmp_start_rows = dict["start_rows"]
    wmp_end_rows = dict["end_rows"]

    wmp_df = wmp_df.rename(columns={"Pos": "WMP_Pos"})
   
    # JOIN THE DATAFRAMES [["Arq", "max_Plays_iTunes", "ID", "Location"]]
    df = iTu_df.merge(wmp_df[["Arq", "max_Plays_WMP", "WMP_Pos"]], on="Arq", how="inner")

    # DROP SOME COLS. (FILE IS CREATED ONLY WHEN SAVING THE DF TO XLSX)
    # df = df.drop(columns=["Plays","Art_sort","Title_sort","Priority", "Pos", "Plays"])

    # XLSX REPORT
    # Save(df,output="Miss_files_merged.xlsx")

    # LEFT JOINS RESULTS IF DEAD TRACKS (IT SHOULD NOT MISS ANY RECORDS)
    if Srch_dead and df.shape[0]>0:
       Miss_ID = [x for x in df["Miss_ID"]] 
       Miss_plays = [x for x in df["Miss_Plays"]]
       Miss_Len = [x for x in df["Miss_Len"]]
       Len = [x for x in df["Len"]]
       Ratio = [x for x in df["Ratio"]]
       #Miss_Art = [x for x in df["Miss_Art"]]
       #Miss_Title = [x for x in df["Miss_Title"]]

    merged_rows = df.shape[0]
    print("\nThe merged df has",df.shape[0],"tracks")

    # SELECT ONLY RELEVANT ROWS (IN CASE IT"S NOT THE DEAD TRACKS PL)
    if not Srch_dead:
       df = df[df["max_Plays_iTunes"] != df["max_Plays_WMP"]]

    diff_plays = df.shape[0]
    if not Srch_dead:
       print("\nThe merged df has",df.shape[0],"tracks where the WMP and iTunes play counts differ")

    # POPULATES LISTS
    Arq = [x for x in df["Location"]]
    ID = [x for x in df["ID"]]
    WMP_Pos = [x for x in df["WMP_Pos"]]
    iTunes_plays = [x for x in df["max_Plays_iTunes"]]
    WMP_plays = [x for x in df["max_Plays_WMP"]]
    nbr_files = len(Arq)

    print()
    wmp_cnt = 0
    iTu_cnt = 0
    del_cnt = 0
    # COMPARE AND CHANGE THE FILES
    for i in range(nbr_files):
        # TRACK METADATA
        cols = ["Art","Title","Genre"]
        iTu_track = iTu_App.GetITObjectByID(*ID[i])
        track_dict = Read_PL.iTunes_tag_dict(iTu_track, cols)
        track_meta = track_dict["Art"] +" - "+ track_dict["Title"]
        # CHANGE TAGS, BOTH ITUNES AND WMP CAN BE UPDATED AT THE SAME TIME NOW (NO LONGER XOR)
        # STARTS
        if Srch_dead:
           # MISS TRACK METADATA
           miss_track = iTu_App.GetITObjectByID(*Miss_ID[i])
           miss_dict = Read_PL.iTunes_tag_dict(miss_track,cols)
           miss_meta = miss_dict["Art"] +" - "+ miss_dict["Title"]
           # CHECKS
           Live_match = Genre_is_live(track_dict["Genre"])==Genre_is_live(miss_dict["Genre"])
           # CHECK SIMILARITY AGAIN (BASED ON TAGS INSTEAD OF FILE NAME)
           dict = similar_ratio(track_meta,miss_meta,thres = thres)
           ratio = dict["ratio"]
           match = dict["match"]
           prt1 = "(match ratio "+ str(Ratio[i]) + ")"
           if Ratio[i] != ratio:
              prt1 = prt1 + "(Updated from "+ str(Ratio[i])+")"
              Ratio[i] = ratio
           prt2 = "(miss count: "+ str(Miss_plays[i]) + ")"
        else:
           Updt = True   
           prt1 = ""
           prt2 = ""
        
        # MAXIMUM OF THE 3 POSSIBLE SOURCES
        max_plays = max(iTunes_plays[i], WMP_plays[i], Miss_plays[i] if Srch_dead else 0)

        # MESSAGES
        print("\nChecking file",i+1,"of",nbr_files,prt1)
        print("Current:",track_meta,"--Length: "+Len[i] if Srch_dead else "")
        if Srch_dead:
           print("Missing:",miss_meta,"--Length:",Miss_Len[i])
        print("WMP plays:",WMP_plays[i],"// iTunes plays:",iTunes_plays[i],prt2)

        # UPDATE OR NOT?
        if Srch_dead:
           if not (Live_match and match):
              print("Not updating...")
              user_inp = "N"
           elif Len[i] == Miss_Len[i]:
               print("Lengths match, updating...")
               user_inp = "Y"
           elif Ratio[i]>0.849:
               print("\nMatch ratio is high, updating...")
               user_inp = "Y"   
           else:   
               user_inp = input("\nPress any key to update/delete (N to not): ")
           # WILL UPDATE?    
           Updt = (user_inp.upper() != "N")

        iTu_updt_cnt = Updt and iTunes_plays[i]<max_plays
        WMP_updt_cnt = Updt and WMP_plays[i]<max_plays
        if iTu_updt_cnt or WMP_updt_cnt:
           print("Updating counts:","iTunes" if iTu_updt_cnt else "","WMP" if WMP_updt_cnt else "")
        else:
           print("Not updating counts")   

        # SETS COUNTS
        # ITUNES COUNTS
        if 0 <= iTunes_plays[i] < max_plays and Updt:
           # CHANGE ITUNES TRACK PLAY COUNT
           print("Doublechecking iTunes plays:",iTu_track.PlayedCount,"// Changing...")
           iTu_track.PlayedCount = max_plays
           iTu_cnt = iTu_cnt+1
        
        # WMP COUNTS
        if 0 <= WMP_plays[i] < max_plays and Updt:
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
              WMP_track.setItemInfo("UserPlayCount", str(max_plays))
              print("(doublecheck WMP count:",WMP_track.getiteminfo("UserPlayCount"),")")

        # REMOVE DEAD TRACKS AFTER COUNT UPDT
        if Srch_dead:
           if Updt:
              print("\nDeleting dead track")
              try:
                 m = Miss_ID[i]
                 miss_track = iTu_App.GetITObjectByID(*m)
                 Read_PL.Remove_track_from_PL(miss_track)
              except:
                 print("Missing track has been deleted already")
              else:    
                 print("Dead track has been deleted!")
                 del_cnt = del_cnt+1
           else:
               print("Not deleting...")
    
    
    print("\nUpdated",wmp_cnt,"WMP plays")
    print("Updated",iTu_cnt,"iTunes plays")
    print("Deleted",del_cnt,"dead iTunes tracks")

    print("\nThe iTunes df has",iTu_start_rows,"tracks before deduping (",iTu_end_rows,"after)")

    print("\nThe WMP df has",wmp_start_rows,"tracks before deduping (",wmp_end_rows,"after)")

    print("\nThe merged df has",merged_rows,"tracks")
    if not Srch_dead:
       print("\nThe merged df has",diff_plays,"tracks where the WMP and iTunes play counts differ")

    # MISSING
    miss_nbr = len(miss_files)
    print("\nFiles not found in WMP:",miss_nbr)
    for i in range(miss_nbr):   
        print("Missing",i+1,"of",miss_nbr,":",miss_files[i])
    print()

# CALLS PROGRAM 
Sync_plays(PL_name="AAA", Do_lib=1, Srch_dead=0, thres=0.85, rows=None)