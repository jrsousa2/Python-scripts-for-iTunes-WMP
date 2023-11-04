import win32com.client
import pandas as pd
import Read_PL as iTunes

# ORDER OF THE COLS. IN THE DF (BUT THEY CAN BE SPECIFIED ANY WAY)
# THE BELOW IS JUST SO THE RIGHT HEADERS GO WITH THE RIGHT COLS.
order_list_wmp = ["PL_nbr","PL_name","Pos","ID","Arq","Art","Title","AA","Album","Genre","Year","Group","Bitrate","Len","Plays","Skips","Added"]

tag_dict = {
    "Art" : "Artist",
    "AA" : "WM/AlbumArtist",
    "Album" : "WM/AlbumTitle", # OR Album
    "Genre" : "WM/Genre", # OR Genre
    "Year" : "WM/Year",
    "Group" : "WM/ContentGroupDescription",
    "Bitrate" : "Bitrate",
    "Plays": "UserPlayCount",
    "Added": "AcquisitionTime" #AcquisitionTimeYearMonthDay
    }

# OS OBJETOS ABAIXO SAO RECONHECIDOS POR QQ FUNCAO DESSE MODULO
# OBJECT wmp IS THE Player
def Init_wmp():
    global wmp
    global library
    global playlists

    wmp = win32com.client.Dispatch('WMPlayer.OCX')
    library = wmp.mediaCollection.getAll()

    # get the playlist collection
    playlists = wmp.playlistCollection.getAll()


# create a dictionary to store attribute names
# THERE'S A CODE TO OBTAIN ALL PROPERTIES OF TRACK (COMMENTED OUT)
def tag_dict_wmp(item,cols):
    dict = {}
    # for i in range(item.attributeCount):
    #    k = item.getAttributeName(i)
    #    print("Attrib:",k,"Value:",item.getItemInfo(k))
    # PROPERTIES
    for key in cols:
        if key in tag_dict.keys():
           dict[key] = item.getItemInfo(tag_dict[key])
    if "Plays" in cols: 
        dict["Plays"] = int(dict["Plays"]) 
    if "Skips" in cols: 
        dict["Skips"] = 0
    if "Bitrate" in cols: 
        dict["Bitrate"] = int(dict["Bitrate"])/1000   
    if "Added" in cols: 
        dict["Added"] = pd.to_datetime(dict["Added"], format="%d/%b/%Y %I:%M:%S %p")    
    if "Arq" in cols: 
        dict["Arq"] = item.sourceURL
    if "Title" in cols: 
        dict["Title"] = item.name
    if "Len" in cols: 
        dict["Len"] = item.durationString
    if "ID" in cols: 
        dict["ID"] = 0    
    return dict

# FINDS A PLAYLIST
def Get_WMP_PL_by_nbr(srch_PL):
    dict = {}
    Achou = False
    PL_nbr = 0
    for PL in playlists:
        PL_name = PL.name
        if PL_name == srch_PL:
           Achou = True
           break
        PL_nbr = PL_nbr+1   
    dict["res"] = Achou          
    dict["PL_no"] = PL_nbr
    return dict    

# PLAYLISTS
def Read_WMP_PL(col_names,PL_name=None,PL_no=None,Do_lib=False,rows=None,Modify_cols=True):

    # CREATES A COPY OF THE COL. LIST SO IT'S NOT MODIFIED OUTSIDE OF THIS FUNCTION
    if not Modify_cols:
       col_names = col_names[:]
    
    
    Init_wmp()
    # LISTA A SER PROCESSADA
    By_name = False
    if not Do_lib and PL_name is not None:
       #print("\nReading WMP playlists...this may take a while")
       read_PL = wmp.playlistCollection.getByName(PL_name).Item(0)
       PL_no = 0
       user_inp = "0"
       By_name = True
       # THIS PART USED TO TAKE TOO LONG
       #dict = Get_WMP_PL_by_nbr(PL_name)
       #if dict["res"]:
          #PL_no = dict["PL_no"]
    elif not Do_lib and PL_no == None:
         nbr_PLs = playlists.Count
         # READS A PL
         print("\nReading WMP playlists...this may take a while")
         print("Select from the following WMP playlists")
         print("Number of playlists:",nbr_PLs)
         for j in range(nbr_PLs):
             playlist = playlists[j]
             # PL_name = playlist.Name
             print(j, ":", PL_name)
         user_inp = input("\nEnter comma-separated lists to process: ")
    elif not Do_lib and PL_no != None:
         user_inp = str(PL_no)
    elif Do_lib:
         # THIS IS THE LIBRARY
         user_inp = "0"
         read_PL = library
         PL_name = "library"
         PL_nbr = 0
    
    # data IS A LIST OF LISTS
    data = []
    # A PL SELECIONADA
    res_list = user_inp.split(",")
    nbr_PLs = len(res_list)
    for k in range(nbr_PLs):
        if not Do_lib and By_name:
           print("\nProcessing playlist",k+1,"of",nbr_PLs,":",PL_name)
           PL_nbr = 0
        elif not Do_lib:
             # PL_name = playlists[k].name
             PL_nbr = int(res_list[k])
             read_PL = playlists[PL_nbr]
             PL_name = read_PL.Name
             print("\nProcessing playlist",k+1,"of",nbr_PLs,":",PL_name)
        elif Do_lib: 
             print("\nProcessing music library")
        
        # PROCESS SPECIFIED NUMBER OF ROWS
        if rows is None:
           if not Do_lib:
              numtracks = read_PL.count
           else:
              numtracks = len(library)   
        else:
            numtracks = min(rows, read_PL.count if not Do_lib else len(library))

        # DISPLAY MESSAGE
        print("\ntracks: ",read_PL.Count,"(processing",numtracks,")\n")

        # LOGIC TO DISPLAY IN THE LOG
        tam = max(numtracks // 20, 1)
        
        # ORDER LIST SO COLUMN HEADERS ALWAYS MATCH THEIR VALUES
        col_names = iTunes.order_list(col_names,order_list=order_list_wmp)
        # THE RANGE FOR ITEMS IN A WMP PL IS NOT 0 TO (N-1) (IT'S 1 TO N)
        for m in range(0,numtracks):
            if not Do_lib:
               track = read_PL.Item(m)
            else:
               track = library[m]    
            
            # ONLY DOES AUDIO
            if track.getiteminfo("MediaType")=="audio":
               # THE SOURCE (PLAYLIST/LIBRARY)
               tag_list = [PL_nbr,PL_name]
               # THE TRACK POSITION
               tag_list.append(m)
               dict = tag_dict_wmp(track,col_names)
               for key in col_names:
                   value = dict[key]
                   tag_list.append(value)
               #ADD ROW TO LIST, BEFORE CREATING DF
               data.append(tag_list)
               if (m+1) % tam==0:
                   print("Row. no: ",m+1)
        #print("")
    # DATAFRAME
    # ADDS COL. PL IF IT WASN'T INCLUDED
    if "PL_nbr" not in col_names:
        col_names.append("PL_nbr") 
    if "PL_name" not in col_names:
        col_names.append("PL_name")
    if "Pos" not in col_names:
        col_names.append("Pos")    
    # ORDER THE LIST SO COLUMN HEADERS MATCH THEIR VALUES
    col_names = iTunes.order_list(col_names,order_list=order_list_wmp)
    df = pd.DataFrame(data, columns=col_names)
    # SETS YEAR TYPE TO INTEGER
    if "Year" in col_names:
        df['Year'] = pd.to_numeric(df['Year'], errors="coerce")
        df['Year'] = df['Year'].fillna(0)
    
    # ORDERS DF BY ART/TITLE (CONVERTED TO UNICODE)
    # df = Order(df, col_names)
    # VALUE RETURNED IS A DICT
    dict = {"Player": wmp, "Lib": library, "media": wmp.mediaCollection, "PLs": playlists, "PL": read_PL, "PL_no": PL_nbr, \
            "PL_Name": PL_name, "tracks": 1, "DF": df}
    return dict


# add some tracks to the playlist
# mc = wmp.mediaCollection
# filename = "D:\\MP3\Favorites\\The Maxx - Cocaine.mp3"
# track = library.getByAttribute('MediaPath', filename)
# track = library.getByPath("D:\\MP3\Favorites\\The Maxx - Cocaine.mp3")
# print("Test:",track.sourceURL)