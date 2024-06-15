# FUNCTIONS THAT READ AND CREATE PLAYLISTS
# iTunes API
import win32com.client
import pandas as pd
from os.path import exists
from unidecode import unidecode
#from pywin32 import datetime
#import datetime
from re import sub
from traceback import print_exc
from WMP_Read_PL import order_list

# ORDER OF THE COLS. IN THE DF (BUT THEY CAN BE SPECIFIED ANY WAY)
# THE BELOW IS JUST SO THE RIGHT HEADERS GO WITH THE RIGHT COLS.
order_list_itunes = ["PL_name","Pos","ID","Arq","Art","Title","AA","Album","Genre","Year","Group","Bitrate","Len","Covers","Plays","Skips","Added"]

# create a dictionary to store attribute names
iTu_tag_dict = {
"Arq" : "Location",
"Art" : "Artist",
"Title" : "Name",
"AA" : "AlbumArtist",
"Album" : "Album",
"Genre" : "Genre",
"Year" : "Year",
"Group" : "Grouping",
"Bitrate" : "Bitrate",
"Len" : "Time",
"Covers" : "Artwork.Count",
"Plays" : "PlayedCount",
"Skips" : "SkippedCount",
"Added" : "DateAdded"}

# THE BELOW OBJECTS ARE RECOGNIZED BY ANY FUNCTION IN THIS MODULE
def Init_iTunes():
    global iTunesApp
    global Sources
    global playlists
    global PL_ID_dict
    global PL_name_dict

    iTunesApp = win32com.client.Dispatch("iTunes.Application.1")
    Sources = iTunesApp.Sources

    for i in range(1,Sources.Count+1):
          source = Sources.Item(i)
          # ESSA VARIAVEL (playlists) DEVE SER DISPONIVEL PARA TODAS AS FUNCOES
          if source.Kind == 1:
             playlists = source.Playlists
             # THE BELOW GLOBAL DICT GIVES THE NAME OF A PL BY THE PL_ID
             PL_name_dict = {}
             PL_ID_dict = {}
             for j in range(1,playlists.Count+1):
                 playlist = playlists.Item(j)
                 PL_name = playlist.Name
                 PL_ID = playlist.playlistID
                 PL_ID_dict[PL_ID] = PL_name
                 PL_name_dict[PL_name] = j

    dict = {}
    dict['iTunesApp'] = iTunesApp
    dict['Sources'] = Sources
    dict['playlists'] = playlists
    return dict

# PLAYLISTS (NAO ESTA SENDO USADO, PODE SER DELETADO DEPOIS)
def PL_nbr_by_name(PL_name):
    dict = {}
    PL_nbr = PL_name_dict.get(PL_name,-1)
    Achou = PL_nbr != -1
    dict["res"] = Achou          
    dict["PL_nbr"] = PL_nbr
    return dict

# PLAYLISTS
def PL_name_by_ID(PL_Id):
    Achou = False
    PL_name = PL_ID_dict.get(PL_Id,"")
    return PL_name


# PLAYLISTS
def Read_PL(col_names,PL_name=None,PL_nbr=None,Do_lib=False,rows=None,Modify_cols=True,Len_type="num"):
    global playlists

    # CREATES A COPY OF THE COL. LIST SO IT'S NOT MODIFIED OUTSIDE OF THIS FUNCTION
    if not Modify_cols:
       col_names = col_names[:]

    # LISTA A SER PROCESSADA (PRIORIDADE EH DADA A LIBRARY)
    if PL_name is not None and not Do_lib:
       dict = PL_nbr_by_name(PL_name)
       if dict["res"]:
          PL_nbr = dict["PL_nbr"]
    #print("\n")
    if not Do_lib:
       if PL_nbr == None:
          # READS A PL
          print("Select from the following playlists")
          print("Number of playlists:",playlists.Count)
          for j in range(1,playlists.Count+1):
              playlist = playlists.Item(j)
              PL_name = playlist.Name
              print(j, ":",PL_name)
          user_inp = input("\nEnter comma-separated lists to process: ")
       else:
           user_inp = str(PL_nbr)
    else:
        # THIS IS THE LIBRARY
        user_inp = "0"  
        PL_nbr = None
        read_PL = iTunesApp.LibraryPlaylist
        PL_name = "library"
    
    #num = listarray(k)
    #Set playlist = playlists.Item(num)
    # data IS A LIST OF LISTS
    data = []
    PL_list = []
    # A PL SELECIONADA
    res_list = user_inp.split(",")
    nbr_PLs = len(res_list)
    for k in range(nbr_PLs):
        if not Do_lib:
           PL_nbr = res_list[k]
           read_PL = playlists.Item(PL_nbr)
           PL_name = read_PL.Name
           print("\nProcessing playlist",PL_nbr,":",PL_name)
        else: 
            print("\nProcessing music library")
        
        # BLOCO QUE SCANEIA A PL
        tracks = read_PL.Tracks

        # PROCESS SPECIFIED NUMBER OF ROWS
        if rows==None:
           numtracks = tracks.Count
        else:
            numtracks = min(rows, tracks.Count)

        # DISPLAY MESSAGE
        print("\ntracks: ",tracks.Count,"(processing",numtracks,")\n")
        
        # LOGIC TO DISPLAY IN THE LOG
        tam = max(numtracks // 20, 1)
        
        # ADD PLAYLIST NAME TO LIST
        PL_list.append(PL_name)
        # ORDER LIST SO COLUMN HEADERS ALWAYS MATCH THEIR VALUES
        col_names = order_list(col_names,order_list=order_list_itunes)
        # INICIA LISTA (m eh necessario para indicar posicao na playlist)
        # REMEMBER THAT POS IS ONLY USED TO REFERENCE THE iTUNES DB, NOT THE LISTS
        # THE RANGE FOR ITEMS IN AN ITUNES PL IS NOT 0 TO (N-1) (IT'S 1 TO N)
        for m in range(1,numtracks+1):
            track = tracks.Item(m)
            if track.Kind == 1:
               # THE SOURCE (PLAYLIST/LIBRARY)
               list = [PL_name]
               # THE TRACK POSITION
               list.append(m)
               for key in col_names:
                   if key=="Covers":
                      value = track.Artwork.Count
                   elif key=="ID":
                        value = track.GetITObjectIDs()
                   else:
                       value = getattr(track, iTu_tag_dict[key])
                   if key=="Len" and Len_type=="num":
                      value =  time_to_sec(value)   
                      # value = pd.to_timedelta(secs, unit="s") 
                   if key=="Added":
                      year = value.year
                      month = value.month
                      day = value.day
                      hour = value.hour
                      minute = value.minute
                      second = value.second
                      # CONVERTS PYWINTYPE DATE TO PANDAS DATETIME
                      value = pd.to_datetime(f"{year}-{month}-{day} {hour}:{minute}:{second}")
                   list.append(value)
               # ADD ROW TO LIST, BEFORE CREATING DF 
               data.append(list)
               if (m+1) % tam==0:
                   print("Row. no: ",m+1)
        #print("")
    # DATAFRAME
    # ADDS COL. PL IF IT WASN'T INCLUDED
    if "PL_name" not in col_names:
        col_names.append("PL_name")
    if "Pos" not in col_names:
        col_names.append("Pos") 
    # ORDER THE LIST OF COLUMNS TO MATCH THE ABOVE ORDER
    # SO COLUMN HEADERS ALWAYS MATCH THEIR VALUES
    col_names = order_list(col_names,order_list=order_list_itunes)
    df = pd.DataFrame(data, columns=col_names)

    # IF LEN IS SELECTED, TRANSFORM INTO SECONDS (DUPLICATED)
    if False and "Len" in col_names and Len_type=="num":
       df["Len"] = pd.to_timedelta(df["Len"], unit="s")
    # ORDERS DF BY ART/TITLE (CONVERTED TO UNICODE)
    df = Order(df, col_names)
    # VALUE RETURNED IS A DICT
    dict = {"App": iTunesApp, "PLs": playlists, "PL": read_PL, "PL_nbr": PL_nbr, \
            "PL_Name": PL_name, "PL_list": PL_list, "tracks": tracks, "DF": df}
    return dict

# REASSIGNS PLAYLIST
def Reassign_PL(PL_Name):
    count = 0
    result = 0
    New_PL_name = ""
    tracks = None
    for j in range(1,playlists.Count+1):
        playlistName = playlists.Item(j).Name
        if playlistName==PL_Name:
           count=count+1
           result = j
    # Do if there"s only one match
    if count==1:
       read_PL = playlists.Item(result)
       New_PL_name = read_PL.Name
       tracks = read_PL.Tracks
    else:
        print("Either the playlist doesn't exist or there is more than one of the same!")
        result = 0   
    print("\nDoublecheck playlist: Before:",PL_Name,"\ After:",New_PL_name,"\n")
    dict = {'PL_nbr': result, 'PL_Name': New_PL_name, 'tracks': tracks}
    return dict

def Cria_PL(PL_name,recria="N",Create_list=False):
    PL = iTunesApp.LibrarySource.Playlists.ItemByName(PL_name)
    # CRIA PLAYLIST SE NAO EXISTE
    if PL is None:
       iTunesApp.CreatePlaylist(PL_name)
       PL_exists = False
    else:
        PL_exists = True
        if recria.lower()=="y":
           PL.Delete() 
           iTunesApp.CreatePlaylist(PL_name)
    PL = iTunesApp.LibrarySource.Playlists.ItemByName(PL_name)
    # READS THE PL
    File_list = []
    dict = {}
    if Create_list:
       tracks = PL.Tracks
       numtracks = tracks.Count
       for m in range(1,numtracks+1):
             track = tracks.Item(m)
             if track.Kind == 1:
                #PLID = track.GetITObjectIDs()
                File = track.location
                File_list.append(File)
       dict["Files"] = set(File_list)
    dict["PL"] = PL
    dict["PL_exists"] = PL_exists
    return dict

def Add_track_to_PL(playlists,PL_name,track,File_list=None): #,PL_files
    if File_list is None or (File_list is not None and track.location not in File_list):
       try:
           PL = playlists.ItemByName(PL_name)
           PL.Addtrack(track)
       except Exception:
           print("Cant't add track to",PL_name,"playlist")

# TRACK IS A TRACK FROM SOME PLAYLIST
def Remove_track_from_PL(track): #,PL_files
    PL_Id = track.PlaylistId
    PL_name = PL_name_by_ID(PL_Id)
    try:
        #This_PL.RemoveTrack(track)
        track.delete()
        #This_PL.Tracks.Remove(track)
    except Exception as e:
        # Print the exception type and message
        print(type(e).__name__, e)
        print_exc()
        print("Can't remove track from",PL_name,"playlist") 

def Add_file_to_PL(playlists,PL_name,arq): #,PL_files
    try:
        This_PL = playlists.ItemByName(PL_name)
    except Exception:
        print("Cant't add file to",PL_name,"playlist")
        pass 
    else:
        if This_PL is not None: #and arq.lower() not in PL_files[PL]
           try:
               This_PL.AddFile(arq)
           except Exception:
               print("File not found!")   
           #PL_files[PL].add(arq.lower())

def Cria_skip_list(playlists,PL,dic):
    try:
        This_PL = playlists.ItemByName(PL)
    except Exception:
        print("Cant't read playlist",PL)
        pass 
    else:
        if This_PL is not None:
           tracks = This_PL.Tracks
           numtracks = tracks.Count
           print("tracks to skip from playlist",PL,":",numtracks,"\n")
           dic[PL] = set([])
           for m in range(1,numtracks+1):
               track = tracks.Item(m)
               if track.Kind == 1:
                  file = track.Location
                  file = file.lower()
                  dic[PL].add(file)

# SORTS BY ART AND TITLE BASED ON THE ANSI ENCODING RATHER THAN UTF-8
def Order(df, col_names):
    if "Art" in col_names:
        # Cria nova lista baseada nos nomes
        Art = [x for x in df['Art']]
        Art_sort_list = []
        Priority_list = []
        for i in range(0,len(Art)):
            art_uni = unidecode(Art[i])
            sort_vl = art_uni.lower()
            sort_vl = sort_vl.replace("-","")
            sort_vl = sort_vl.replace("'","")
            #sort_vl = sort_vl.replace("+","")
            #sort_vl = "".join(c for c in sort_vl if c.isalnum() or c in [".",","," ","+"])
            # REMOVING DUPE SPACES AGAIN */
            sort_vl = sub(' +',' ',sort_vl)
            # FIRST SORT COL.
            if len(art_uni)>0 and art_uni[0].isdigit():
               priority = 2
            else:
                priority = 1    
            # SECOND SORT COL.
            if len(sort_vl)>=1 and sort_vl[0].lower()=="*":
               sort_vl = sort_vl[1:]
            if len(sort_vl)>=2 and sort_vl[0:2].lower()=="a ":
               sort_vl = sort_vl[2:]
            if len(sort_vl)>=4 and sort_vl[0:4].lower()=="the ":
               sort_vl = sort_vl[4:]
            Art_sort_list.append(sort_vl)
            Priority_list.append(priority)
        # DO THE SAME THING FOR TITLE IF PRESENT
        if "Title" in col_names:
            Title = [x for x in df["Title"]]
            Title_sort_list = []
            for i in range(len(Title)):
                title_uni = unidecode(Title[i])
                Title_sort_list.append(title_uni)
            df["Title_sort"] = Title_sort_list    
        
        # ADDS COLS. TO DF
        df["Art_sort"] = Art_sort_list
        df["Priority"] = Priority_list
        # SORTS THE DF 
        if "Title" in col_names:
            df = df.sort_values(by=["Priority","Art_sort","Title_sort"], ascending=True)
        else:
            df = df.sort_values(by=["Priority","Art_sort"], ascending=True) 
    return df

def time_to_sec(time_str):
    # Split the time string into minutes and seconds
    min, sec = map(int, time_str.split(':'))
    # Calculate the total time in seconds
    total_sec = min * 60 + sec
    return total_sec

# READS LIBRARY TO GET ID'S (SUPPOSED TO BE FASTER)
def Read_lib_miss(rows=None):
    # THIS IS THE LIBRARY
    read_PL = iTunesApp.LibraryPlaylist
    # BLOCO QUE SCANEIA A PL
    tracks = read_PL.Tracks
    
    # data IS A LIST OF LISTS
    data = []
    print("\nProcessing iTunes music library")

    # PROCESS SPECIFIED NUMBER OF ROWS
    if rows==None:
        numtracks = tracks.Count
    else:
        numtracks = min(rows, tracks.Count)

    # DISPLAY MESSAGE
    print("\ntracks: ",tracks.Count,"(processing",numtracks,")\n")
        
    # LOGIC TO DISPLAY IN THE LOG
    tam = max(numtracks // 5, 1)
    
    # REMEMBER THAT POS IS ONLY USED TO REFERENCE THE iTUNES DB, NOT THE LISTS
    # THE RANGE FOR ITEMS IN A PL IS NOT 0 TO (N-1) (IT'S 1 TO N)
    for m in range(1,numtracks+1):
        track = tracks.Item(m)
        if m % tam==0:
           print("Row. no:",m,"of",numtracks)
        if track.Location=="" and track.Kind == 1:
           # THE SOURCE (PLAYLIST/LIBRARY), THE TRACK POSITION
           list = [track.GetITObjectIDs(), track.Artist, track.Name, track.PlayedCount, track.Time]
           # ADD ROW TO LIST, BEFORE CREATING DF 
           data.append(list)
           
    # DATAFRAME
    col_names = ["Miss_ID", "Miss_Art", "Miss_Title", "Miss_Plays", "Miss_Len"]
    df = pd.DataFrame(data, columns=col_names)

    # VALUE RETURNED IS A DICT
    dict = {"App": iTunesApp, "PLs": playlists, "tracks": tracks, "DF": df}
    return dict

# CODE TO OBTAIN ALL PROPERTIES OF TRACK AT A TIME (REQUESTED ONES IN COLS)
def iTunes_tag_dict(track,cols,Len_type="num"):
    dict = {}
    # PROPERTIES
    for key in cols:
        if key=="Covers":
           value = track.Artwork.Count
        elif key=="ID":
           value = track.GetITObjectIDs()
        else:
           value = getattr(track, iTu_tag_dict[key])
        if key=="Len" and Len_type=="num":
           value =  time_to_sec(value)   
        if key=="Added":
           year = value.year
           month = value.month
           day = value.day
           hour = value.hour
           minute = value.minute
           second = value.second
           # CONVERTS PYWINTYPE DATE TO PANDAS DATETIME
           value = pd.to_datetime(f"{year}-{month}-{day} {hour}:{minute}:{second}")
        dict[key] = value
        #list.append(value)
    return dict



# INITIALIZES iTunes
Init_iTunes()

