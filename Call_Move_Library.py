# READS THE ITUNES XML FILES TO OBTAIN INFO ON MISSING TRACKS LOCATION
# REASSIGNS THE TRACKS TO MOVE LIBRARY TO ANOTHER DRIVE
# IN THIS VERSION IT USES THE XLM FILE TO IDENTIFY MISSING ID'S
# IT DOESN'T MOVE THE FILES AS IT GOES

import Read_PL
import xml.etree.ElementTree as ET
import pandas as pd
from urllib.parse import unquote
from os.path import exists
import ctypes
from ctypes import wintypes
from sys import exit

# XML COLS THAT WE WANT TO KEEP
keep_lst = ['Location','Track ID','Artist','Name']

# Define constants
FOLDERID_Music = '{4BD8D571-6D19-48D3-BE97-422220080E43}'

# Convert the string GUID to a ctypes GUID
class GUID(ctypes.Structure):
    _fields_ = [
        ('Data1', wintypes.DWORD),
        ('Data2', wintypes.WORD),
        ('Data3', wintypes.WORD),
        ('Data4', wintypes.BYTE * 8)
    ]

# READS XML LIBRARY
def parse_itunes_library_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    tracks = []
    for dict_entry in root.findall("./dict/dict/dict"):
        track = {}
        iter_elem = iter(dict_entry)
        for elem in iter_elem:
            if elem.tag == "key":
                key = elem.text
                next_elem = next(iter_elem)
                if key in keep_lst:
                   if key == "Location":
                       # Extract only the file path from the location URL
                       location = next_elem.text.split("file://localhost")[1]
                       location = location.lstrip("/")
                       track[key] = unquote(location)
                   else:
                       track[key] = next_elem.text
        tracks.append(track)
    
    return tracks

def xml_to_dataframe(xml_file):
    tracks = parse_itunes_library_xml(xml_file)
    return pd.DataFrame(tracks)

def convert_string_to_guid(guid_string):
    guid = GUID()
    ctypes.windll.ole32.CLSIDFromString(ctypes.c_wchar_p(guid_string), ctypes.byref(guid))
    return guid

def Music_folder():
    # Initialize SHGetKnownFolderPath
    SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
    SHGetKnownFolderPath.argtypes = [ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)]
    SHGetKnownFolderPath.restype = ctypes.HRESULT

    # Convert the string GUID to a ctypes GUID
    guid_music = convert_string_to_guid(FOLDERID_Music)

    # Retrieve the path
    path_ptr = ctypes.c_wchar_p()
    result = SHGetKnownFolderPath(ctypes.byref(guid_music), 0, None, ctypes.byref(path_ptr))

    music_folder = ""
    if result == 0:
        music_folder = path_ptr.value
        print("\nMusic folder path:", music_folder)
    else:
        print("\nFailed to retrieve the Music folder path. Error code:", result)
    return music_folder    

# IN CASE FIRST FUNCTION FAILS
def Music_folder_def():
    import os

    # Gets the value of the USERPROFILE environment variable
    user_profile = os.getenv('USERPROFILE')
    # Builds the path to the Music folder
    music_folder = os.path.join(user_profile, 'Music')

    print("Music folder path:", music_folder)
    return music_folder

# TO SHORTEN FUNCTION CALLS 
# Just replaces the current drive with the new one
def Cur_to_new(file,Cur_drive,Dest_drive):
    return file.replace("/", "\\").replace(Cur_drive+":\\", Dest_drive+":\\")

# RETRIEVES MUSIC FOLDER
music_folder = Music_folder()
xml_file_path = music_folder + "\\iTunes\\iTunes Music Library.xml"

# IF THE FILE DOESN'T EXIST, TRY THE OTHER FUNCTION
if not exists(xml_file_path):
   music_folder = Music_folder_def()
   xml_file_path = music_folder + "\\iTunes\\iTunes Music Library.xml"

if exists(xml_file_path):
   # THIS PART READS THE XML FILE
   print("Reading XML (this may take a while...)")
   df_xml = xml_to_dataframe(xml_file_path)
    # Reorder columns based on the list
   df_xml = df_xml.reindex(columns=keep_lst)
   df_xml['Track ID'] = pd.to_numeric(df_xml['Track ID'])
   df_xml = df_xml.sort_values(by='Track ID', ascending=True)

   # Display the first source that was read, the XML file
   print("\nThe XML df has",df_xml.shape[0],"tracks")

    # ITUNES ACTUAL LIBRARY
   col_names =  ["Art" , "Title", "ID"] 
   dict = Read_PL.Read_PL(col_names,Do_lib=True,rows=None) 
   App = dict['App']
   PLs = dict['PLs']
   df_lib = dict['DF']
   PL_name = dict["PL_Name"]

   ID = [x for x in df_lib["ID"]]
   Track_ID = [id[3] for id in ID]

   df_lib['Track ID'] = Track_ID
   df_lib = df_lib.sort_values(by='Track ID', ascending=True)

   # Join df1 and df2 on the 'Track_no' column
   df = pd.merge(df_lib, df_xml, on='Track ID', how="inner")

   print("\nThe iTunes library has",df_lib.shape[0],"tracks")

   print("\nThe merged df has",df.shape[0],"tracks")

   # LIST CREATION (list comprehension) 
   Arq = [x for x in df['Location']]
   ID = [x for x in df["ID"]]
   # Assuming 'df' is your DataFrame and 'X', 'Y', 'Z', 'W' are the column names
   match = [(x == z) and (y == w) for x, y, z, w in zip(df["Artist"], df["Name"], df["Art"], df["Title"])]
   mismatch = [1 if not x else 0 for x in match]
   mism_files = sum(mismatch)

   if mism_files==0:
      print("\nXLM and library are in sync")
   else:
       print("\nXLM and library are out of sync, can't proceed")
       exit()

    # GETS INPUT FROM USER
   Cur_drive = input("\nEnter the current drive: ")
   Cur_drive = Cur_drive.strip()
   Dest_drive = input("\nEnter the destination drive: ")
   Dest_drive = Dest_drive.strip()

   # COUNTS HOW MANY FILES CAN BE UPDATED
   Check_drive = [i for i, file in enumerate(Arq) if exists(Cur_to_new(file,Cur_drive,Dest_drive))]
   max_files = len(Check_drive)

   nbr_files = max_files
   nbr_files_inp = input(f"\nNumber of files to update (less than {max_files}) (blank for ALL): ")
   if nbr_files_inp != "":
      try:
         nbr_files = max(int(nbr_files_inp), max_files)
      except:
         nbr_files = max_files

   # PLAYLIST WITH MIGRATED FILES
   migrated_PL = "Updt_location"
   call_PL = Read_PL.Cria_PL(migrated_PL,recria="n")

    # 1st CHECK
   print("\nUpdating file location from",Cur_drive,"to",Dest_drive)
   cnt = 0
   miss = 0
   up_to_date = 0
   found = [1] * len(Arq)
   for j in range(nbr_files):
       i = Check_drive[j]
       New_loc = Cur_to_new(Arq[i],Cur_drive,Dest_drive)
       m = ID[i]
       track = App.GetITObjectByID(*m)
       if exists(New_loc) and not exists(Arq[i]) and match[i] and New_loc != Arq[i]:
          print("Updating",i+1,"of",nbr_files,":",Arq[i],"-> ",Dest_drive + ":\\")
          track.location = New_loc
          Read_PL.Add_track_to_PL(PLs,migrated_PL,track)
          cnt = cnt + 1
       elif exists(New_loc):
          up_to_date = up_to_date+1
       else:
          miss = miss+1
          found[i] = 0

   print("\n\nUpdated",cnt,"of",nbr_files,"(",miss,"not found)")
   print(up_to_date,"files already up-to-date")
   df["Found"] = found

   # TO SAVE DEAD TRACKS TO EXCEL
   print("\nSaving dead tracks to Excel...")
   if exists("D:\\iTunes\\Excel"):
      file_nm = "D:\\iTunes\\Excel\\Dead_tracks.xlsx"
   else:  
       file_nm = music_folder + "\\Dead_tracks.xlsx"

   # save the dataframe to an Excel file
   df_dead = df[ df["Found"] == 0]
   df_dead.to_excel(file_nm, index=False)
