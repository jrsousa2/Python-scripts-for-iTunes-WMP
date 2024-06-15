# READS THE ITUNES XML FILES TO OBTAIN INFO ON MISSING TRACKS LOCATION
# REASSIGNS THE TRACKS TO MOVE LIBRARY TO ANOTHER DRIVE

import Read_PL
import xml.etree.ElementTree as ET
import pandas as pd
from urllib.parse import unquote
from os.path import exists

# XML COLS THAT WE WANT TO KEEP
keep_lst = ['Location','Track ID','Artist','Name']

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

# Path to your iTunes library XML file
print("Reading XML...")
xml_file_path = "E:\\Backup\\iTunes Music Library.xml"
df_xml = xml_to_dataframe(xml_file_path)

# Reorder columns based on the list
df_xml = df_xml.reindex(columns=keep_lst)
df_xml['Track ID'] = pd.to_numeric(df_xml['Track ID'])
df_xml = df_xml.sort_values(by='Track ID', ascending=True)

# ITUNES ACTUAL LIBRARY
col_names =  ["Art" , "Title", "ID"] 
dict = Read_PL.Read_PL(col_names,Do_lib=True,rows=None) 
App = dict['App']
# PLs = dict['PLs']
df_lib = dict['DF']

ID = [x for x in df_lib["ID"]]
Track_ID = [id[3]-1 for id in ID]

df_lib['Track ID'] = Track_ID
df_lib = df_lib.sort_values(by='Track ID', ascending=True)

# Join df1 and df2 on the 'Track_no' column
df = pd.merge(df_lib, df_xml, on='Track ID', how='inner')

print("\nThe XML df has",df_xml.shape[0],"tracks")

print("\nThe iTunes library has",df_lib.shape[0],"tracks")

print("\nThe merged df has",df.shape[0],"tracks")

# LIST CREATION (list comprehension) 
Arq = [x for x in df['Location']]
ID = [x for x in df["ID"]]
# Assuming 'df' is your DataFrame and 'X', 'Y', 'Z', 'W' are the column names
match = [(x == z) and (y == w) for x, y, z, w in zip(df["Artist"], df["Name"], df["Art"], df["Title"])]
mismatch = [1 if not x else 0 for x in match]
nbr_files = len(Arq)

# 1st CHECK
print("\nUpdating file location from D: to E:")
print("Misaligned files:",sum(mismatch),"\n")
cnt = 0
miss = 0
up_to_date = 0
found = []
for i in range(nbr_files):
    New_loc = Arq[i].replace("/", "\\")
    New_loc = New_loc.replace("D:\\", "E:\\")
    m = ID[i]
    track = App.GetITObjectByID(*m)
    if exists(New_loc) and match[i] and New_loc != Arq[i]:
       found.append(1)
       print("Updating",i+1,"of",nbr_files,":",Arq[i],"-> E:\\")
       track.location = New_loc
       cnt = cnt + 1
    elif exists(New_loc):
         up_to_date = up_to_date+1
         found.append(1)
    else:
        miss = miss+1
        found.append(0)

print("Updated",cnt,"of",nbr_files,"(",miss,"not found)")
print(up_to_date,"files already up-to-date")
df["Found"] = found

print("Saving dead tracks to Excel...")
file_nm = "D:\\iTunes\\Excel\\Dead_tracks.xlsx"
# save the dataframe to an Excel file
df_dead = df[ df["Found"] == 0]
df_dead.to_excel(file_nm, index=False)
