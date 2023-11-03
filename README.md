# Python-scripts-for-iTunes
I just started on GitHub, eventually I plan to share my Python scripts to manage your iTunes music library.

11/03/2023 Edit: I've just shared three of the python codes that I use to manage my music collection (with 
roughly 63,000 mp3 files) in both Tunes and Windows Media Player.

The main code Call_Save_to_Excel.py can be run from VS Code (or any other suitable python compiler/editor) and will
create an Excel file with a list of your mp3 files and chosen tags, either through an interface with iTunes or Windows Media Player.
You can choose to export a list of files (plus tags) either for all of your library or for a playlist (or set of playlists).

This is a basic code and is useful to give first time users some insights into the libraries of functions, methods and
properties of both WMP and iTunes. For example:
  How can you directly reference a track in the iTunes library with a tuple of 4 ID numbers?
  How do you read the file tags that WMP keeps in the metadata, some of which are intrinsic to the file and some of which are extrinsic?
  How can you sync play counts between iTunes and WMP?

Here's a brief description of the parameters:

# MAIN CODE Call_Save_to_Excel.py calls Read_PL.py (module for iTunes) and WMP_Read_PL.py (module for WMP)

def Save_Excel(PL_name=None,PL_nbr=None,Do_lib=False,rows=None,iTunes=True):
**iTunes:** If True, will use iTunes library or playlists (if False, it will use WMP)
**PL_name**: if supplied and Do_lib=False, it will create a list of files and tags for the playlist PL_name
**PL_nbr**: If supplied and Do_lib=False, it will create a list of files and tags for the playlist number PL_nbr
**rows**: Whether the source is the whole library or a playlist, this option limits the number of files to read,
if necessary.

If none of the above parameters (PL_name, PL_Nbr) is supplied, and Do_lib=False, the code will display a list of
playlists along with their respective numbers (per the iTunes or WMP libraries), and the user will be able to enter 
all playlists that they wish to process separated by comma (the program has instruction prompts and other messages to guide the user).

**col_names**: This variable is entered straight into the code and identifies the tags that the user
wants to extract from the mp3 files. (The code can be easily modified to have this variable as a parameter to call
the main macro with.) Here's an example:
col_names =  ["Arq","Art","Title","Year"]

Here's a list of all tags that can be extracted:
Tag	Description	iTunes	WMP
ID	4-dimension tuple that id's a file in the iTunes library or playlist	x	
Arq	the name of the file	x	x
Art	the artist	x	x
Title	the song title	x	x
AA	the album artist	x	x
Album	the album name	x	x
Genre	the genre	x	x
Year	the release year	x	x
Group	the Grouping tag	x	
Bitrate	the file bit rate	x	x
Len	the length of the track	x	x
Covers	the number of covers/art work the file has	x	
Plays	the number of plays	x	x
Skips	the number of skips	x	
Added	the datetime the file was added to the library	x	x
![image](https://github.com/jrsousa2/Python-scripts-for-iTunes/assets/94881602/e05ba46c-01f6-4e1a-97bd-d41e3136132f)


**Be sure to change the default folder** that the file will be saved to in the main code.
This is done in lines 41 and 43 of the main codes and is set to:
file_nm = "D:\\iTunes\\Excel\\" + user_inp + ".xlsx"
