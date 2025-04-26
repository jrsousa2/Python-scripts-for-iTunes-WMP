I just started on GitHub, eventually I plan to share my Python scripts to help you manage your iTunes music library.

<br>11/03/2023 Edit: I've just shared some of the various python codes that I use to manage my music collection (with 
roughly 64,000 mp3 files) in both <b>iTunes and Windows Media Player</b>.

<br>The main code Call_Save_to_Excel.py can be run from VS Code (or any other suitable python compiler/editor) and will
create an Excel file with a list of your mp3 files and chosen tags, either through an interface with iTunes or Windows Media Player.
You can choose to export a list of files (plus tags) either for all of your library or for a playlist (or set of playlists).

<br>These are relatively basic codes and can be useful to give first time users some insights into the libraries of functions, methods and
properties of both the WMP and iTunes COM API's. (Asking ChatGPT to create this type of code doesn't work, it gets lots of these API questions wrong -- ChatGPT struggles with tricky questions.)

<br>For example:
  <br>&nbsp;&nbsp;&nbsp;How can you directly reference a track in the iTunes library with a tuple of 4 ID numbers?
  <br>&nbsp;&nbsp;&nbsp;Even better, how can you directly reference a track in the iTunes library with a persistent ID (that sticks regardless of session)?
  <br>&nbsp;&nbsp;&nbsp;How can you locate and load the iTunes XML library (the fastest way to read all the database of music files)?
  <br>&nbsp;&nbsp;&nbsp;How do you read the file tags that WMP keeps in the metadata, some of which are embedded into the file and some of which are external?
  <br>&nbsp;&nbsp;&nbsp;How can you sync play counts between iTunes and WMP? (Code Call_Sync_Plays.py does that by taking the max of both).

<br>Here's a brief description of the parameters of the codes:

<br><b>Code Call_Save_to_Excel.py calls Read_PL.py (module for iTunes) and WMP_Read_PL.py (module for WMP)</b>

def Save_Excel(PL_name=None,PL_nbr=None,Do_lib=False,rows=None,iTunes=True):
<br>**iTunes:** If True, will use iTunes library or playlists (if False, it will use WMP)
<br>**Do_lib**: If True, will run on the whole library instead of a playlist (or list of playlists)
<br>**PL_name**: If supplied and Do_lib=False, will create a list of files and tags for the playlist PL_name
<br>**PL_nbr**: If supplied and Do_lib=False, will create a list of files and tags for the playlist number PL_nbr
<br>**rows**: Whether the source is the whole library or a playlist, this option limits the number of files to read,
if necessary.
  
If none of the above parameters (PL_name, PL_Nbr) is supplied, and Do_lib=False, the code will display a list of
playlists along with their respective numbers (per the iTunes or WMP libraries), and the user will be able to enter 
all playlists that they wish to process separated by comma (the program has instruction prompts and other messages to guide the user).

<br>**col_names**: This variable is entered straight into the code and identifies the tags that the user
wants to extract from the mp3 files. 
<br>(The code can be easily modified to have this variable as a parameter to call the main macro with.) 
<br>Here's an example:
<br>col_names =  ["Arq","Art","Title","Year"]

<br>Here's a list of all the tags that can be extracted:
![image](https://github.com/jrsousa2/Python-scripts-for-iTunes/assets/94881602/3db6168a-3ea3-496e-a42d-3bbfc333a211)


<br>**Be sure to change the default folder** that the Excel file will be saved to in the main code.
<br>This is done in lines 41 and 43 of the main codes and is currenty set to:
<br>file_nm = "D:\\iTunes\\Excel\\" + user_inp + ".xlsx"

<br>Finally, this is a snapshot of one output file:
![image](https://github.com/jrsousa2/Python-scripts-for-iTunes/assets/94881602/e3d63161-f639-4c6c-9374-b4ffcb8339de)

##### Feel free to send me an e-mail if you have question on how these codes work. 
<i>My contact can be found in Google</i>.