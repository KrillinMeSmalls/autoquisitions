Download Current Version of the Executable:
https://github.com/KrillinMeSmalls/autoquisitions/blob/main/autoquisitions.exe

Required in the same directory
alpha*.xlsx
upmr*.xlsx
gain*.xlsx
loss*.xlsx

Optional:
code*.txt	# PASCODE on the left, Custom Label on the right - you can use a space or a tab as the delimiter

Example codes.txt: ----
#PASCODE 	DIRECOTRATE
FT1CF99D 	Sq_Overhead
FT1CF99F 	CYO-A
FT1CF99G	CYO-B
FT1CF99H	CYO-C
-----------------------

Optional:
hold*.txt	# One PASCODE or POSITION_NUMBER per line to not auto fill

Example holds.txt: ----
#PASCODES (Sq Overhead)
FT1CF99D

#BILLETS
10600211C
10598681C
-----------------------

Notes:
Default BLSDM files (alpha roster, upmr, gains listing, & loss listing) will work - no need to re-structure or save them in any other format
No need to remove default header/footer rows - it will work with all, some, or none of them
Officer Billets are ignored (only reads Enlisted Ranks)
All filenames are NOT case sensitive, and it will ignore the middle of the filename as represented with the '*'
All files must be in the same directory, but it can be any directory with write access
If there are multiple files starting with the same name, it will prompt you for which one to use
Output filename format is "<DATE>_autoquisitions_<TIME>.xlsx"
Once all required (and optional) files are in place, just double click the executable
If you don't wish to replace pascodes with custom labels, you don't need the code*.txt file
If you don't wish to hold back any billets from being auto-filled, you don't need the hold*.txt file
Any new .csv files left in the directory when the program exits are an indication that the program failed
If it appears to fail, save/close all Excel open windows, then try again
The program is complete when you see the "Vikings, DOMINATE!" output