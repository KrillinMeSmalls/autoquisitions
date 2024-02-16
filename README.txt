
Required in the same directory (filenames must start with these 4-5 characters - case-insensitive.  It will not find files starting with anything else such as "20240201 UMPR...")
alpha*.xlsx
upmr*.xlsx
gain*.xlsx
loss*.xlsx

Optional in the same directory:
code*.txt			# Used to replace the PASCODE with customizable identifiers.  Format is PASCODE on the left with the custom label on the right. You can use either a space or a tab as the delimiter between the two fields.
hold*.txt			# Used to exclude a billet or entire PASCODE from being processed by this program.  Format is one PASCODE or POSITION_NUMBER per line.  
				## For both optional files, you may include comment lines as shown below as long as they are entirely on other lines.

---- Example of codes.txt: ----
#PASCODE 	DIRECOTRATE
AB1CD23E 	Sq_Overhead
AB1CD23F 	Flight A
AB1CD23G	Flight B
AB1CD23H	Flight C
-------------------------------

---- Example of holds.txt: ----
#PASCODES (Sq Overhead)
AB1CD23E

#BILLETS
10400111A
10598765B
-------------------------------


How To Read the Output File "<DATE>_autoquisitions_<TIME>.xlsx":

- You can probably notice the output is in three undivided sections:
-- The top section shows all billet position numbers without an assigned member.  These include billets that were empty or that will be empty because of losses.
--  The middle section shows all members that are not assigned to a billet.  Keep in mind this is after the script optimized your manning, so this section includes all members on the losses listing, any gains that were unable to be placed into an open billet, and potentially previously unassigned members.
-- The bottom section is your listing of all billeted members and their position numbers.  This includes possible multi-incumbent positions.  This script will not double up and positions, with the exception of positions marked as higher-ratio exception.  For instance, the current hard coded ratio for 3-lvls is 6 members to 1 position number.
-- Below the bottom section you may see two lines of additional information.  In the event members are found on the Alpha Roster, but not the UPMR (or vice versa), those names will be identified here.  They should still be accounted for in the sections above and should not impact the overall output.  These are just highlighted for your awareness.

- So, what about these extra columns with 'Y's and such:
-- GAIN/LOSS should be obvious.  This will be marked according to where the member was identified on either gains or losses listings.  A blank field means they are on neither.
-- REQUISTION will only contain information if the billet is being marked empty.  It may highlight a billet as 3-lvl to help identify possible positions only filled by accessions versus PCSing members.
-- HOLD BILLET will only contain a 'Y' if the billet position matched an entry in the optional "hold*.txt" file.  This billet will remain empty if already empty or has an assigned member that is on the losses roster.  Use case for this would be Sq leadership billets (SEL/Ops Supt) that you need to advertise specifically without this program placing a member on the gains roster in the billet.
-- MODIFIED BILLET will be marked whenever a change is made/suggested from what is shown on the BLSDM reports.  This includes billets that had losses removed from them, empty billets that had gains assigned to them, over-billeted positions that were corrected, and such.  If this column has a 'Y' you probably need to be making an official change on your books.
-- NEW INCUMBENT will be marked when the assigned member was not in this position previously.  The billet was previously empty, had a member removed because they were on the losses roster, or the previous billet occupant was re-billeted to a more appropriate billet.
-- FAILED PLACEMENT will be marked on members from the gains roster unassigned to a billet, previously unassigned members that remain unassigned, or over-billeted positions that remain unassigned.  This program will make matches based on rank & AFSC.  If it cannot find an appropriate billet to match the members information, the member will not be assigned to a position number.  These will require manual evaluation to determine if they should occupy a non-matching billet or remain unassigned.
-- OVERLOADED POSITION is a quick look to know which, if any, billets exceed the allowed ratio.  Again, this is typically 1-to-1, or in the case of 3-levels it is currently 6-to-1.  You will likely see a FAILED PLACEMENT when you see an OVERLOADED POSITION
-- OFFICER is only there to help with filtering within Excel.
-- REMARKS are intended to help with tracking where any moved members were on the books previously.


Additional Notes:
Default BLSDM files (alpha roster, upmr, gains listing, & loss listing) will work - no need to re-structure or save them in any other format
No need to remove default header/footer rows - it will work with all, some, or none of them
All filenames are NOT case sensitive, and it will ignore the middle of the filename as represented with the '*'
All files must be in the same directory, but it can be any directory with write access
If there are multiple files starting with the same name, it will prompt you for which one to use
Once all required (and optional) files are in place, just double click the executable
If you don't wish to replace pascodes with custom labels, you don't need the code*.txt file
If you don't wish to hold back any billets from being auto-filled, you don't need the hold*.txt file
Any new .csv files left in the directory when the program exits are an indication that the program failed - some future version will remove the .csv files being created
If it appears to fail, save/close all Excel open windows, then try again.  It is important that none of these files are already open in Excel!
The program is complete when you see the credits output and a new .xlsx output file in the "<DATE>_autoquisitions_<TIME>.xlsx" format
