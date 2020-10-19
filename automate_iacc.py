"""
Get formatted list of attendees/dar al aman tutors and date
input in the form of array of names, ignore ids for now
"""
import gspread
import oauth2client

gc = gspread.oauth()

sh = gc.open("Attendance Records")

print(sh.sheet1.get('B1'))

## think input method later (consider outlook, phone directly, as automated as possible)

#date = input("Day of event?") ##MM/DD/YYYY, must match sheet

"""
identify column by date
for each name, format as needed and look for row
if found, change value from 0 to 1
if not found: either the person is not there or it's a issue of spelling (case & spaces are taken care of, so at least one letter is off)
in that case: see how different the name is from the other names, if it's a very small difference (2 letters off) assume they're the same person
otherwise, create a new row
much printing throughout
"""
