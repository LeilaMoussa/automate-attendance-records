"""
Get formatted list of attendees/dar al aman tutors and date
input in the form of array of names, ignore ids for now
actually, ids may be a more reliable form of input (less typos)
"""
## think input method later (consider outlook, phone directly, as automated as possible)

import gspread
import oauth2client

gc = gspread.oauth()

sh = gc.open("Attendance Records")

worksheet = sh.worksheet('Fall 20') # This semester's sheet.

attendees = ['Leila Farah Moussa', 'Hanane Nour Moussa']

#date = input("Day of event?\n") ##MM/DD/YY, must match sheet
date = '10/20/20'

date_column = worksheet.find(date, 1, None).col  # this of course assumes the date is there, otherwise, create a column

for name in attendees:
    try:
        name_row = worksheet.find(name, None, 1).row
        worksheet.update_cell(name_row, date_column, 1)
    except:
        print("Adding new name...")
        # Insert name at first row
        worksheet.insert_row([name, ''] + [0]*(date_column-3) + [1], 2) ## ISSUE!! last column for each row should be the sum: how to use excel functions from here?

print("All done!")

