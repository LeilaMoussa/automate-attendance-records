import gspread
import oauth2client
import csv
import string
import time

gc = gspread.oauth()

semester = 'Fall 20'
board_members = ('Leila Farah Moussa', 'Hanane Nour Moussa', 'Yassir Benabdallah', 'Naeem Nisar Sheikh')

def removeBadBytes(filename):
    # This is some weird encoding issue I don't know how to fix.
    fi = open(filename, 'rb')
    data = fi.read()
    fi.close()
    fo = open('mynew.csv', 'wb')  # Temp file.
    fo.write(data.replace(b'\x00', b''))
    fo.close()
    f = open('mynew.csv')
    r = csv.DictReader((line.replace('ÿþ','') for line in f), delimiter='\t')
    return r

def getRangeEnd(first_empty_row, number_rows, date_column) -> str:
    """
    Returns the last cell we're writing to in A1 notation
    """
    col = string.ascii_uppercase[date_column-1]
    row = first_empty_row + number_rows - 1
    cell = col + str(row)
    return cell

def getFirstEmptyRow(worksheet) -> int:
    row = 2
    while worksheet.cell(row, 1) != '':
        row += 1
    return row

def update_meetings():
    print("Okay!")
    
    sh = gc.open("Attendance Records")
    worksheet = sh.worksheet(semester)

    #filename = input("Attendance list path?\n")
    filename = 'sheet.csv'
    reader = removeBadBytes(filename)

    attendees = set()
    for i, data_row in enumerate(reader):
        # Very bad encoding problems with Excel files turned to CSV!
        if i == 0:
            date = data_row['Timestamp'].split(',')[0]
            date = date[:-2]  # Just for formatting, remove '20'.
        attendees.add(data_row['Full Name'])

    date_column = worksheet.find(date, 1, None).col
    # All dates are already there.
    
    missing = []
    existing = 0
    for person in attendees:
        split = person.split('<')
        if len(split) == 1:
            name, _id = person, ''
        elif len(split) == 2:
            name, _id = split[0], split[1][:5]
        name = name.title()  # Capitalize names.

        try:
            name_row = worksheet.find(name, None, 1).row
            print("Name already exists")
            existing += 1
            worksheet.update_cell(name_row, date_column, 1)
        except:
            print("Adding new name...")
            missing.append([name, _id]+[0]*(date_column-3)+[1])

    print("Waiting for a minute to avoid exhausting the resource. Go drink some water or something.")
    time.sleep(60)
    first_empty_row = getFirstEmptyRow(worksheet)

    last_cell = getRangeEnd(first_empty_row, len(missing), date_column)
    worksheet.update(f'A{first_empty_row}:{last_cell}', missing)

    print("All done!")
    return
    ########################################
    print("The 3 most active non-board members...")
    count_column = worksheet.find('Count of attendance so far', 1, None).col
    counts = worksheet.col_values(count_column)  # This is an ordered list.
    print("counts", counts)
    max_val = max(counts)
    rows = [i+2 for i, x in enumerate(counts) if x == max_val]
    for row in rows:
        frequent_attendee = worksheet.cell(row, 1)
        print("freq attendee", frequent_attendee)
        if frequent_attendee not in board_members:
            print(frequent_attendee, ", ")
    print(f'have come {max_val} times.')
    
    print("The most popular monday meeting so far...")
    # find single max value in last row
    event_row = worksheet.find('Event Attendance', None, None) or 44  # resolves to boolean or actual choices?
    events = worksheet.row_values(event_row)
    print("events row", events)
    max_att = 0
    for i, val in events:
        if val > max_att:
            max_att = val
            letter = i+2 ## not sure about this, need to look at 'events'
    event_date = worksheet.cell(1, letter)
    print("was on the", event_date)

def update_tutoring():
    print("Okay!")
    # for later

if __name__ == '__main__':
    event = input('"M" for monday meeting, "T" for Dar Al Aman:\n')
    if event == 'M':
        update_meetings()
    elif event == 'T':
        update_tutoring()
    print("Have a nice day <3")
