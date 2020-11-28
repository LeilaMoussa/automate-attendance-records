import gspread
import oauth2client
import csv
import string

gc = gspread.oauth()

semester = 'Fall 20'

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

def update_meetings():
    print("Okay!")
    
    sh = gc.open("Attendance Records")
    worksheet = sh.worksheet(semester)

    filename = input("Attendance list path?\n")
    reader = removeBadBytes(filename)

    attendees = set()
    for i, data_row in enumerate(reader):
        # print("data row", data_row)
        # Very bad encoding problems with Excel files turned to CSV!
        # Solution: start with a CSV.
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
            # print("Name already exists")
            existing += 1
            worksheet.update_cell(name_row, date_column, 1)
        except:
            # print("Adding new name...")
            missing.append([name, _id]+[0]*(date_column-3)+[1])

    first_empty_row = existing + 2

    last_cell = getRangeEnd(first_empty_row, len(missing), date_column)
    worksheet.update(f'A{first_empty_row}:{last_cell}', missing)

    print("All done!")
    # 3 most active members so far
    # find at most three occurrences of the max value in count column
    # print corresponding people
    
    # most popular event so far
    # find single max value in last row
    # print its corresponding date

def update_tutoring():
    print("Okay!")
    # for later

if __name__ == '__main__':
    event = input('"M" for monday meeting, "T" for Dar Al Aman:\n')
    if event == 'M':
        update_meetings()
    elif event == 'T':
        update_tutoring()
