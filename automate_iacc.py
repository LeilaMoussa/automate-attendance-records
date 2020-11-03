import gspread
import oauth2client
import csv

gc = gspread.oauth()

def removeBadBytes(filename):
    fi = open(filename, 'rb')
    data = fi.read()
    fi.close()
    fo = open('mynew.csv', 'wb')
    fo.write(data.replace(b'\x00', b''))
    fo.close()
    f = open('mynew.csv')
    r = csv.DictReader((line.replace('ÿþ','') for line in f), delimiter='\t')
    # This is some weird encoding issue I don't know how to fix
    return r

def update_meetings ():
    sh = gc.open("Attendance Records")

    worksheet = sh.worksheet('Fall 20') # This semester's sheet.

    #date = input('Day of event?\n') ##MM/DD/YY, must match sheet
    #date = '10/20/20'
    
    #attendees = ['Leila Farah Moussa', 'Hanane Nour Moussa']

    filename = 'C:\\Users\\mouss\\Downloads\\meetingAttendanceList.csv'
    #filename = input("Attendance list path?\n")
    reader = removeBadBytes(filename)

    date_column = worksheet.find(date, 1, None).col  # this of course assumes the date is there, otherwise, create a column

    attendees = []

    for i, data_row in enumerate(reader):
        if i == 0:
            date = data_row['Timestamp'].split(',')[0]
            print("date from file", date)
        attendees.append(data_row['Full Name'])

    attendees = set(attendees)
    
    for person in attendees:
        split = person.split('<')
        print(split)
        if len(split) == 1:
            name = person
        elif len(split) == 2:
            name, _id = split[0], split[1][:5]
            print(_id)

        print(name)
        try:
            name_row = worksheet.find(name, None, 1).row
            print("Name already exists")
            worksheet.update_cell(name_row, date_column, 1)
        except:
            print("Adding new name...")
            # Insert name at first row
            worksheet.insert_row([name, ''] + [0]*(date_column-3) + [1], 2) ## ISSUE!! last column for each row should be the sum:
                                                ##how to use excel functions from here?
            # so, insert or just write in the first empty row i find?

    print("All done!")

if __name__ == '__main__':
    event = input('"M" for monday meeting, "T" for Dar Al Aman:\n')
    if event == 'M':
        # input later
        update_meetings()
    elif event == 'T':
        update_tutoring()
    
