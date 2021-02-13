import gspread
import oauth2client
import csv
import string
import pprint

gc = gspread.oauth()

semester = 'Fall 20'
board_members = ('Leila Farah Moussa') # And co.

def removeBadBytes(filename):
    # This is some weird encoding issue I don't know how to fix.
    fi = open(filename, 'rb')
    data = fi.read()
    fi.close()
    fo = open('mynew.csv', 'wb')
    fo.write(data.replace(b'\x00', b''))
    fo.close()
    f = open('mynew.csv')
    r = csv.DictReader((line.replace('ÿþ','') for line in f), delimiter='\t')
    return r

def get_first_empty_row(recs: list) -> int:
    for i, row in enumerate(recs):
        if row['Name'] == '':
            return i

def get_csv_data(reader) -> tuple:
    attendees = set()
    for i, data_row in enumerate(reader):
        # Encoding problems with Excel files turned to CSV!
        if i == 0:
            date = data_row['Timestamp'].split(',')[0]
            date = date[:-2]
        attendees.add(data_row['Full Name'])
    return attendees, date

def find_name(recs: list, name: str) -> int:
    for i, row in enumerate(recs):
        if row['Name'] == name:
            return i
    return -1

def populate_records(attendees: set, recs: list, date: str) -> None:
    first_empty = get_first_empty_row(recs)
    if first_empty is None:
        print('You need more empty rows.')
    
    for person in attendees:
        split = person.split('<')
        if len(split) == 1:
            name, _id = person, ''
        elif len(split) == 2:
            name, _id = split[0].strip(), split[1].strip()[:5]
        name = name.title()

        index = find_name(recs, name)
        if index < 0:
            # First time attending.
            recs[first_empty]['Name'] = name
            recs[first_empty]['ID'] = _id
            recs[first_empty][date] = 1
            first_empty += 1
        else:
            recs[index][date] = 1
            recs[index]['ID'] = _id

def write_back(recs: list, worksheet) -> None:
    number_rows = len(recs)
    lists = []
    for elt in recs:
        lists.append(list(elt.values()))
    worksheet.update(f'A2:{number_rows+1}', lists)

def update_meetings():
    print("Okay!")
    
    sh = gc.open("Attendance Records")
    worksheet = sh.worksheet(semester)

    recs = worksheet.get_all_records()
    
    del recs[-1]
    for row in recs:
        del row['Count of attendance so far']

    #pp = pprint.PrettyPrinter()
    #pp.pprint(recs)

    filename = input("Attendance list path?\n")
    # filename = 'sheet.csv'
    reader = removeBadBytes(filename)
    attendees, date = get_csv_data(reader)

    populate_records(attendees, recs, date)

    write_back(recs, worksheet)

    print("All done!")
    return

    ########################################
    # May or may not come back here...
    
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
    event_row = worksheet.find('Event Attendance', None, None) or 44  # Should be length of records!
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
    # event = 'M'
    if event == 'M':
        update_meetings()
    elif event == 'T':
        update_tutoring()
    print("Have a nice day <3")
