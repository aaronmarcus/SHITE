from icalevents.icalevents import events
import datetime
import dateutil
import openpyxl

# create table for cell letter lookup
cellNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']

# work out the dates required for last month
todayDate = datetime.date.today()
thisMonthStart = datetime.date(todayDate.year, todayDate.month, 1)
lastMonthEnd = thisMonthStart - datetime.timedelta(days=1)
lastMonthStart = datetime.date(lastMonthEnd.year, lastMonthEnd.month, 1)
startDate = lastMonthStart - datetime.timedelta(days=1)

# create valid events array
WWCalendar = []

# get ical file for the last month (plus one day before as well to check if the first day is 15 or 20)
calendar = events("http://www.cbunit.co.uk/ccs/ical/AM2363.ics", start=startDate, end=lastMonthEnd)

print(len(calendar))

for i in calendar:
    print(i)

print("\n \n")

# take only WW events from the file
for x in range(len(calendar)):
    if "WW - " in calendar[x].summary:
        WWCalendar.append(calendar[x])

for i in WWCalendar:
    print(i)

# open the blank expenses form
workbook = openpyxl.load_workbook(
    "C:/Users/aaron/OneDrive - University of Derby/Personal/Work/CB/Expense Claims/Expenses_Form_August_2016 - Blank.xlsx")
# set the correct sheet
mainSheet = workbook['August 2016']

# put in the correct value
for day in range(len(WWCalendar)):
    cellName = cellNames[day - 1]
    if (day == 0):
        continue
    elif ("Onsite" in WWCalendar[day].summary) and ("Onsite" not in WWCalendar[day - 1].summary):
        mainSheet[cellName + "7"] = 1
    elif ("Onsite" in WWCalendar[day].summary) and ("Onsite" in WWCalendar[day - 1].summary):
        mainSheet[cellName + "8"] = 1
    elif ("Travel") in WWCalendar[day].summary:
        mainSheet[cellName + "9"] = 1

# add the current month
mainSheet["N3"] = todayDate.strftime('%d/%m/%Y')
mainSheet["N3"].number_format = 'mmm-yy'

# save the new file
workbook.save(
    filename="C:/Users/aaron/OneDrive - University of Derby/Personal/Work/CB/Expense Claims/Expenses_Form_August_2016 - AM" + lastMonthStart.strftime(
        "%B%Y") + "TEST.xlsx")









