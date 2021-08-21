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

# create array to hold each day
calendar = []

# get ical file for the last month (plus one day before as well to check if the first day is 15 or 20)
print("Getting calendar")
iCal = events("http://www.cbunit.co.uk/ccs/ical/AM2363.ics", start=startDate, end=lastMonthEnd)

# sort the calendar by date
print("sorting calendar")
iCal.sort(key=lambda x: x.start)

# sort the iCal file into a 2D array to hold the events by day
for i in range(len(iCal)):
    # store the event in array
    tempList = [iCal[i]]
    if (i == 0):
        # if the first item then append to list
        calendar.append(tempList)
    else:
        # if events are on the same day then store together
        if (iCal[i].start.day == iCal[i-1].start.day):
            calendar[-1].append(iCal[i])
        else:
            # store event as new day
            calendar.append(tempList)

for i in calendar:
    print(i)

# open the blank expenses form
workbook = openpyxl.load_workbook(
    "C:/Users/aaron/OneDrive - University of Derby/Personal/Work/CB/Expense Claims/Expenses_Form_August_2016 - Blank.xlsx")
# set the correct sheet
mainSheet = workbook['August 2016']

print(len(calendar))

#got through the days and check each for the summery contents
for day in range(len(calendar)):
    cellName = cellNames[day + 1]
    if (day == 0):
        continue
    elif ("Onsite" in calendar[day][0].summary) and ("Onsite" not in calendar[day - 1][0].summary):
        mainSheet[cellName + "7"] = 1
    elif ("Onsite" in calendar[day][0].summary) and ("Onsite" in calendar[day - 1][0].summary):
        mainSheet[cellName + "8"] = 1
    elif ("Onsite" not in calendar[day][0].summary) and (len(calendar[day]) > 1):
        print("ping")
        print(calendar[day][0].summary)
        for event in calendar[day]:
            if ("Travel" in event.summary) or ("travel" in event.summary):
                print("pong")
                mainSheet[cellName + "9"] = 1
                break

# add the current month
mainSheet["N3"] = todayDate.strftime('%d/%m/%Y')
mainSheet["N3"].number_format = 'mmm-yy'

# save the new file
workbook.save(
    filename="C:/Users/aaron/OneDrive - University of Derby/Personal/Work/CB/Expense Claims/Expenses_Form_August_2016 - AM" + lastMonthStart.strftime(
        "%B%Y") + "TEST.xlsx")
