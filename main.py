from icalevents.icalevents import events
import datetime
import dateutil
import openpyxl

# create table for cell letter lookup
cellNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']

# work out the dates required for last month
todayDate = datetime.date.today()
thisMonthStart = datetime.datetime(todayDate.year, todayDate.month, 1, 0, 0, 0, 0)
lastMonthEnd = (thisMonthStart - datetime.timedelta(days=1)) + datetime.timedelta(hours = 23, minutes=59, seconds=59)
lastMonthStart = datetime.date(lastMonthEnd.year, lastMonthEnd.month, 1)
startDate = lastMonthStart - datetime.timedelta(days=1)
print(lastMonthEnd)

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

# open the blank expenses form
workbook = openpyxl.load_workbook(
    "C:/Users/aaron/OneDrive - Cloudbass/HR/Expense Claims/Expenses_Form_August_2016 - Blank.xlsx")
# set the correct sheet
mainSheet = workbook['August 2016']

print(len(calendar))

#got through the days and check each for the summery contents
for day in range(len(calendar)):
    cellName = cellNames[day + 1]
    print(f"{calendar[day][0].start.day}, {calendar[day][0].start.month}, {calendar[day][0].start.year}")
    if (day == 0):
        continue
    elif ("Onsite" in calendar[day][0].summary) and ("Onsite" not in calendar[day - 1][0].summary):
        mainSheet[cellName + "7"] = 1
    elif ("Onsite" in calendar[day][0].summary) and ("Onsite" in calendar[day - 1][0].summary):
        mainSheet[cellName + "8"] = 1
    elif ("Onsite" not in calendar[day][0].summary) and (len(calendar[day]) > 1):
        print(calendar[day][0].summary)
        for event in calendar[day]:
            if ("Travel" in event.summary) or ("travel" in event.summary) or ("TVL" in event.summary) or ("tvl" in event.summary):
                print(event.summary)
                mainSheet[cellName + "9"] = 1
                break

# add the current month
mainSheet["N3"] = lastMonthStart.strftime('%d/%m/%Y')
mainSheet["N3"].number_format = 'mmm-yy'

# save the new file
workbook.save(
    filename= "C:/Users/aaron/OneDrive - Cloudbass/HR/Expense Claims/" + lastMonthStart.strftime(
        "%m") + " - Expenses_Form_August_2016 - AM" + lastMonthStart.strftime(
        "%B%Y") + ".xlsx")
