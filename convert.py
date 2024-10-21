from xlwt import Workbook
from time import sleep
from os import system
import re

found = False
while not found:
    # INSERT FILE NAME HERE
    filename = ""
    try:
        with open(f"Input/{filename}", "r") as f:
            email = f.read()
        found = True
    except FileNotFoundError:
        print("File not found...")
        found = True

newFilename = filename.replace(".txt", "")
newFilename = newFilename.replace('Shoutbomb', '')

queries = {
    "Hold text notices sent for the month": 0,
    "Hold cancel notices sent for the month": 0,
    "Overdue text notices sent for the month": 0,
    "Overdue items eligible for renewal, text notices sent for the month": 0,
    "Overdue items ineligible for renewal, text notices sent for the month": 0,
    "Overdue (text) items renewed successfully by patrons for the month": 0,
    "Overdue (text) items unsuccessfully renewed by patrons for the month": 0,
    "Renewal text notices sent for the month": 0,
    "Items eligible for renewal text notices sent for the month": 0,
    "Items ineligible for renewal text notices sent for the month": 0,
    "Items (text) renewed successfully by patrons for the month": 0,
    "Items (text) unsuccessfully renewed by patrons for the month": 0,
}

libraries = {
    "Atkinson": 0, 
    "Bay View": 0, 
    "Brown Deer MAIN": 0, 
    "Brown Deer Drive-Up": 0, 
    "Capitol": 0,
    "Center St.": 0,
    "Central MAIN": 0,
    "Central Drive-Up": 0,
    "Cudahy MAIN": 0,
    "Cudahy Locker": 0,
    "East MAIN": 0, 
    "East Locker": 0, 
    "Franklin MAIN": 0, 
    "Franklin Locker": 0, 
    "Good Hope": 0, 
    "Greendale": 0, 
    "Greenfield": 0,
    "Hales Corners": 0,
    "Martin Luther King": 0, 
    "Mitchell St": 0, 
    "North Shore": 0, 
    "Oak Creek MAIN": 0,
    "Oak Creek Locker": 0,
    "Shorewood MAIN": 0, 
    "Shorewood Locker": 0, 
    "South Milwaukee": 0,
    "St. Francis": 0,
    "Tippecanoe": 0, 
    "Villard": 0, 
    "Washington Park": 0, 
    "Wauwatosa": 0, 
    "West Allis": 0, 
    "West Milwaukee": 0, 
    "Whitefish Bay MAIN": 0, 
    "Whitefish Bay Locker": 0,  
    "Zablocki": 0, 
}

workbook = Workbook()
splittedEmail = email.split("=TOTALS BY BRANCH=")[0]

def parse(data, query):
    for line in data.splitlines():
        for key in query.keys():
            if key in line:
                newLine = line.split(" = ")
                newLine = int(newLine[1])
                query[key] = newLine
    return query


def parseNoticeTotals(branch, queries):
    matches = re.findall(r'([^=]+)=\s*(\d+)\s+', branch)
    parsed_data = {query: 0 for query in queries}
    
    for text_part, number in matches:
        text_part = text_part.strip()
        number = int(number.strip()) 
        
        for query in queries:
            if query in text_part:
                parsed_data[query] = number
                break 
    
    return parsed_data

# First Sheet
totalsByBranch = workbook.add_sheet(f"Totals {newFilename}")

emailText = splittedEmail.split("=TOTALS=")[0]
queriesList = list(queries.keys())
for query in queriesList:
    totalsByBranch.write(int(queriesList.index(query)+1), 0, query)
row = 0
column = 0
for branch in emailText.split("Branch:: "):
    for library in libraries:
        row = 0
        if library == branch[0:len(library)]:
            column += 1
            totalsByBranch.write(0, column, library)
            libQueries = parseNoticeTotals(branch, queries.copy())
            for query in libQueries.values():
                row += 1
                totalsByBranch.write(row, column, query)

row = 0 
column += 1
totals = parse(splittedEmail.split("=TOTALS=")[1], queries.copy())
totalsByBranch.write(row, column, f"Totals")
for query in totals.values():
    row += 1
    totalsByBranch.write(row, column, query)

# Second Sheet
def parseTotalsSent(branch, libraries):
    # Regex to find occurrences of "<library name> sent for the month, this many text notices = <number>"
    matches = re.findall(r'(.*?)\s*sent for the month,\s*this many text notices\s*=\s*(\d+)', branch)
    parsed_data = {library: 0 for library in libraries}
    
    # Iterate over matches
    for library_name, number in matches:
        library_name = library_name.strip()
        number = int(number.strip())
        for key in parsed_data.keys():
            if library_name.startswith(key):
                parsed_data[key] = number
                break
    
    return parsed_data

column = 1
textNotices = workbook.add_sheet(f"Text Notices Sent {newFilename}")
textNotices.write(1, 0, "Total Text Notices")
splittedEmail = email.split("=TOTALS BY BRANCH=")[1]
emailText = splittedEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[0]
values = parseTotalsSent(emailText, libraries)
for library in libraries:
    textNotices.write(0, column, library)
    textNotices.write(1, column, values[library])
    column += 1

# Third Sheet
column = 1
registeredUsers = workbook.add_sheet(f"Registered Patrons {newFilename}")
registeredUsers.write(1, 0, "Total Registered Users")
emailText = splittedEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[1]

def parseRegisteredPatrons(text, libraries):
    matches = re.findall(r'(.*?)\s+has\s+(\d+)\s+registered patrons for text notices', text)
       
    registered_patrons = {library: 0 for library in libraries}
    
    for library_name, number in matches:
        library_name = library_name.strip()
        number = int(number.strip())
        for key in registered_patrons.keys():
            if library_name.startswith(key):
                registered_patrons[key] = number 
                break
    
    return registered_patrons

patronValues = parseRegisteredPatrons(emailText, libraries)
print(patronValues)

for library in libraries:
    registeredUsers.write(0, column, library)
    registeredUsers.write(1, column, patronValues[library])
    column += 1

# Save the workbook
workbook.save(f"Output/{filename.replace('.txt', '.xls')}")
print("Saved Successfully...")
