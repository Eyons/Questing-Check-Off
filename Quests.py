import json
from pprint import pprint

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from difflib import SequenceMatcher

def similar(a,b):
    return SequenceMatcher(None, a, b).ratio()

with open('Tormented.json') as f:
    data = json.load(f)

# pprint(list(data['quests'][0].items()))

# for i in range(data['quests']):
#     for k,v in list(data['quests'][i].items()):
#         print (k,v)


# Creates a list of all quests that are not completed
questIndex = []
for i in data['quests']:
    for k, v in i.items():
        if k == "status" and v != "COMPLETED":
            questIndex.append(i)


# Setups the API for accessing the specified google sheet
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Copy of Level 3 to All Capes--Ultimate Runescape Guide.xlsx').sheet1

# Column that was desired is the Goals tab, aka column 1
goals = sheet.col_values(1)

# temp = sheet.cell(1,1).value
# sheet.update_cell(1,1, temp+'*')

# Creates a list of the quests without certain words in them
questNames = []
for i in range(len(questIndex)):
    if 's Task' in questIndex[i]['title'] or 'Elemental Workshop' in questIndex[i]['title']:
        continue
    elif '(miniquest)' in questIndex[i]['title']:
        questNames.append(questIndex[i]['title'].replace(' (miniquest)', ''))
    elif ' (saga)' in questIndex[i]['title']:
        questNames.append(questIndex[i]['title'].replace(' (saga)', ''))
    else:
        questNames.append(questIndex[i]['title'])

# Manages finding quests in the spreadsheet with similar ones in the given json and updating them with an asterisk
for i in goals:
    if "*" in i:
        i = i.replace('*','')

nonFoundQuests = []
for i in range(len(questNames)):
    found = False
    for j in range(len(goals)):
        if similar(questNames[i], goals[j]) > .85:
            # cellVal = sheet.cell(j+1, 1).value
            # sheet.update_cell(j+1, 1, cellVal + "*")
            print("Updating {0} with {1}* in row {2}".format(questNames[i], 0, j+1))
            found = True
            break
    if not found:
        nonFoundQuests.append(i)


    


# Stores the remaining quests that weren't found in specified file
file = open("tempfile.txt","w")
for i in nonFoundQuests:
    print("Storing {0} in file".format(questNames[i]))
    file.write(questNames[i] + "\n")
file.close()

print("Done")