import requests
import urllib3
import pandas as pd 
import json
from datetime import datetime

# Todo: Create a list with all activities, store as file and append to it.
# Create New Excel-file pr week.
# Finn første match og erstatt resten.


auth_url = "https://www.strava.com/oauth/token"
activities_url = "https://www.strava.com/api/v3/clubs/10971/activities"

payload = {
    'client_id': "60365",
    'client_secret': 'd6eaed2e3d1f1d8a154044d5984769536dc06674', 
    'refresh_token': '18a7f1b532db6ffb0085055a93831b5ad5f4b00a',
    'grant_type': "refresh_token",
    'f':'json'
}

access_token = requests.post(auth_url,data=payload, verify=False).json()['access_token']
print('Access token:"%s"\n' % access_token)

header = {'Authorization': 'Bearer ' + access_token}

now = datetime.now().strftime("%Y.%m.%d %H%M")
FileName = 'Activitylist %s.xlsx' % now

fullname = list()
activityname = list()
activitytype = list()
activityduration = list()
activitydistance = list()
fullLine = list()
allColumns = ['Name', 'Activity', 'Type', 'Duration', 'Distance']
loop = True

readpage = 1
pagesize = 50
activities = 0



while loop:
    param = {'per_page': pagesize, 'page': readpage}
    response = requests.get(activities_url,headers=header, params=param)

    readpage = readpage + 1
    data = response.json()
    activities = activities + len(data)

    for line in data :
        fullname =          fullname +          [line['athlete']['firstname'] +"#"+ line['athlete']['lastname']]
        activityname =      activityname +      [line['name']]
        activitytype =      activitytype +      [line['type']]
        activityduration =  activityduration +  [line['elapsed_time']]
        activitydistance =  activitydistance +  [line['distance']]
        
        # Break in right place
        taglist = line['name'].split("#")
        print("taglist[0]: %s" % taglist[0])

        # if taglist[0] in ("UkeStart","") :
        #     loop = False
        #     break
    
    print("data has %d items" % len(data))
    print("readpage: %i" % readpage)
    if activities>300: 
        loop = False

    if len(data)<pagesize :
        loop = False

print("activities found: %i" % activities)

TheData = list(zip(fullname,activityname,activitytype,activityduration,activitydistance))

# Write data to an Excel spreadsheet.
df = pd.DataFrame(data=TheData, columns=allColumns)
df.to_excel(FileName, index=False)
