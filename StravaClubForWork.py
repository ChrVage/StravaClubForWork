import requests
import urllib3
import pandas as pd 
import json
from datetime import datetime

# Verifiser at å legge inn et par aktiviteter tilbake i tid for å se om de dukker opp før noe som allerede er i listen
# Del opp i funksjoner
# Les inn Config fil med minste mulige data, men data om admin-bruker-token, siste dato 
# Legg inn alle datoer etter den som er lagret i config-fil som skjult aktivitet på Atea Strava Admin med dato hver midnatt. yyyy.mm.dd#StravaClubForWork_Date
# Les ut ny fil fra API, lag ID kolonne og dato kolonne
#  Lag ID kolonne bestående av navn#sekunder aktivitet#distanse (Det som ikke endrer seg.).
#  Legg inn dagens dato, til aktitet fra Atea Strava Admin markerer datoskillet til dagen før. (Det kan finnes dager uten aktivitet)
# Fjern Atea Strava Admin - aktiviteter.
# Les inn eksisterende fil som inneholder alt fra i år.
# Bruk filen som nettopp er laget som utgangspunkt for ny fil
# Fjern de første linjene i eksisterende fil (Antall = maks det som finnes i eksisterende datasett +?:Hvor mange aktiviteter kan være slettet eller lagt til?).
# Lagre ny fil med alle aktiviteter (Sjekk at PowerBI klarer å hente json dataene herfra - eventuelt lagre som .csv (test komma i aktivitetsnavn?).)
# Sjekk mot medlemslisten hvem som har like navn hver mandag. 
# Lag oversikt over antall medlemmer i klubben.
# Lag trekningsliste i Excel for forrige uke hver gang man starter på ny uke.
#   * Fjern de som har kun en aktivitet, og de som ikke har navnebror og har mer enn 4 aktiviteter.
#   * Fjern aktiviteter som er mindre enn 900 sekunder
#   * Gi alle aktiviteter random nummer, laveste vinner.


auth_url = "https://www.strava.com/oauth/token"
activities_url = "https://www.strava.com/api/v3/clubs/10971/activities"
members_url = ""

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

# Write data to an Excel spreadsheet pr week
df = pd.DataFrame(data=TheData, columns=allColumns)
df.to_excel(FileName, index=False)
