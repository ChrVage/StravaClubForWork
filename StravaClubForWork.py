import requests
import urllib3 #trengs denne?
import pandas as pd 
import json
from datetime import datetime

# Legg inn alle datoer etter den som er lagret i config-fil som skjult aktivitet på Atea Strava Admin med dato hver midnatt. yyyy.mm.dd#StravaClubForWork_Date
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

def authenticate(client_id,client_secret,refresh_token):
    payload = {
    'client_id': client_id,
        'client_secret': client_secret, 
        'refresh_token': refresh_token,
        'grant_type': "refresh_token",
       'f':'json'
    }

    auth_url = "https://www.strava.com/oauth/token"

    return requests.post(auth_url,data=payload, verify=False).json()['access_token']

def create_dates_by_clubadmin(date):
    pass

def get_new_activities(access_token,club_id):
    # create dictionary of user activities
    activities = pd.DataFrame(
        columns = [
                "Athlete",
                "Name",
                "Type",
                "Duration",
                "Distance",
                "Date",
                "id"
        ]
    )

    loop = True

    readpage = 1
    pagesize = 50
    activity_count = 0
    activities_url = "https://www.strava.com/api/v3/clubs/%s/activities" % club_id
    activity_row = 0
    activity_date = datetime.now()
    while loop:
        header = {'Authorization': 'Bearer ' + access_token}
        param  = {'per_page': pagesize, 'page': readpage}
        response = requests.get(activities_url,headers=header, params=param)
        data = response.json()

        readpage = readpage + 1
        activity_count = activity_count + len(data)
    
        for line in data :
            
            # Check for new date in activities
            taglist = line['name'].split("#")
            if len(taglist)==2:
                if taglist[1] == "AteaClubForWork_Date":
                    activity_date = datetime.strptime(taglist[0], "%Y-%m-%d")
                    continue

            activities.at[activity_row, 'Athlete']  = line['athlete']['firstname'] +"#"+ line['athlete']['lastname']
            activities.at[activity_row, 'Name']     = line['name']
            activities.at[activity_row, 'Type']     = line['type']
            activities.at[activity_row, 'Duration'] = line['elapsed_time']
            activities.at[activity_row, 'Distance'] = line['distance']
            activities.at[activity_row, 'Date']     = activity_date.strftime("%Y-%m-%d")
            activities.at[activity_row, 'id']       = "%s#%s#%s#%s" % (activities.at[activity_row, 'Athlete'], 
                                                                       activities.at[activity_row, 'Duration'], 
                                                                       activities.at[activity_row, 'Distance'],
                                                                       activity_date.strftime("%d"))

            activity_row = activity_row + 1
 

        print("Page: %i, len(data): %d " % (readpage-1,len(data)))
        # https://stackoverflow.com/questions/17071871/how-to-select-rows-from-a-dataframe-based-on-column-values
        if activity_count>300: 
            loop = False

        if len(data)<pagesize :
            loop = False
    
    print("activities found: %i" % activity_count)
    return activities

def main():
    # Read config.json file
    with open("config.json", "r") as jsonfile:
        config = json.load(jsonfile)

    access_token = authenticate(config["client_id"],config["client_secret"],config["refresh_token"])
    print('Access token:"%s"\n' % access_token)
    date = config['last_date']
    print('Date:%s' % date)
    # date = create_dates_by_clubadmin(date)
    new_activities = get_new_activities(access_token,config["club_id"])

    i = len(new_activities.index)
    print("len new_activities: %s" % i)
    print("Oldest activity(%d): %s %s %s" % (i-1,new_activities.at[i-1, 'Athlete'],new_activities.at[i-1, 'Name'],new_activities.at[i-1, 'Date']))

    # Write data to an Excel spreadsheet pr week
    FileName = 'Activitylist %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    
    df = pd.DataFrame(new_activities)
    df.to_excel(FileName, index=False)
    
    with open("config.json", "w") as jsonfile:
        myJSON = json.dump(config, jsonfile, indent=2) # Writing to the file
        jsonfile.close()

if __name__ == "__main__":
    main()