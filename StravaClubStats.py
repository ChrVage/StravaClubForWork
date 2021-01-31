import requests
import urllib3 #trengs denne?
import pandas as pd
import json
import errno
from datetime import datetime

# Dette må gjøres før programmet når v1.0:

# Rydd config.json og oppdater Readme

# Finn og fjern de siste aktivitetene i den gamle filen som finnes i den nye (basert på id)

# Legg inn manuelle aktiviteter i Atea Strava Admin etter dato i config.json fil (eller kanskje faktiske innlagte dato?)

# Sjekk mot medlemslisten hvem som har like navn hver mandag
# Lag fil med run-statistics
#   Lag oversikt over antall medlemmer i klubben
#   Hvor mange aktiviteter som var nye side sist

# Lag trekningsliste i Excel for forrige uke hver gang man starter på ny uke
#   * Fjern de som har kun en aktivitet, og de som ikke har navnebror og har mer enn 4 aktiviteter
#   * Fjern aktiviteter som er mindre enn 900 sekunder
#   * Gi alle aktiviteter random nummer, laveste vinner

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

# Read all stored activities from file. Create file if it doesn't exist
def read_activities_from_file(file_name,activities):
    try:
        activities = pd.read_pickle(file_name)
        print('Opens file: %s' % file_name)
    except OSError as e:
        if e.errno == errno.ENOENT:
            activities.to_pickle(file_name)
            print('Creates file: %s' % file_name)
        else:
            print('Oops')
            raise

    return activities

# Get all new activities from Strava API
def get_new_activities_from_strava(access_token,club_id,activities):
    loop = True

    readpage = 1
    pagesize = 50
    activities_url = "https://www.strava.com/api/v3/clubs/%s/activities" % club_id
    counter = 0
    activity_date = datetime.now()
    

    while loop:
        header = {'Authorization': 'Bearer ' + access_token}
        param  = {'per_page': pagesize, 'page': readpage}
        response = requests.get(activities_url,headers=header, params=param)
        data = response.json()

        readpage = readpage + 1
      
    
        for line in data :
            # Check for new date in activities
            taglist = line['name'].split("#")
            if len(taglist)==2:
                if taglist[1] == "AteaClubStats_Date":
                    activity_date = datetime.strptime(taglist[0], "%Y-%m-%d")
                    continue

            activities.at[counter, 'Athlete']  = line['athlete']['firstname'] +"#"+ line['athlete']['lastname']
            activities.at[counter, 'Name']     = line['name']
            activities.at[counter, 'Type']     = line['type']
            activities.at[counter, 'Duration'] = line['elapsed_time']
            activities.at[counter, 'Distance'] = line['distance']
            activities.at[counter, 'Date']     = activity_date.strftime("%Y-%m-%d")
            activities.at[counter, 'id']       = "%s#%s#%s#%s" % (  activities.at[counter, 'Athlete'], 
                                                                    activities.at[counter, 'Duration'], 
                                                                    activities.at[counter, 'Distance'],
                                                                    activities.at[counter, 'Date'])

            counter = counter + 1
        
        if counter>300: 
            loop = False

        if len(data)<pagesize :
            loop = False
    
    print("activities found: %i" % (counter-1) )
    return activities

# Check "old" activites, replace when the same activity exist in "new" list. Append all that does not exist.
def remove_duplicate_activities(all_activities,new_activities):
    # appen new activities backwards and reset index
    all_activities = all_activities.append(new_activities.iloc[::-1])
    all_activities.reset_index(drop=True, inplace=True)
    
    return all_activities

def main():
    # Read config.json file
    with open("config.json", "r") as jsonfile:
        config = json.load(jsonfile)

    # Get an access token to authenticate when getting data from Strava
    access_token = authenticate(config["client_id"],config["client_secret"],config["refresh_token"])
    print('Access token:"%s"\n' % access_token)
        
    # Define data columns
    columns = [ "Athlete", "Name", "Type", "Duration", "Distance", "Date", "id" ]
    
    # Get stored data
    all_activities = pd.DataFrame(columns=columns)
    all_activities = read_activities_from_file(config["data_file"], all_activities)

    # Get data from Strava
    new_activities = pd.DataFrame(columns=columns)
    new_activities = get_new_activities_from_strava(access_token, config["club_id"], new_activities)
    new_activities.set_index('id')

    # Add the new activites to the data already stored, but skip existing activities
    # fin_activities = pd.DataFrame(columns=columns)
    # fin_activities = remove_duplicate_activities(all_activities, new_activities)

    # Debug: Write the new activities to an Excel file
    # FileName = 'Activitylist %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    # new_activities.to_excel(FileName, index=False)

    # Debug: Write the old, stored activities to an Excel file
    # FileName = 'AllData %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    # all_activities.to_excel(FileName)

    # Write the new and the old activities to an Excel file
    FileName = 'FinData %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    fin_activities.to_excel(FileName)

    # Debug: Write the new and the old activities to the pickle file for store for next run
    fin_activities.to_pickle("ClubData.pkl")

# Run the main() function only when this file is called as main.
if __name__ == "__main__":
    main()