#########################################################
# Todo:
## 1
## 2
# Identifiser og flagg uvanlige aktiviteter pr type
#  * Mangler Crop
#  * Fjern aktiviteter som er fjernet fra activities, om det finnes en annen med samme bruker+dato + type/navn
## 3
# Legg opp til at config inneholder flere tokens, 
#  * les fra alle tokens (Alle må følge ASA)
#  * Legge inn info om hvem som har lest aktiviteten.
#  * Sjekk om noen har id som ikke andre har.
## 4
# Lag trekningsliste i Excel for forrige uke hver gang man starter på ny uke
#   * Nummerer aktivitetene som er mer enn 900 sekunder pr medlem
#   * Gi alle aktiviteter med nummer>1 og <5 random nummer, laveste vinner (Manuell sjekk om en aktivitet har fått 2 pga navnebror)
#   * Luke ut de som ikke jobber i Atea Norge
## 5
# Sjekk mot medlemslisten hvem som har like navn hver mandag
# Lag fil med run-statistics
#   Lag oversikt over antall medlemmer i klubben
#   Hvor mange aktiviteter som var nye side sist
#########################################################

import requests
import pandas as pd
import json
import errno
from datetime import datetime
from datetime import timedelta
import urllib3

# Disable warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Get access token based on client_id, client_secret and refresh_token
def authenticate(client_id, client_secret, refresh_token):
    url = "https://www.strava.com/oauth/token"

    data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'refresh_token': refresh_token,
        'grant_type': "refresh_token",
        'f':'json'
        }

    return requests.post(url, data=data, verify=False).json()['access_token']

# Create activities for admin user - placeholder for date at the end of each day
def create_date_activities(access_token,access_token_write):
    activities_url = "https://www.strava.com/api/v3/athlete/activities"
    header = {'Authorization': 'Bearer ' + access_token}
    param  = {'per_page': 10, 'page': 1}
    response = requests.get(activities_url,headers=header, params=param)
    data = response.json()        

    # Set found_date 7 days back
    found_date = datetime.now()- timedelta(days=7)
    found_date = found_date.replace(hour=0, minute=0, second=0, microsecond=0)

    # Find last date registered in Strava
    for line in data :
        taglist = line['name'].split("#")
        if len(taglist)==2:
            if taglist[1] == "AteaClubStats_Date":
                found_date = datetime.strptime(taglist[0], "%Y-%m-%d")
                break
    # Set newest_date to yesterday @ 23:59 (The last date to write)
    newest_date = datetime.now() - timedelta(days=1)
    newest_date = newest_date.replace(hour=23, minute=59, second=0, microsecond=0)

    # Set write_date to next date to write
    write_date = found_date + timedelta(days=1, hours=23, minutes=59)
    
    url = "https://www.strava.com/api/v3/activities"
    header = {'Authorization': 'Bearer ' + access_token_write}

    # Loop to create all date-activities       
    while newest_date >= write_date:
        activity_name = '%s#AteaClubStats_Date' % write_date.strftime("%Y-%m-%d")
        strava_date = write_date.strftime("%Y-%m-%dT23:59") # ISO 8601 formatted date time

        data = {
            'name': activity_name,
            'type': "run",
            'start_date_local': strava_date,
            'elapsed_time': 1,
            'distance': 0
            }
        
        # Create the placeholder-activity
        response = requests.post(url=url, headers=header, data=data)
        
        # Increase the date with 1 day
        write_date = write_date + timedelta(days=1)

# Read all stored activities from file. Create file if it doesn't exist
def read_activities_from_file(file_name,activities):
    try:
        activities = pd.read_pickle(file_name)
    except OSError as e:
        if e.errno == errno.ENOENT:
            activities.to_pickle(file_name)
        else:
            raise

    return activities


# Get all new activities from Strava API
def get_new_activities_from_strava(access_token,club_id,activities):
    loop = True

    readpage = 1
    pagesize = 50
    url = "https://www.strava.com/api/v3/clubs/%s/activities" % club_id
    counter = 0
    activity_date = datetime.now()
    
    while loop:
        header = {'Authorization': 'Bearer ' + access_token}
        param  = {'per_page': pagesize, 'page': readpage}
        response = requests.get(url, headers=header, params=param )
        data = response.json()

        readpage = readpage + 1
      
        for line in data :
            # There are no date in the data, so manual activities are created as placeholders, date will change when these activites are found
            taglist = line['name'].split("#")
            if len(taglist)==2:
                if taglist[1] == "AteaClubStats_Date":
                    activity_date = datetime.strptime(taglist[0], "%Y-%m-%d")
                    continue
                
            # Assign values to dataframe
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
        
        if len(data)<pagesize :
            loop = False
    
    print("activities found: %i" % (counter-1) )
    return activities

# Check "old" activites, replace when the same activity exist in "new" list. Append all that does not exist.
def remove_duplicate_activities(stored_activities,new_activities):
    # appen new activities backwards and reset index
    stored_activities = stored_activities.append(new_activities.iloc[::-1])
    # stored_activities.reset_index(drop=True, inplace=True)
    stored_activities.drop_duplicates(subset=['id'], keep='last', inplace=True, ignore_index=True)

    return stored_activities

def main():
    # Read config.json file
    with open("config.json", "r") as jsonfile:
        config = json.load(jsonfile)

    # Get an access token to authenticate when getting data from Strava
    access_token = authenticate(config["clients"][0]["client_id"],config["clients"][0]["client_secret"],config["clients"][0]["refresh_token"])

    # Get an access token to authenticate when writing data to Strava
    access_token_write = authenticate(config["clients"][0]["client_id"],config["clients"][0]["client_secret"],config["clients"][0]["refresh_token_write"])

    #Create manual activities to determine date on activity
    create_date_activities(access_token,access_token_write) 

    # Define data columns
    columns = [ "Athlete", "Name", "Type", "Duration", "Distance", "Date", "id" ]
    
    # Get stored data
    stored_activities = pd.DataFrame(columns=columns)
    stored_activities = read_activities_from_file("ClubData.pkl", stored_activities)

    # Get data from Strava
    new_activities = pd.DataFrame(columns=columns)
    new_activities = get_new_activities_from_strava(access_token, config["club_id"], new_activities)
    new_activities.set_index('id')

    # Add the new activites to the data already stored, but skip existing activities
    all_activities = pd.DataFrame(columns=columns)
    all_activities = remove_duplicate_activities(stored_activities, new_activities)

    # Debug: Write the new activities to an Excel file
    FileName = 'ClubData %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")

    # Write the dataset to file
    all_activities.to_pickle("ClubData.pkl")

    # all_activities.drop(inplace=True)

    all_activities.to_excel(FileName)

# Run the main() function only when this file is called as main.
if __name__ == "__main__":
    main()