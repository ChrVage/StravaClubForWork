import requests
import urllib3 #trengs denne?
import pandas as pd
import json
import errno
from datetime import datetime
from datetime import timedelta

# Dette må gjøres før programmet når v1.0:

# Legg opp til at config inneholder flere tokens, les fra alle (Alle må følge ASA)

# Legge inn info om hvem som har lest aktiviteten.
# Sjekk om noen har id som ikke andre har.

# Les ut med access token fra både CV og ASA - for å se forskjeller... Dette i test-kode

# Finn og fjern de siste aktivitetene i den gamle filen som finnes i den nye (basert på id)

# Legg inn manuelle aktiviteter i Atea Strava Admin etter dato i config.json fil (eller kanskje faktiske innlagte dato?)

# Lag trekningsliste i Excel for forrige uke hver gang man starter på ny uke
#   * Nummerer aktivitetene som er mer enn 900 sekunder pr medlem
#   * Gi alle aktiviteter med nummer>1 og <5 random nummer, laveste vinner (Manuell sjekk om en aktivitet har fått 2 pga navnebror)

# Sjekk mot medlemslisten hvem som har like navn hver mandag
# Lag fil med run-statistics
#   Lag oversikt over antall medlemmer i klubben
#   Hvor mange aktiviteter som var nye side sist


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

# Create placeholders for date by creating manual acitivities at the end of each day
def create_date_activities(access_token):
    activities_url = "https://www.strava.com/api/v3/athlete/activities"
    header = {'Authorization': 'Bearer ' + access_token}
    param  = {'per_page': 10, 'page': 1}
    response = requests.get(activities_url,headers=header, params=param)
    data = response.json()        

    found_date = datetime.now()

    # Find last date registered
    for line in data :
        # Check for new date in activities
        taglist = line['name'].split("#")
        if len(taglist)==2:
            if taglist[1] == "AteaClubStats_Date":
                found_date = datetime.strptime(taglist[0], "%Y-%m-%d")
                break
    
    # If no date is found, create last 7 days.
    if found_date>=datetime.now() :
        found_date = found_date - timedelta(days=7)
        found_date = found_date.replace(hour=0, minute=0, second=0, microsecond=0)


    newest_date = datetime.now() - timedelta(days=1)
    newest_date = newest_date.replace(hour=23, minute=59, second=0, microsecond=0)
    write_date = found_date + timedelta(days=1, hours=23, minutes=59)
    
    # Loop to create all date-activities       
    while newest_date >= write_date:
        Activity_Name = '%s#AteaClubStats_Date' % write_date.strftime("%Y-%m-%d")
        print("Activity_Name:  %s" % Activity_Name)
        print("write_date:     %s" % write_date)
        write_date = write_date + timedelta(days=1)


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

    #Create manual activities
    create_date_activities(access_token) #access_token must be write for ASA

    # Define data columns
    columns = [ "Athlete", "Name", "Type", "Duration", "Distance", "Date", "id" ]
    
    # Get stored data
    all_activities = pd.DataFrame(columns=columns)
    all_activities = read_activities_from_file("ClubData.pkl", all_activities)

    # Get data from Strava
    new_activities = pd.DataFrame(columns=columns)
    new_activities = get_new_activities_from_strava(access_token, config["club_id"], new_activities)
    new_activities.set_index('id')

    # Add the new activites to the data already stored, but skip existing activities
    fin_activities = pd.DataFrame(columns=columns)
    fin_activities = remove_duplicate_activities(all_activities, new_activities)

    # Debug: Write the new activities to an Excel file
    FileName = 'Activitylist %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    new_activities.to_excel(FileName, index=False)

    # Debug: Write the old, stored activities to an Excel file
    FileName = 'AllData %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    all_activities.to_excel(FileName)

    # Write the new and the old activities to an Excel file
    FileName = 'FinData %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    fin_activities.to_excel(FileName)

    # Debug: Write the new and the old activities to the pickle file for store for next run
    fin_activities.to_pickle("ClubData.pkl")

# Run the main() function only when this file is called as main.
if __name__ == "__main__":
    main()