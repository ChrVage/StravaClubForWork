#########################################################
# Todo:
#   -Sjekk newline - verdi str/int/true/false
#   -Legg til header på statistikken.
#   -Kopier viktige filer til en fornuftig plass (Lagre f.eks. medlemmer pr dato)
#   -Lag system for Sykle til jobben
#       -Forms - registrering av Strava-brukernavn og epost.  
#   -Fjern aktiviteter som er fjernet fra activities, om det finnes en annen med samme bruker+dato + type/navn
#   -Sjekk mot medlemslisten hvem som har like navn
#   -Rydd i datetime - håndtering.
#   -lag kode som tester forskjellige tokens om det gir forskjellig resultat
#       -les fra flere tokens om noen mangler
#   -Sjekk om disse mangler ennå fra ASA (Olav L, Glenn Ø, Victoria R, Anett V, Christian V)
#########################################################

import requests
import pandas as pd
import json
import errno
from datetime import datetime, timedelta, time
import urllib3
import os
import numpy as np
from stat import S_IREAD, S_IRGRP, S_IROTH
import locale
from math import floor
locale.setlocale(locale.LC_ALL, '')

# Disable warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
pd.options.mode.chained_assignment = None  # default='warn'

# Define global variables
data_columns = [ "Athlete", "Name", "Distance", "Moving time", "Elapsed time", "Elevation gain", "Type", "Workout type", "Date", "id", "Duration (s)", "Duration" ]
stat_columns = [ "Timestamp", "Execution time (sec)", "Since last run (hrs)", "Stored activities", "API activities", "Appended", "Appended/New" ]

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

# Read dataframe from Excel. Create file if it doesn't exist
def read_df_from_excel(file_name,df):
    try:
        df = pd.read_excel(file_name)
    except OSError as e:
        if e.errno == errno.ENOENT: # No such file or directory, create new file
            df.to_excel(file_name, index=False)
        else:
            raise
    return df

# Write dataframe to Excel. Change name if error
def write_df_to_excel(file_name,df):
    try:
        df.to_excel('%s' % file_name, index=False)
    except OSError as e:
        if e.errno == errno.EACCES: # Permission denied: File already open
            file_name2 = '%s %s.xlsx' % (file_name, datetime.now().strftime("%Y.%m.%d %H%M"))
            df.to_excel(file_name2, index=False)
            print('Permission denied when saving file. File saved as: "%s". Please rename to "%s"' % (file_name2, file_name))
        else:
            raise

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
      
        # Loop the json-file
        for line in data :
            # There are no date in the data, so manual activities are created as placeholders, date will change when these activites are found
            taglist = line['name'].split("#")
            if len(taglist)==2:
                if taglist[1] == "AteaClubStats_Date":
                    activity_date = datetime.strptime(taglist[0], "%Y-%m-%d")
                    continue

            athlete         = line['athlete']['firstname'] +"#"+ line['athlete']['lastname']
            seconds_elapsed = line['elapsed_time']
            meters          = line['distance']

            duration_s = seconds_elapsed

            # If duration is over 45 minutes and speed is lower than 1 m/s, limit to 45 min and at least 1 m/s
            if seconds_elapsed>2700 and seconds_elapsed>meters:
                duration_s = max(2700,meters)

            duration = datetime(year=1,month=1,day=1) + timedelta(seconds=duration_s)
            duration = duration.time()
            
            # Assign values to dataframe
            activities.at[counter, 'Athlete']       = athlete
            activities.at[counter, 'Name']          = line['name']
            activities.at[counter, 'Distance']      = meters
            activities.at[counter, 'Moving time']   = line['moving_time']
            activities.at[counter, 'Elapsed time']  = seconds_elapsed
            activities.at[counter, 'Elevation gain']= line['total_elevation_gain']
            activities.at[counter, 'Type']          = line['type']
            activities.at[counter, 'Workout type']  = line.get('workout_type','na') # Sometimes, this value isn't available
            activities.at[counter, 'Date']          = activity_date.replace(hour=0, minute=0, second=0, microsecond=0)
            activities.at[counter, 'id']            = "%s#%s#%s#%s" % ( athlete, 
                                                                        seconds_elapsed, 
                                                                        meters,
                                                                        activity_date.strftime("%Y-%m-%d"))
            activities.at[counter, 'Duration (s)']  = duration_s
            activities.at[counter, 'Duration']      = duration.strftime("%H:%M:%S")
            
            counter = counter + 1
        
        if len(data)<pagesize :
            loop = False
    
    return activities

# Check "old" activites, replace when the same activity exist in "new" list. Append all that does not exist.
def remove_duplicate_activities(stored_activities,new_activities):
    # Append new activities backwards and reset index
    stored_activities = stored_activities.append(new_activities.iloc[::-1])
    stored_activities.drop_duplicates(subset=['id'], keep='last', inplace=True, ignore_index=True)

    return stored_activities

def create_subset(df,exclude_athletes):
    # Read table to check which subsets to create
    config_subset = pd.DataFrame()
    config_subset = read_df_from_excel("config_subset.xlsx", config_subset )

    # Set end date to yesterday to check for subsets that ended 
    end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1) 

    # Select lines from config_subset that matches 
    index_list = config_subset.index[config_subset['End date']==end_date].tolist()

    if len(index_list)==0: # No subset to create on this date
        return

    for index in index_list:
        start_date = config_subset.at[index, 'Start date'] 
        setup =     config_subset.at[index, 'Setup'] 
        file_name = config_subset.at[index, 'Filename']
        new_line = config_subset.at[index, 'Newline']

        # Find rows in the specified date interval
        subset_df = df.loc[(df['Date']>=start_date) & (df['Date']<=end_date)]

        if setup=="Trekning":
            # Create draw-column with random numbers
            no_count = 10000
            subset_df[setup] = np.random.randint(1, no_count, subset_df.shape[0])
            # Disqualify activities that are too short
            subset_df.loc[subset_df['Elapsed time'] < 900, setup] = no_count
            subset_df.loc[subset_df['Elapsed time'] < 900, 'Kommentar'] = "For kort økt"
            # Sort dataset with the winner on top
            subset_df.sort_values(by=[setup], inplace=True)
            if new_line:
                new_start_date = start_date + timedelta(days=7)
                new_file_name  = "Trekning uke %s.xlsx" % new_start_date.strftime("%V-%Y")

                config_subset = config_subset.append(  {'End date': end_date + timedelta(days=7),
                                                        'Start date': new_start_date, 
                                                        'Setup': setup, 
                                                        'Filename': new_file_name, 
                                                        'Newline': True}, 
                                                        ignore_index=True)

        if setup=="Minutter":
            no_count = 0
            # Set actual number of minutes 
            subset_df[setup]         = subset_df['Duration (s)']/60

        # Remove activities from people not working in Atea
        subset_df.loc[subset_df.Athlete.isin(exclude_athletes), setup] = no_count
        subset_df.loc[subset_df.Athlete.isin(exclude_athletes), 'Kommentar'] = "Jobber ikke i Atea Norge"

        subset_stat = subset_df[['Type']].drop_duplicates().reset_index(drop=True)

        # Create stats pr activity type
        index_stat_list = subset_stat.index.tolist()
 
        for index_stat in index_stat_list:
            activity_type = subset_stat.at[index_stat, 'Type']
            subset_stat.at[index_stat,'Athlete']        = len(subset_df['Athlete'].loc[subset_df['Type']==activity_type].drop_duplicates())
            subset_stat.at[index_stat,'Name']           = len(subset_df.loc[subset_df['Type']==activity_type].drop_duplicates())
            subset_stat.at[index_stat,'Distance']       = '{0:n}'.format(subset_df['Distance'].loc[subset_df['Type']==activity_type].sum())         #1 000
            subset_stat.at[index_stat,'Moving time']    = str(timedelta(seconds=floor(subset_df['Moving time'].loc[subset_df['Type']==activity_type].sum())))      #h:m:s
            subset_stat.at[index_stat,'Elapsed time']   = str(timedelta(seconds=floor(subset_df['Elapsed time'].loc[subset_df['Type']==activity_type].sum())))     #h:m:s
            subset_stat.at[index_stat,'Elevation gain'] = '{0:n}'.format(subset_df['Elevation gain'].loc[subset_df['Type']==activity_type].sum())   #1 000
            subset_stat.at[index_stat,'Duration (s)']   = str(timedelta(seconds=floor(subset_df['Duration (s)'].loc[subset_df['Type']==activity_type].sum())))     #h:m:s
        
        subset_stat.at[index_stat+1,'Athlete']        = len(subset_df['Athlete'].drop_duplicates())
        subset_stat.at[index_stat+1,'Name']           = subset_stat['Name'].sum()
        subset_stat.at[index_stat+1,'Distance']       = '{0:n}'.format(subset_df['Distance'].sum())         #1 000
        subset_stat.at[index_stat+1,'Moving time']    = str(timedelta(seconds=floor(subset_df['Moving time'].sum())))      #h:m:s
        subset_stat.at[index_stat+1,'Elapsed time']   = str(timedelta(seconds=floor(subset_df['Elapsed time'].sum())))     #h:m:s
        subset_stat.at[index_stat+1,'Elevation gain'] = '{0:n}'.format(subset_df['Elevation gain'].sum())   #1 000
        subset_stat.at[index_stat+1,'Type']           = 'Totalt'
        subset_stat.at[index_stat+1,'Duration (s)']   = str(timedelta(seconds=floor(subset_df['Duration (s)'].sum())))     #h:m:s
        
        # Adding a blank row on top
        subset_stat.at[-1,'Athlete'] = ['']
        #subset_stat.at[-1,data_columns] = [data_columns] # Add headlines
        subset_stat.sort_index(inplace=True)

        # Append the stats on the end of the subset
        subset_df = subset_df.append(subset_stat, ignore_index=True)

        # Write the new subset to Excel
        write_df_to_excel(file_name, subset_df)
        print("Subset created: %s" % file_name)

    # Delete lines for created subsets
    config_subset.drop(index_list,inplace=True)
    # Update config_subset - file
    write_df_to_excel("config_subset.xlsx", config_subset)

def main():
    # Set start time for run statistics
    start_time = datetime.now()

    # Read config.json file
    with open("config.json", "r") as jsonfile:
        config = json.load(jsonfile)

    # Get an access token to authenticate when getting data from Strava
    #access_token = authenticate(config["clients"][0]["client_id"],config["clients"][0]["client_secret"],config["clients"][0]["refresh_token"])
    access_token = authenticate(config["clients"][1]["client_id"],config["clients"][1]["client_secret"],config["clients"][1]["refresh_token"])

    # Get an access token to authenticate when writing data to Strava
    access_token_write = authenticate(config["clients"][0]["client_id"],config["clients"][0]["client_secret"],config["clients"][0]["refresh_token_write"])

    #Create manual activities to determine date on activity
    create_date_activities(access_token,access_token_write) 
    
    # Get stored data
    data_file_name = 'ClubData %s.xlsx' % config["club_id"]
    stored_activities = pd.DataFrame(columns=data_columns)
    stored_activities = read_df_from_excel(data_file_name, stored_activities)
    print("Stored activities: %i" % len(stored_activities))

    # Get data from Strava
    api_activities = pd.DataFrame(columns=data_columns)
    api_activities = get_new_activities_from_strava(access_token, config["club_id"], api_activities)
    api_activities.set_index('id')
    print("Api activities:    %i" % len(api_activities))

    # Add the new activites to the data already stored, but skip existing activities
    all_activities = pd.DataFrame(columns=data_columns)
    all_activities = remove_duplicate_activities(stored_activities, api_activities)
    print("All activities:    %i, %i added" % (len(all_activities), len(all_activities)-len(stored_activities)))

    # Check for and create subsets of the data
    create_subset(stored_activities, config["exclude_athletes"])

    # Debug: Write the new activities to an Excel file
    file_name = 'ClubData %s.xlsx' % datetime.now().strftime("%Y.%m.%d %H%M")
    write_df_to_excel(file_name, all_activities)

    # Write the dataset to file
    write_df_to_excel(data_file_name, all_activities)
    
    # Write a copy to TP2B
    strava_data_file = 'C:/Users/ChrVage/Atea/NO-ATEA alle - Strava Data/StravaData.xlsx'
    write_df_to_excel(strava_data_file, all_activities)

    # Write run statistics
    run_statistics = pd.read_excel("RunStats.xlsx")
    data = [ datetime.now(), 
             (datetime.now() - start_time).total_seconds(), 
             (datetime.now() - run_statistics['Timestamp'].iloc[-1])*24, 
             len(stored_activities), 
             len(api_activities), 
             len(all_activities)-len(stored_activities), 
             (len(all_activities)-len(stored_activities))//len(api_activities)]
    this_run = pd.DataFrame([data], columns=stat_columns )
    run_statistics = run_statistics.append(this_run, ignore_index=True )
    write_df_to_excel('RunStats.xlsx', run_statistics)

    print("Last run: %s" % datetime.now().strftime("%Y-%m-%d %H:%M"))


# Run the main() function only when this file is called as main.
if __name__ == "__main__":
    main()