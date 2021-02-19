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


# Read dataframe from Excel. Create file if it doesn't exist
def read_df_from_excel(file_name,df):
    try:
        df = pd.read_excel(file_name + ".xlsx")
    except OSError as e:
        if e.errno == errno.ENOENT: # No such file or directory, create new file
            df.to_excel(file_name + ".xlsx", index=False)
        else:
            raise
    return df

# Write dataframe to Excel. Change name if error
def write_df_to_excel(file_name,df):
    try:
        df.to_excel('%s.xlsx' % file_name, index=False)
    except OSError as e:
        if e.errno == errno.EACCES: # Permission denied: File already open
            file_name2 = '%s %s.xlsx' % (file_name, datetime.now().strftime("%Y.%m.%d %H%M"))
            df.to_excel(file_name2, index=False)
            print('Permission denied when saving file. File saved as: "%s". Please rename to "%s.xlsx"' % (file_name2, file_name))
        else:
            raise


# Get all new activities from Strava API
def get_members_from_club(access_token,club_id,members):
    loop = True

    readpage = 1
    pagesize = 50
    url = "https://www.strava.com/api/v3/clubs/%s/members" % club_id
    counter = 0
    activity_date = datetime.now()
    
    while loop:
        header = {'Authorization': 'Bearer ' + access_token}
        param  = {'per_page': pagesize, 'page': readpage}
        response = requests.get(url, headers=header, params=param )
        data = response.json()

        readpage = readpage + 1
      
        for line in data :
            # Assign values to dataframe
            members.at[counter, 'Firstname']    = line['firstname'] 
            members.at[counter, 'Lastname']     = line['lastname']
            members.at[counter, 'Membership']   = line['membership']
            members.at[counter, 'Owner']        = line['owner']
            # members.at[counter, 'Organization'] = line['type']
            
            counter = counter + 1
        
        if len(data)<pagesize :
            loop = False
    
    return members

# Check "old" activites, replace when the same activity exist in "new" list. Append all that does not exist.
def remove_duplicate_members(old_df,new_df):
    # Append new rows backwards and reset index
    stored_activities = stored_activities.append(new_activities.iloc[::-1])
    stored_activities.drop_duplicates(subset=['Firstname','Lastname'], keep='last', inplace=True, ignore_index=True)

    return stored_activities

def main():
    # Read config.json file
    with open("config.json", "r") as jsonfile:
        config = json.load(jsonfile)

    # Get an access token to authenticate when getting data from Strava
    access_token = authenticate(config["clients"][0]["client_id"],config["clients"][0]["client_secret"],config["clients"][0]["refresh_token"])

    # Define data columns
    member_columns = [ "Firstname", "Lastname", "Membership", "Admin", "Owner", "Organization" ]   

    # Get stored data
    member_file_name = 'ClubMembers %s' % config["club_id"]
    api_members = pd.DataFrame(columns=member_columns)
    api_members = get_members_from_club(access_token, config["club_id"], api_members)
    print("Member count: %i" % len(api_members))

    # Write the dataset to file
    write_df_to_excel(member_file_name, api_members)

# Run the main() function only when this file is called as main.
if __name__ == "__main__":
    main()