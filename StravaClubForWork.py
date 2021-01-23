import requests
import urllib3
import pandas as pd 

auth_url = "https://www.strava.com/oauth/token"
activities_url = "https://www.strava.com/api/v3/clubs/10971/activities"
athlete_activities_url = "https://www.strava.com/api/v3/athlete/activities" 
getClubMembersById_url = "https://www.strava.com/api/v3/clubs/10971/members"

payload = {
    'client_id': "60365",
    'client_secret': 'd6eaed2e3d1f1d8a154044d5984769536dc06674',
    'refresh_token': '295e4ffa5db843b8482fcdbd2ca9a5ae2c1a24a7',
    'grant_type': "refresh_token",
    'f':'json'
}

access_token = requests.post(auth_url,data=payload, verify=False).json()['access_token']
print("Access token:{0}\n",access_token)



#https://www.strava.com/athletes/12664466
header = {'Authorization': 'Bearer ' + access_token}
param = {'per_page': 10, 'page': 1}
#param = {'per_page': 10, 'page': 1, 'before':'', 'after':''}
my_dataset = requests.get(activities_url,headers=header, params=param).json()
#my_dataset = requests.get(getClubMembersById_url,headers=header).json()


firstname = my_dataset[0]['athlete']['firstname']
lastname = my_dataset[0]['athlete']['lastname']
name = my_dataset[0]['name']
TheData = list(zip(firstname,lastname,name));
print(TheData)

df = pd.DataFrame(data = TheData, columns=['Firstname', 'Lastname', 'Name'])
df.index.name = 'Index'
df.to_csv('Activitylist.txt',index=True,header=True)
