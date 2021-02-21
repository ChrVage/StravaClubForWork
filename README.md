# StravaClubStats
Collecting and storing activities from a Strava Club for statistics.

- Data:
  - Athlete name (First name and first letter of last name)
  - Activity name
  - Type of workout
  - Duration (seconds)
  - Distance (meters)
  - Activity date

* Authenticate against Strava (for reading and writing)
* Create placeholder activities to set date on the activities (date is not a part of the clubs-api)
* Read stored activity data from file
* Get new activities from Strava
* Append data and remove existing (duplicate) activities.
* Store new dataset to file
* Create subsets automatically each week 
