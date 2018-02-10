# SetStravaCommute
Fairly basic Google Script that can be easily reworked for different GPS coordinates.

This will return the default number of Strava activities (30 at the time of writing 10/02/2018) as JSON and then run a series of tests to identify if the bike ride is likely to be a "commute".

## Commute tests
The following tests are used to confirm if this is likely to be a commute
- Check if the start or end location is in the range of GPS co-rdinates defined at the beginning of the script (this works to within 2 decimal places)
- Check if the activity is a "Ride" type
- Check if the Commute field is already set to 'true'
- Check if the activity is using any of the default names for an activity

## Known Issues
Because the script is checking a large number of activities frequently there's a chance if you were to begin an activity at your "Work" location it may set your activity name and the 'commute' field repeatedly, as a workaround just rename the activity to anything other than the default Strava activity names