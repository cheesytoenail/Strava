$URL = "https://www.strava.com/api/v3/activities?access_token=" + $Key + "&per_page=200"

[String]$Key = "APIKey"

$Results = Invoke-RestMethod -Method Get -Uri $URL


