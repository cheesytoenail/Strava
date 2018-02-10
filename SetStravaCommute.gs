function SetStravaCommute() {
    // Office GPS Values (https://www.geoplaner.com/)
    var maxLat = 51.51
    var minLat = 51.50
    var maxLng = -0.42
    var minLng = -0.43

    // Tokens
    var readToken = "Bearer ########################################"
    var writeToken = "Bearer ########################################" // http://yizeng.me/2017/01/11/get-a-strava-api-access-token-with-write-permission/

    // Strava API
    var apiUrl = "https://www.strava.com/api/v3"
    var apiMethAth = "/athlete"
    var apiMethAct = "/activities"

    // Make GET Request and convert JSON
    var requestGetUrl = apiUrl + apiMethAth + apiMethAct
    var paramsGet = {
        "method": "GET",
        "headers": {
            "Authorization": readToken
        }
    }
    var responseGetJson = UrlFetchApp.fetch(requestGetUrl, paramsGet)
    var responseGetObj = JSON.parse(responseGetJson)

    // Iterate through JSON/Test/Update activity
    for (var index in responseGetObj) {

        // Test start GPS for if value present and then test if co-ordinates are correct
        if (responseGetObj[index].start_latlng !== null) {
            var maxLatStart = responseGetObj[index].start_latlng[0] <= maxLat
            var minLatStart = responseGetObj[index].start_latlng[0] >= minLat
            var minLngStart = responseGetObj[index].start_latlng[1] <= maxLng
            var maxLngStart = responseGetObj[index].start_latlng[1] >= minLng
        } else {
            var minLatStart = false
            var maxLatStart = false
            var minLngStart = false
            var maxLngStart = false
        }

        // Test end GPS for if value present and then test if co-ordinates are correct
        if (responseGetObj[index].end_latlng !== null) {
            var maxLatEnd = responseGetObj[index].end_latlng[0] <= maxLat
            var minLatEnd = responseGetObj[index].end_latlng[0] >= minLat
            var minLngEnd = responseGetObj[index].end_latlng[1] <= maxLng
            var maxLngEnd = responseGetObj[index].end_latlng[1] >= minLng
        } else {
            var minLatEnd = false
            var maxLatEnd = false
            var minLngEnd = false
            var maxLngEnd = false
        }

        // Test for "Ride" value
        if (responseGetObj[index].type !== null) {
            var actTypeTest = responseGetObj[index].type == "Ride"
        } else {
            var actTypeTest = false
        }

        // Test for "Commute" boolean
        if (responseGetObj[index].commute !== null) {
            var actComTest = responseGetObj[index].commute == true
        } else {
            var actComTest = false
        }

        // Test for default activity name
        // (we can use this to manually override scipt by renaming activity to anything else)
        var actNameTest = false
        if (responseGetObj[index].name !== null) {
            switch (responseGetObj[index].name) {
                case "Morning Ride":
                    var actNameTest = true
                case "Afternoon Ride":
                    var actNameTest = true
                case "Evening Ride":
                    var actNameTest = true
            }
        }

        // Group the tests together
        var actStartTest = minLatStart && maxLatStart && minLngStart && maxLngStart
        var actEndtest = minLatEnd && maxLatEnd && minLngEnd && maxLngEnd

        // If either GPS tests true and type test true, update activity
        if ((actStartTest || actEndtest) && actTypeTest && !(actComTest) && actNameTest) {

            // PUT vars
            var actId = "/" + responseGetObj[index].id
            var actPut = responseGetObj[index]
            actPut.commute = true
            actPut.name = "Commute"
            actPut = JSON.stringify(actPut)

            // Make PUT Request and update activity
            var requestPutUrl = apiUrl + apiMethAct + actId
            var paramsPut = {
                "method": "PUT",
                "headers": {
                    "Authorization": writeToken
                },
                "payload": actPut,
                "contentType": "application/json"
            }
            var responsePutJson = UrlFetchApp.fetch(requestPutUrl, paramsPut)
        }
    }

    // Log PUT writes (may be useful for identifying Oauth issues)
    Logger.log(responsePutJson);
}
