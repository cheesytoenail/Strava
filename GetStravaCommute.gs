function GetStravaCommute() {
    // Pre-Req function 
    // https://stackoverflow.com/questions/21131224/sorting-json-object-based-on-attribute
    function sortByKey(array, key) {
        return array.sort(function(a, b) {
        var x = a[key]
        var y = b[key]
        return ((x < y) ? -1 : ((x > y) ? 1 : 0))
            }
        )
    }
    
    // Google Sheet vars
    var ss = SpreadsheetApp.getActiveSpreadsheet() ; // Get Active Doc
    
    // Check for summary sheet
    if (ss.getSheetByName("Summary") == null) {
        ss.insertSheet().setName("Summary")
        var sheet = ss.getSheetByName(startYear)
    } else {
        var sheet = ss.getSheetByName(startYear)
    }
    
    // Summary Sheet
    var sheetSum = ss.getSheetByName("Summary")
    ss.setActiveSheet(sheetSum)
    ss.moveActiveSheet(0)
    
    // Check for default sheet and remove if present
    var sheetDef = ss.getSheetByName("Sheet1")
    if (sheetDef) {
        ss.setActiveSheet(sheetDef)
        ss.deleteActiveSheet()
    }
    
    // Tokens
    var readToken = "Bearer "

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
    responseGetObj = sortByKey(responseGetObj, 'id')

    // Iterate through JSON/Test/Update activity
    for (var index in responseGetObj) {

        // Test for "Commute" boolean
        if (responseGetObj[index].commute !== null) {
            var actComTest = responseGetObj[index].commute == true
        } else {
            var actComTest = false
        }

        // If tests true
        if (actComTest) {
            // Process the data
            var id = responseGetObj[index].id
            var timeDate = responseGetObj[index].start_date_local.split('T')
            var startDate = timeDate[0]
            var startTime = timeDate[1].slice(0,-4)
            var timeElapsed = ((responseGetObj[index].elapsed_time / 60) / 1440)
            var timeMoving = ((responseGetObj[index].moving_time / 60) / 1440)
            var actDist = (responseGetObj[index].distance / 1000).toFixed(1)
            var speedAverage = (responseGetObj[index].average_speed * 3.6).toFixed(1)
            var startYear = startDate.split('-')[0]

            // Costs calculations
            var columnSum = 1
            var columnValueSum = sheetSum.getRange(2, columnSum, sheetSum.getLastRow()).getValues()
            columnValueSum = String(columnValueSum).split(',')
            var startYearString = String(startYear)
            var searchResultSum = columnValueSum.indexOf(startYearString)

            if (searchResultSum == null) {
                var emptyASum = sheetSum.getRange(sheetSum.getLastRow()+1,1,1)
                var emptyBSum = sheetSum.getRange(sheetSum.getLastRow()+1,2,1)
                var emptyCSum = sheetSum.getRange(sheetSum.getLastRow()+1,3,1)
                var emptyDSum = sheetSum.getRange(sheetSum.getLastRow()+1,4,1)
                var emptyESum = sheetSum.getRange(sheetSum.getLastRow()+1,5,1)
                var emptyFSum = sheetSum.getRange(sheetSum.getLastRow()+1,6,1)
                var emptyGSum = sheetSum.getRange(sheetSum.getLastRow()+1,7,1)

                var totalDistFormSum = "SUM('" + startYear + "'!I:I)"
                var totalCostFormSum = "SUM('" + startYear + "'!J:J)"

                emptyASum.setValue([startYear]).setFontWeight("bold").setHorizontalAlignment("right")
                emptyBSum.setValue("Single Fare").setFontWeight("bold").setHorizontalAlignment("right")
                emptyCSum.setHorizontalAlignment("right")
                emptyDSum.setValue("Total").setFontWeight("bold").setHorizontalAlignment("right")
                emptyESum.setFormula([totalDistFormSum]).setHorizontalAlignment("right")
                emptyFSum.setValue("Distance").setFontWeight("bold").setHorizontalAlignment("right")
                emptyGSum.setFormula([totalCostFormSum]).setHorizontalAlignment("right")
            }

            // Set column headers
            if (ss.getSheetByName(startYear) == null) {
                ss.insertSheet().setName(startYear)
                var sheet = ss.getSheetByName(startYear)
                sheet.getRange('A1').setValue('Activity ID').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('B1').setValue('Date').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('C1').setValue('Time').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('D1').setValue('Total Time').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('E1').setValue('Active Time').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('F1').setValue('Distance').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('G1').setValue('Speed').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('H1').setValue('NP').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('I1').setValue('Total Distance').setFontWeight("bold").setHorizontalAlignment("right")
                sheet.getRange('J1').setValue('Total Savings').setFontWeight("bold").setHorizontalAlignment("right")
            } else {
                var sheet = ss.getSheetByName(startYear)
            }

            // Normalised power
            if (responseGetObj[index].hasOwnProperty('weighted_average_watts')) {
                var powerNP = responseGetObj[index].weighted_average_watts
            } else {
                var powerNP = "N/A"
            }

            // Empty cells
            var emptyA = sheet.getRange(sheet.getLastRow()+1,1,1)
            var emptyB = sheet.getRange(sheet.getLastRow()+1,2,1)
            var emptyC = sheet.getRange(sheet.getLastRow()+1,3,1)
            var emptyD = sheet.getRange(sheet.getLastRow()+1,4,1)
            var emptyE = sheet.getRange(sheet.getLastRow()+1,5,1)
            var emptyF = sheet.getRange(sheet.getLastRow()+1,6,1)
            var emptyG = sheet.getRange(sheet.getLastRow()+1,7,1)
            var emptyH = sheet.getRange(sheet.getLastRow()+1,8,1)
            var emptyH = sheet.getRange(sheet.getLastRow()+1,9,1)
            
            // Check for duplicates
            var column = 1
            var columnValues = sheet.getRange(2, column, sheet.getLastRow()).getValues()
            columnValues = String(columnValues).split(',')
            var idString = String(id)
            var searchResultAct = columnValues.indexOf(idString)

            // Assign Vars
            if (searchResultAct == -1) {
                emptyA.setValue([id]).setHorizontalAlignment("right")
                emptyB.setValue([startDate]).setNumberFormat(("dddd, dd/mm/yy")).setHorizontalAlignment("right")
                emptyC.setValue([startTime]).setHorizontalAlignment("right")
                emptyD.setValue([timeElapsed]).setNumberFormat("h:mm:ss").setHorizontalAlignment("right")
                emptyE.setValue([timeMoving]).setNumberFormat("h:mm:ss").setHorizontalAlignment("right")
                emptyF.setValue([actDist]).setNumberFormat("0.0").setHorizontalAlignment("right")
                emptyG.setValue([speedAverage]).setNumberFormat("0.0").setHorizontalAlignment("right")
                emptyH.setValue([powerNP]).setHorizontalAlignment("right")
            }

            // Column resizing  
            sheet.autoResizeColumn(1)
            sheet.autoResizeColumn(2)
            sheet.autoResizeColumn(3)
            sheet.autoResizeColumn(4)
            sheet.autoResizeColumn(5)
            sheet.autoResizeColumn(6)
            sheet.autoResizeColumn(7)
            sheet.autoResizeColumn(8)
        }
    }

    // Log PUT writes (may be useful for identifying Oauth issues)
    //Logger.log();
}
