function groupBy(list, keyGetter) {
    const map = new Map();
    list.forEach((item) => {
         const key = keyGetter(item);
         const collection = map.get(key);
         if (!collection) {
             map.set(key, [item]);
         } else {
             collection.push(item);
         }
    });
    return map;
}

// Create toolbar dropdown items
function onOpen() {
  var ui = SpreadsheetApp.getUi();
 
  ui.createMenu('Strava App')
    .addItem('Sync Data', 'updateSheet')
    .addItem('Logout', 'logout')
    .addToUi();
}

const trackedColumns = {
  Mileage: "Mileage",
  Runs: "SSRuns",
  Date: "Date",
  Pace: "Pace",
  HeartRate: "HeartRate",
  LapData: "Splits"
}

const startRowIndex = 4;
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getActiveSheet();

function setCellData(rowIndex, colIndex, value) {
  var cell = sheet.getRange(rowIndex, colIndex, 1, 1)
  Logger.log(!cell.getValues()[0])
  // if (!cell.getValues()[0]) {
    cell.setValues([[value]])
  // }
}

function updateSheet() {
  var service = getStravaService();
   
  if (!service.hasAccess()) {
    return;
  }

  var columnHeaders = sheet.getRange(1, 1, 2, sheet.getLastColumn());
  var [colNameToIndex, indexToColName] = [...columnHeaders.getValues()[0], ...columnHeaders.getValues()[1]].reduce((acc, cell, index) => {
    var [tmpColNameToIndex, tmpIndexToColName] = acc;
    if (Object.values(trackedColumns).includes(cell)) {
      tmpColNameToIndex.set(cell, index % sheet.getLastColumn())
      tmpIndexToColName.set(index % sheet.getLastColumn(), cell)
    }

    return acc;
  }, [new Map(), new Map()])

  // Print colNameToIndex
  Logger.log(JSON.stringify(Array.from(colNameToIndex.entries())))
  // Print indexToColName
  Logger.log(JSON.stringify(Array.from(indexToColName.entries())))

  var logData = sheet.getRange(startRowIndex, 1, sheet.getLastRow() - (startRowIndex - 1), sheet.getLastColumn());
  var parsedDataMap = logData.getValues().reduce((acc, row, rowIndex) => {
    var parsedRow = row.reduce((acc, col, index) => {
      if (indexToColName.has(index)) {
        return {
          ...acc,
          [indexToColName.get(index)]: col
        }
      }
      return acc;
    }, {rowIndex: rowIndex + startRowIndex})

    if (Object.keys(parsedRow).includes(trackedColumns.Date) && !!parsedRow[trackedColumns.Date]) {
      var date = parsedRow[trackedColumns.Date];
      var dateStr = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, "0")}-${date.getDate().toString().padStart(2, "0")}`
      acc.set(dateStr, parsedRow);
    }

    return acc;
  }, new Map)
  Logger.log(JSON.stringify(Array.from(parsedDataMap.entries())))
  var parsedDataRows = Array.from(parsedDataMap.values())

  var firstRowIndexMissingData = parsedDataRows.findIndex(row => {
    return Object.values(row).some(col => !col)
  })
  var lastRowIndexMissingData = parsedDataRows.length - 1 - [...parsedDataRows].reverse().findIndex(row => Object.values(row).some(col => !col))
  Logger.log(firstRowIndexMissingData)
  Logger.log(lastRowIndexMissingData)
  var firstRowDate = new Date(parsedDataRows[firstRowIndexMissingData][trackedColumns.Date]);
  var lastRowDate = new Date(parsedDataRows[lastRowIndexMissingData][trackedColumns.Date]);

  Logger.log(JSON.stringify(parsedDataRows));

  var activities = getStravaActiviesInRange(service, firstRowDate, lastRowDate)

  Logger.log(JSON.stringify(activities));

  Logger.log(activities.map(item => item.start_date_local))

  var groupedActivities = groupBy(activities, activity => {
    var dateKey = activity.start_date_local.split("T")[0];

    return dateKey
  })
  Logger.log("Grouped: " + JSON.stringify(Array.from(groupedActivities.entries())[0]))
  Array.from(groupedActivities.entries()).forEach((entry, index) => {
    const [dateKey, activities] = entry
    Logger.log("Date Key: " + dateKey)
    Logger.log("Activities: " + activities)

    var aggregateData = activities.reduce((acc, activity) => {
      Logger.log("Activity: " + activity)
      var secondPerMeter = Math.pow(activity.average_speed, -1)
      var secondPerMile = secondPerMeter * 1609
      var pace = `${Math.trunc(secondPerMile / 60).toString()}:${Math.round(secondPerMile % 60).toString().padStart(2, "0")}`

      return {
        totalDistance: activity.distance + (acc.totalDistance ?? 0),
        runs: [...acc.runs ?? [], activity.distance],
        paces: [...acc.paces ?? [], pace],
        averageHeartRates: [...acc.averageHeartRates ?? [], activity.average_heartrate],
        lapDataList: [...(index <= 20 ? [getLapData(service, activity.id)] : []), ...acc.lapDataList ?? []],
        activityDataList: [...(index <= 20 ? [getActivityData(service, activity.id)] : []), ...acc.activityDataList ?? []]
      }
    }, {})

    Logger.log(aggregateData)

    Logger.log(dateKey);
    var sheetData = parsedDataMap.get(dateKey);
    // Logger.log("Activity: " + JSON.stringify(activity));
    Logger.log(sheetData.rowIndex);
    Logger.log(colNameToIndex.get(trackedColumns.Mileage) + 1);


    setCellData(sheetData.rowIndex, colNameToIndex.get(trackedColumns.Mileage) + 1, Math.round((aggregateData.totalDistance / 1609) * 100) / 100);
    setCellData(sheetData.rowIndex, colNameToIndex.get(trackedColumns.Runs) + 1, aggregateData.runs.map(run => Math.round((run / 1609) * 100) / 100).join(", "));
    setCellData(sheetData.rowIndex, colNameToIndex.get(trackedColumns.Pace) + 1, aggregateData.paces.join(", "));
    setCellData(sheetData.rowIndex, colNameToIndex.get(trackedColumns.HeartRate) + 1, aggregateData.averageHeartRates.join(", "));
    if (aggregateData.lapDataList?.length != 0) {
      Logger.log(aggregateData.lapData)
      setCellData(sheetData.rowIndex, colNameToIndex.get(trackedColumns.LapData) + 1, aggregateData.lapDataList.map(laps => laps.map(lap => {
        var secondPerMeter = Math.pow(lap.average_speed, -1)
        var secondPerMile = secondPerMeter * 1609
        var pace = `${Math.trunc(secondPerMile / 60).toString()}:${Math.round(secondPerMile % 60).toString().padStart(2, "0")}`

        Logger.log("Lap: " + JSON.stringify(lap))

        return `${lap.moving_time} (${pace})`
      }).join(", ")).join(" | "));
    }
    if (aggregateData.activityDataList?.length != 0) {
      aggregateData.activityDataList.forEach(activityData => {
        Logger.log("Activity Data: " + JSON.stringify(activityData))
      })
    }
  })

  // Logger.log(JSON.stringify(Array.from(columnIndexMap.entries())))

  // var columncolumnHeaders.getValues().find(cell => {
  //   Logger.log(cell);

  //   return true
  // });

  // var range = sheet.getRange(1, 2, sheet.getLastRow(), 1).getValues().map((item, index) => {
  //   var timestamp = Date.parse(item);

  //   if (isNaN(timestamp) == true) {
  //     return
  //   }

  //   var date = new Date(timestamp);
  //   return {
  //     key: `${date.getFullYear().toString()}-${(date.getMonth() + 1).toString().padStart(2, "0")}-${date.getDate().toString().padStart(2, "0")}`,
  //     rowIndex: index + 1
  //   }
  // }).filter(item => !!item)

  // range.forEach(item => {
  //   let runningData = data.get(item.key)

  //   if (!value) {
  //     return;
  //   }

  //   let distance = Math.round((runningData.distance / 1609) * 100) / 100
  //   let laps = runningData.laps.join("|")

  //   sheet.getRange(item.rowIndex, 13, 1, 1).setValues([[distance]])
  //   sheet.getRange(item.rowIndex, 14, 1, 1).setValues([[value]])
  // })

  // Logger.log(OAuth2.getServiceNames())
}

function getStravaActiviesInRange(service, fromDate, toDate) {
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = `?after=${fromDate.getTime() / 1000}&before=${toDate.getTime() / 1000}&per_page=100`;
 
    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
     
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
     
    return JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
}

function getActivityData(service, activityId) {
    var lapEndpoint = `https://www.strava.com/api/v3/activities/${activityId}`;

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    return JSON.parse(UrlFetchApp.fetch(lapEndpoint, options));
}

function getLapData(service, activityId) {
    var lapEndpoint = `https://www.strava.com/api/v3/activities/${activityId}/laps`;

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    return JSON.parse(UrlFetchApp.fetch(lapEndpoint, options));
}
 
// Get athlete activity data
function getStravaActivityData() {
   
  // get the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
 
  // call the Strava API to retrieve data
  var data = callStravaAPI();
   
  // empty array to hold activity data
  var stravaData = [];
     
  // loop over activity data and add to stravaData array for Sheet
  data.forEach(function(activity) {
    var arr = [];
    arr.push(
      activity.id,
      activity.name,
      activity.type,
      activity.distance,
      activity.start_date
    );
    stravaData.push(arr);
  });

  var range = sheet.getRange(1, 2, sheet.getLastRow(), 1).getValues().map((item, index) => {
    var timestamp = Date.parse(item);

    if (isNaN(timestamp) == true) {
      return
    }

    var date = new Date(timestamp);
    return {
      key: `${date.getFullYear().toString()}-${(date.getMonth() + 1).toString().padStart(2, "0")}-${date.getDate().toString().padStart(2, "0")}`,
      rowIndex: index + 1
    }
  }).filter(item => !!item)

  range.forEach(item => {
    let runningData = data.get(item.key)

    if (!value) {
      return;
    }

    let distance = Math.round((runningData.distance / 1609) * 100) / 100
    let laps = runningData.laps.join("|")

    sheet.getRange(item.rowIndex, 13, 1, 1).setValues([[distance]])
    sheet.getRange(item.rowIndex, 14, 1, 1).setValues([[value]])
  })
  
  Logger.log("App has no access yet.");

  // paste the values into the Sheet
  // sheet.getRange(sheet.getLastRow() + 1, 1, stravaData.length, stravaData[0].length).setValues(stravaData);
}
 
function logout(service) {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava App')
    .addItem('Logout', 'getStravaActivityData')
    .addToUi();

  var service = getStravaService();
  service.reset()
}

function showDialog() {

}

// call the Strava API
function callStravaAPI() {
   
  // set up the service
  var service = getStravaService();
   
  if (service.hasAccess()) {
    Logger.log('App has access.');
     
  var ui = SpreadsheetApp.getUi();
 
  ui.createMenu('Strava App')
    .addItem('Logout', 'getStravaActivityData')
    .addToUi();

    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = `?before=${new Date().getTime() / 1000}&per_page=200`;
 
    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
     
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
     
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));

    Logger.log(JSON.stringify(response))
    Logger.log("Activity Length: " + response.length)

    var map = response.reduce((acc, curr) => {
      var key = curr.start_date_local.split("T")[0]

      var lapEndpoint = `https://www.strava.com/api/v3/activities/${curr.id}/laps`;
  
      var headers = {
        Authorization: 'Bearer ' + service.getAccessToken()
      };
      
      var options = {
        headers: headers,
        method : 'GET',
        muteHttpExceptions: true
      };
      
      var response = JSON.parse(UrlFetchApp.fetch(lapEndpoint, options));

      var laps = response?.map((item) => {
        return {
          lapIndex: item.lap_index,
          elapsedTime: item.elapsed_time,
          distance: item.distance
        }
      }) ?? []
      var lapsStr = laps.map(lap => lap.elapsedTime).join(",")


      Logger.log(curr)

      if (acc.has(key)) {
        acc.set(key, {distance: acc.get(key).distance + curr.distance, laps: [acc.get(key).laps, lapsStr]})
      } else {
        acc.set(key, {distance: curr.distance, laps: [lapsStr]})
      }
      return acc
    }, new Map())

    return map;  
  }
  else {
    Logger.log("App has no access yet.");
     
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();
     
  var html = HtmlService.createHtmlOutput(`Click the link to complete authentication with Strava. <a target="_blank" href="${authorizationUrl}">Strava Authentication</a> <input type="button" value="Close" onclick="google.script.host.close()"/>`)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'My custom dialog');

    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}