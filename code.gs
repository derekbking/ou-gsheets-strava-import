const METERS_TO_MILE = 1609.34

const colors = {
  LongRun: "#99ccff",
  Workout: "#c27ba0",
  Recovery: "#d9d2e9",
  Easy: "#ccffcc",
  Off: "#ffffff"
}

const trackedColumns = {
  Type: "Type",
  Mileage: "Mileage",
  Runs: "SSRuns",
  Date: "Date",
  Pace: "Pace",
  HeartRate: "HeartRate",
  LapData: "Splits",
};

const startRowIndex = 4;
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getActiveSheet();

// Create toolbar dropdown items
function onOpen() {
  const service = getStravaService();

  var ui = SpreadsheetApp.getUi();

  ui.createMenu("Strava App")
    .addItem("Sync Data", "updateSheet")
    .addItem("Logout", "logout")
    .addToUi();
}

function setCellData(rowIndex, colIndex, value) {
  var cell = sheet.getRange(rowIndex, colIndex, 1, 1);
  cell.setValues([[value]]);
}

function updateSheet() {
  const service = getStravaService();

  if (!service.hasAccess()) {
    Logger.log("User not logged in. Started authorization flow.")
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();

    var html = HtmlService.createHtmlOutput(`<p style='font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif;'>Click the link to complete authentication with Strava.</p><a target="_blank" href="${authorizationUrl}" style='padding: 1rem; background-color: #fc4c01; display: block; border-radius: 5px; text-align: center; color: white; font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif; text-decoration: none;'>Continue to Strava</a>`)
      .setWidth(400)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Strava Authentication');
    return;
  }

  var columnHeaders = sheet.getRange(1, 1, 2, sheet.getLastColumn());
  var [colNameToIndex, indexToColName] = [
    ...columnHeaders.getValues()[0],
    ...columnHeaders.getValues()[1],
  ].reduce(
    (acc, cell, index) => {
      var [tmpColNameToIndex, tmpIndexToColName] = acc;
      if (Object.values(trackedColumns).includes(cell)) {
        tmpColNameToIndex.set(cell, index % sheet.getLastColumn());
        tmpIndexToColName.set(index % sheet.getLastColumn(), cell);
      }

      return acc;
    },
    [new Map(), new Map()]
  );

  // Print colNameToIndex
  Logger.log(JSON.stringify(Array.from(colNameToIndex.entries())));
  // Print indexToColName
  Logger.log(JSON.stringify(Array.from(indexToColName.entries())));

  var logData = sheet.getRange(
    startRowIndex,
    1,
    sheet.getLastRow() - (startRowIndex - 1),
    sheet.getLastColumn()
  );
  var parsedDataMap = logData.getValues().reduce((acc, row, rowIndex) => {
    var color = logData.getCell(rowIndex + 1, colNameToIndex.get(trackedColumns.Type) + 1).getBackground();

    // Do not parse off days.
    if (color == colors.Off) {
      return acc;
    }

    // Parse cell data for column indexes that we track.
    var parsedRow = row.reduce(
      (acc, col, index) => {
        if (indexToColName.has(index)) {
          return {
            ...acc,
            [indexToColName.get(index)]: col,
          };
        }
        return acc;
      },
      { rowIndex: rowIndex + startRowIndex, color: color }
    );

    // Only append parsed data if date column was successfully found.
    if (Object.keys(parsedRow).includes(trackedColumns.Date) && !!parsedRow[trackedColumns.Date]) {
      var date = parsedRow[trackedColumns.Date];
      var dateStr = `${date.getFullYear()}-${(date.getMonth() + 1)
        .toString()
        .padStart(2, "0")}-${date.getDate().toString().padStart(2, "0")}`;
      acc.set(dateStr, parsedRow);
    }

    return acc;
  }, new Map());
  var parsedDataRows = Array.from(parsedDataMap.values());

  var firstRowIndexMissingData = parsedDataRows.findIndex((row) => {
    return Object.values(row).some((col) => !col);
  });
  var lastRowIndexMissingData =
    parsedDataRows.length -
    1 -
    [...parsedDataRows]
      .reverse()
      .findIndex((row) => Object.values(row).some((col) => !col));
  var firstRowDate = new Date(
    parsedDataRows[firstRowIndexMissingData][trackedColumns.Date]
  );
  var lastRowDate = new Date(
    parsedDataRows[lastRowIndexMissingData][trackedColumns.Date]
  );

  var activities = getStravaActiviesInRange(service, firstRowDate, lastRowDate);

  var groupedActivities = groupBy(activities, (activity) => {
    var dateKey = activity.start_date_local.split("T")[0];

    return dateKey;
  });
  Array.from(groupedActivities.entries()).forEach((entry, index) => {
    const [dateKey, activities] = entry;
    var sheetData = parsedDataMap.get(dateKey);

    var aggregateData = activities.reduce((acc, activity) => {
      var secondPerMeter = Math.pow(activity.average_speed, -1);
      var secondPerMile = secondPerMeter * 1609;
      var pace = `${Math.trunc(secondPerMile / 60).toString()}:${Math.round(
        secondPerMile % 60
      )
        .toString()
        .padStart(2, "0")}`;

      return {
        totalDistance: activity.distance + (acc.totalDistance ?? 0),
        runs: [...(acc.runs ?? []), activity.distance],
        paces: [...(acc.paces ?? []), pace],
        averageHeartRates: [
          ...(acc.averageHeartRates ?? []),
          `${Math.round(activity.average_heartrate)} bpm`,
        ],
        lapDataList: [
          ...(index <= 10 ? [getLapData(service, activity.id)] : []),
          ...(acc.lapDataList ?? []),
        ],
        activityDataList: [
          // ...(index <= 20 ? [getActivityData(service, activity.id)] : []),
          // ...(acc.activityDataList ?? []),
        ],
      };
    }, {});

    setCellData(
      sheetData.rowIndex,
      colNameToIndex.get(trackedColumns.Mileage) + 1,
      metersToMiles(aggregateData.totalDistance)
    );
    setCellData(
      sheetData.rowIndex,
      colNameToIndex.get(trackedColumns.Runs) + 1,
      aggregateData.runs
        .map((run) => `${metersToMiles(run)} mi`)
        .join("\n")
    );
    setCellData(
      sheetData.rowIndex,
      colNameToIndex.get(trackedColumns.Pace) + 1,
      aggregateData.paces.join("\n")
    );
    setCellData(
      sheetData.rowIndex,
      colNameToIndex.get(trackedColumns.HeartRate) + 1,
      aggregateData.averageHeartRates.join("\n")
    );
    if (aggregateData.lapDataList?.length != 0) {
      setCellData(
        sheetData.rowIndex,
        colNameToIndex.get(trackedColumns.LapData) + 1,
        aggregateData.lapDataList
          .map((laps) =>
            laps
              .map((lap) => {
                var secondPerMeter = Math.pow(lap.average_speed, -1);
                var secondPerMile = secondPerMeter * 1609;
                var pace = `${Math.trunc(
                  secondPerMile / 60
                ).toString()}:${Math.round(secondPerMile % 60)
                  .toString()
                  .padStart(2, "0")}`;

                return `${sheetData.color === colors.Workout ? `${durationToTime(lap.moving_time)} (${lap.distance !== METERS_TO_MILE ? `${metersToMiles(lap.distance)} mi ` : ""}${Math.round(lap.average_heartrate)} bpm)` : `${durationToTime(lap.moving_time)}${lap.distance !== METERS_TO_MILE ? ` (${metersToMiles(lap.distance)} mi)` : ""}`}`;
              })
              .join(", ")
          )
          .join("\n")
      );
    }
  });
}

function logout(service) {
  var service = getStravaService();
  service.reset();
}

// Strava API Functions

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

// Utility Functions

function durationToTime(duration) {
  if (duration < 60) {
    return `${duration} sec`;
  }

  return `${Math.trunc(duration / 60)}:${(duration % 60).toString().padStart(2, "0")}`
}

function metersToMiles(meters) {
  return Math.round((meters / METERS_TO_MILE) * 100) / 100;
}

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
