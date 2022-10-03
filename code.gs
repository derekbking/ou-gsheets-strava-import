/**
 * @OnlyCurrentDoc
 */

const METERS_TO_MILE = 1609.34;
const METERS_TO_FEET = 3.28084;
const CELSIUS_TO_FAHRENHEIT_RATIO = 9 / 5;
const RUN_TYPE_KEY = "Run";

const WEATHER_API_KEY = "bde49a5a6b5c4e4ea69182155222508";
const MAX_LAPS_SYNC = 0;
const MAX_ACTIVITY_SYNC = 15;

const SUMMER_START_MONTH = 5;
const SUMMER_START_DAY = 12;

const SUMMER_END_MONTH = 8;
const SUMMER_END_DAY = 15;

const SHEETS_EXEMPT_FROM_OFF_DAYS = ["Derek"];

const colors = {
  LongRun: "#99ccff",
  Workout: "#c27ba0",
  Recovery: "#d9d2e9",
  Easy: "#ccffcc",
  Off: "#ffff00",
  Off2: "#ffffff",
};

const trackedColumns = {
  Type: "Type",
  Mileage: "Distance",
  Runs: "Activity Distance",
  Date: "Date",
  Cadence: "Cadence",
  Elevation: "Elevation Change",
  Pace: "Average Pace",
  HeartRate: "Average Heart Rate",
  // LapData: "Laps",
  // RawLapData: "RawLapData",
  RawActivityData: "RawActivityData",
  Temperature: "Temperature",
  MovingTime: "Total Moving Time",
  Wind: "Wind",
  Humidity: "Humidity",
};

const startRowIndex = 4;
const ss = SpreadsheetApp.getActiveSpreadsheet();
let sheet = ss.getActiveSheet();

const MAX_API_REQUESTS = 4;
let apiRequests = 0;

// Create toolbar dropdown items
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Only show logout button if user is authenticated.
  ui.createMenu("Strava")
    .addItem("Sync Current Tab", "updateSheet")
    .addItem("View Splits", "viewSplits")
    .addToUi();

  ui.createMenu("Admin")
    .addItem("Update Login", "updateLogin")
    .addItem("Update Column Map", "getColumnMap")
    .addItem("Sync Current Sheet", "updateSheet")
    .addItem("Sync All Sheets", "updateAllSheets")
    .addItem("Logout", "logout")
    .addToUi();
}

function updateLogin() {
  const service = getStravaService(sheet.getName());

  if (!service.hasAccess()) {
    Logger.log("User not logged in. Started authorization flow.");
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl({
      sheetName: sheet.getName(),
    });

    var html = HtmlService.createHtmlOutput(
      `<p style='font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif;'>Click the link to complete authentication with Strava.</p><a target="_blank" href="${authorizationUrl}" style='padding: 1rem; background-color: #fc4c01; display: block; border-radius: 5px; text-align: center; color: white; font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif; text-decoration: none;'>Continue to Strava</a>`
    )
      .setWidth(400)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, "Strava Authentication");
    return;
  }
}

function getLoggedInUsers() {
  Logger.log(OAuth2.getServiceNames(PropertiesService.getScriptProperties()));
}

function createTriggerEven() {
  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(16)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .create();
}

function createTriggerOdd() {
  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(11)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(13)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("updateAllSheets")
    .timeBased()
    .atHour(19)
    .everyDays(1)
    .create();
}

function viewSplits() {
  const [colNameToIndex, indexToColName] = getColumnMap(true);
  const cell = sheet.getActiveCell();
  const activityDataList = JSON.parse(
    getCellData(
      cell.getRowIndex(),
      colNameToIndex.get(trackedColumns.RawActivityData) + 1
    )
  ).reverse();

  var html = HtmlService.createHtmlOutput(
    `
     <script>
       function switchTab(activityIndex, tab) {
         const splits = document.getElementById(\`content-splits-\${activityIndex}\`)
         const splitsTab = document.getElementById(\`tab-splits-\${activityIndex}\`)
         const laps = document.getElementById(\`content-laps-\${activityIndex}\`)
         const lapsTab = document.getElementById(\`tab-laps-\${activityIndex}\`)
         if (tab == 'splits') {
           splits.classList.remove('hidden')
           splitsTab.classList.add('active')
           laps.classList.add('hidden')
           lapsTab.classList.remove('active')
         } else {
           splits.classList.add('hidden')
           splitsTab.classList.remove('active')
           laps.classList.remove('hidden')
           lapsTab.classList.add('active')
         }
       }
       console.log("Test hello");
     </script>
    <style>
      table {
        border: 1px solid #dfdfe8;
      }
 
 
     .tab-button {
       cursor: pointer;
       background: none;
       outline: none;
       border: none;
       position: relative;
       padding-bottom: 5px;
       transition: color 400ms ease;
     }
 
     .tab-button:focus {
       outline: none;
     }
 
     .tab-button.active {
       color: hsl(196, 100%, 47%);
     }
 
     .tab-button.active::after {
       background-color: hsl(196, 100%, 47%);
     }
 
     .tab-button::after {
       position: absolute;
       content: "";
       border-radius: 3px;
       width: 100%;
       height: 3px;
       bottom: 0;
       left: 0;
       right: 0;
       transform: scale(0.6);
       background-color: hsl(180, 8%, 83%);
       transition: transform 200ms ease, opacity 400ms ease, background-color 400ms ease;
       opacity: 0;
     }
 
     .tab-button:hover::after,
     .tab-button.active::after {
       transform: scale(1);
       opacity: 1;
     }
 
      .hidden {
        display: none;
      }
    
      table td, table th {
        padding: 6px 20px;
        border-bottom: 1px solid #dfdfe8;
      }
    
      table tr:last-child td {
        border-bottom: none;
      }
    
      table tr {
        text-align: center;
        font-size: 14px;
      }
    </style>
    <div style='display: flex; gap: 2rem; flex-direction: column; font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif;'>${activityDataList
      .map((activityData, activityIndex) => {
        return `
    <div>
    <h3>${activityData.name} (${metersToMiles(
          activityData.distance
        )} miles) - ${new Date(
          activityData.start_date
        ).toLocaleTimeString()}</h3>
    ${
      !!activityData.laps
        ? `
     <div style='display: flex; gap: 1rem; margin-bottom: .5rem;'>
       <button class="tab-button active" id='tab-laps-${activityIndex}' onclick="switchTab(${activityIndex}, 'laps')">Laps (${
            activityData.laps?.length ?? 0
          })</button>
       <button class="tab-button" id='tab-splits-${activityIndex}' onclick="switchTab(${activityIndex}, 'splits')">Splits (${
            activityData.splits?.length ?? 0
          })</button>
     </div>
     <div id='content-laps-${activityIndex}' class=''>
       ${!activityData.laps ? "No lap data available." : ""}
       ${
         !!activityData.laps
           ? `
       <table style='width: 100%; border-spacing: 0;'>
         <thead>
           <tr style="text-align: center; background-color: #f7f7fa;">
             <th>Lap</th>
             <th>Distance</th>
             <th>Time</th>
             <th>Pace</th>
             <th>Elevation Gain</th>
             <th>HR</th>
           </tr>
         </thead>
         ${activityData.laps
           ?.map((lap, index) => {
             return `<tr><td>${(index + 1).toString()}</td><td>${metersToMiles(
               lap.distance
             )}</td><td>${durationToTime(
               lap.moving_time
             )}</td><td>${speedToPace(lap.average_speed)}</td><td>${Math.round(
               lap.total_elevation_gain * METERS_TO_FEET
             )?.toFixed(1)} ft</td><td>${lap.average_heartrate?.toFixed(
               1
             )}</td></tr>`;
           })
           .join("\n")}
       </table>`
           : ""
       }
     </div>
     <div id='content-splits-${activityIndex}' class="hidden">
       ${!activityData.splits ? "No split data available." : ""}
       ${
         !!activityData.splits
           ? `
 <table style='width: 100%; border-spacing: 0;'>
         <thead>
           <tr style="text-align: center; background-color: #f7f7fa;">
             <th>Split</th>
             <th>Distance</th>
             <th>Time</th>
             <th>Pace</th>
             <th>Elevation Difference</th>
             <th>HR</th>
           </tr>
         </thead>
         ${activityData.splits
           ?.map((lap, index) => {
             return `<tr><td>${(index + 1).toString()}</td><td>${metersToMiles(
               lap.distance
             )}</td><td>${durationToTime(
               lap.moving_time
             )}</td><td>${speedToPace(lap.average_speed)}</td><td>${Math.round(
               lap.elevation_difference * METERS_TO_FEET
             )?.toFixed(1)} ft</td><td>${lap.average_heartrate?.toFixed(
               1
             )}</td></tr>`;
           })
           .join("\n")}
       </table>`
           : ``
       }
     </div>    
 `
        : ""
    }
    </div>`;
      })
      .join("\n")}
    </div>`
  )
    .setWidth(900)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(
    html,
    `Splits ${new Date(activityDataList[0].start_date).toLocaleDateString()}`
  );
}

function getCellData(rowIndex, colIndex) {
  var cell = sheet.getRange(rowIndex, colIndex, 1, 1);
  return cell.getValues();
}

function setCellData(rowIndex, colIndex, value) {
  var cell = sheet.getRange(rowIndex, colIndex, 1, 1);
  cell.setValues([[value]]);
  cell.setHorizontalAlignment("left");
}

function updateAllSheets() {
  let index = 0;
  for (const currentSheet of ss.getSheets()) {
    if (index == 0) {
      index++;
      continue;
    }

    Logger.log("Updating sheet: " + currentSheet.getName());
    try {
      updateSheet(currentSheet);
    } catch (e) {
      Logger.log(`Failed to update sheet ${currentSheet.getName()}: ${e}`);
    }
    index++;
  }
  Logger.log(`Made a total of ${apiRequests} API requests.`);
}

function updateCurrentSheet() {
  updateSheet(sheet);
}

function testUpdateSheet() {
  updateSheet(ss.getSheetByName("Derek"));
}

function updateSheet(suppliedSheet) {
  if (!!suppliedSheet) {
    sheet = suppliedSheet;
  }

  const service = getStravaService(sheet.getName());

  if (!service.hasAccess()) {
    if (!suppliedSheet) {
      try {
        updateLogin();
      } catch (e) {
        Logger.log(`Error initializing login set for ${sheet.getName()} (was this run by automation?): ${e}`);
      }
    }
    return;
  }

  Logger.log("Getting column indexes...");
  const [colNameToIndex, indexToColName] = getColumnMap(true);

  Logger.log(Array.from(colNameToIndex.entries()));

  Logger.log("Reading log data...");
  var logData = sheet.getRange(
    startRowIndex,
    Array.from(indexToColName.keys()).reduce(
      (acc, index) => {
        return index < acc ? index : acc;
      },
      [sheet.getLastColumn()]
    ),
    sheet.getLastRow() - (startRowIndex - 1),
    sheet.getLastColumn()
  );
  var logColors = sheet
    .getRange(
      startRowIndex,
      colNameToIndex.get(trackedColumns.Type) + 1,
      sheet.getLastRow() - (startRowIndex - 1),
      1
    )
    .getBackgrounds();
  Logger.log("Parsing log data...");
  var parsedDataMap = logData.getValues().reduce((acc, row, rowIndex) => {
    var color = logColors[rowIndex];

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
    if (
      Object.keys(parsedRow).includes(trackedColumns.Date) &&
      !!parsedRow[trackedColumns.Date] &&
      Object.keys(parsedRow).includes(trackedColumns.Type) &&
      !!parsedRow[trackedColumns.Type]
    ) {
      var date = parsedRow[trackedColumns.Date];
      var dateStr = `${date.getFullYear()}-${(date.getMonth() + 1)
        .toString()
        .padStart(2, "0")}-${date.getDate().toString().padStart(2, "0")}`;
      acc.set(dateStr, parsedRow);
    }

    return acc;
  }, new Map());
  Logger.log("Finish parsing data");
  var parsedDataRows = Array.from(parsedDataMap.values());

  var firstRowIndexMissingData = parsedDataRows.findIndex((row) => {
    return true;
    // return Object.entries(row).some(
    //   ([colName, col]) => !col && colName != "rowIndex"
    // );
  });
  var lastRowIndexMissingData =
    parsedDataRows.length -
    1 -
    [...parsedDataRows].reverse().findIndex((row) =>
      Object.entries(row).some(([colName, col]) => {
        if (!col && colName != "rowIndex") {
          Logger.log(`Missing data for ${colName}`);
          return true;
        }
        return false;
      })
    );
  var firstRowDate = new Date(
    parsedDataRows[firstRowIndexMissingData][trackedColumns.Date]
  );
  var lastRowDate = new Date(
    parsedDataRows[lastRowIndexMissingData][trackedColumns.Date]
  );

  var activities;

  try {
    Logger.log(
      `Getting Strava activities between ${firstRowDate.toLocaleDateString()} to ${lastRowDate.toLocaleDateString()}`
    );
    activities = getStravaActiviesInRange(service, firstRowDate, lastRowDate);
    Logger.log(`Got ${activities.length} activities.`);
  } catch (e) {
    var html = HtmlService.createHtmlOutput(`An error has occurred. ${e}`)
      .setWidth(400)
      .setHeight(300);
    Logger.log(e);
    SpreadsheetApp.getUi().showModalDialog(html, "Application Error");
    return;
  }

  Logger.log("Grouping activities by date...");
  var groupedActivities = groupBy(activities, (activity) => {
    var dateKey = activity.start_date_local.split("T")[0];

    return dateKey;
  });
  Logger.log("Begin inserting data to sheet.");
  Logger.log(JSON.stringify(Array.from(groupedActivities.keys())));
  var blankActivities = Array.from(groupedActivities.entries()).filter(
    (entry, index) => {
      const [dateKey, activities] = entry;
      var sheetData = parsedDataMap.get(dateKey);

      if (!sheetData) {
        Logger.log("Couldn't find: " + dateKey);
        return -1;
      }

      return !sheetData[trackedColumns.RawActivityData];
    }
  );
  var notBlankActivities = Array.from(groupedActivities.entries()).filter(
    (entry, index) => {
      const [dateKey, activities] = entry;
      var sheetData = parsedDataMap.get(dateKey);

      if (!sheetData) {
        Logger.log("Couldn't find: " + dateKey);
        return -1;
      }

      return !!sheetData[trackedColumns.RawActivityData];
    }
  );
  Array.from([...blankActivities, ...notBlankActivities]).forEach(
    (entry, index) => {
      const [dateKey, activities] = entry;
      Logger.log(`Inserting data for ${dateKey}`);
      var sheetData = parsedDataMap.get(dateKey);

      const splitDate = dateKey.split("-");
      const year = parseInt(splitDate[0]);
      const month = parseInt(splitDate[1]);
      const day = parseInt(splitDate[2]);
      const currentDay = new Date(year, month - 1, day);
      const startSummerDate = new Date(
        year,
        SUMMER_START_MONTH - 1,
        SUMMER_START_DAY
      );
      const endSummerDate = new Date(
        year,
        SUMMER_END_MONTH - 1,
        SUMMER_END_DAY
      );

      if (!sheetData) {
        Logger.log(
          `Skipped inserting Strava data for ${dateKey}. No associated row could be found on the sheet. Is the day marked as an off day?`
        );
        return;
      }

      var isSummerDay =
        currentDay.getTime() >= startSummerDate.getTime() &&
        currentDay.getTime() <= endSummerDate.getTime();
      var isOffDay =
        sheetData.color == colors.Off || sheetData.color == colors.Off2;
      if (
        (isSummerDay || isOffDay) &&
        !SHEETS_EXEMPT_FROM_OFF_DAYS.includes(sheet.getName())
      ) {
        for (let key of Object.values(trackedColumns)) {
          if (key === trackedColumns.Date || key === trackedColumns.Type) {
            continue;
          }

          setCellData(sheetData.rowIndex, colNameToIndex.get(key) + 1, "");
        }

        Logger.log(
          `Clearing data for ${dateKey}. ${
            isOffDay ? "Off day detected." : ""
          }${isSummerDay ? "Summer day detected." : ""}.`
        );
        return;
      }

      var aggregateData = activities.reduce((acc, activity) => {
        var cadence = activity.average_cadence
          ? `${Math.round(activity.average_cadence * 2)} spm`
          : "-";
        var elevation = `${Math.round(
          activity.total_elevation_gain * METERS_TO_FEET
        )} ft`;
        var movingTime = `${durationToTime(activity.moving_time)}`;
        var pace = speedToPace(activity.average_speed);

        return {
          totalDistance:
            (activity.type === RUN_TYPE_KEY ? activity.distance : 0) +
            (acc.totalDistance ?? 0),
          activities: [
            { distance: activity.distance, type: activity.type },
            ...(acc.activities ?? []),
          ],
          activityData: [activity, ...(acc.activityData ?? [])],
          paces: [pace, ...(acc.paces ?? [])],
          cadence: [cadence, ...(acc.cadence ?? [])],
          elevation: [elevation, ...(acc.elevation ?? [])],
          movingTime: [movingTime, ...(acc.movingTime ?? [])],
          averageHeartRates: [
            activity.average_heartrate
              ? `${Math.round(activity.average_heartrate)} bpm`
              : "-",
            ...(acc.averageHeartRates ?? []),
          ],
          activityIds: [...(acc.activityIds ?? []), activity.id],
        };
      }, {});

      var rawActivityDataList;
      try {
        var rawData = sheetData[trackedColumns.RawActivityData];
        rawActivityDataList =
          aggregateData.activityIds.length !==
          (!!rawData ? JSON.parse(rawData).length : 0)
            ? apiRequests <= MAX_API_REQUESTS
              ? aggregateData.activityIds.map((activityId) => {
                  return getActivityData(service, activityId);
                })
              : undefined
            : undefined;
      } catch (e) {
        Logger.log(`Get activity details exception. ${e}`);
      }

      var weatherDataList;
      try {
        var windDataLen = sheetData[trackedColumns.Wind]
          .toString()
          .split("\n")
          .filter((item) => !!item).length;
        var humidityDataLen = sheetData[trackedColumns.Humidity]
          .toString()
          .split("\n")
          .filter((item) => !!item).length;
        var temperatureDataLen = sheetData[trackedColumns.Temperature]
          .toString()
          .split("\n")
          .filter((item) => !!item).length;

        if (
          [windDataLen, humidityDataLen, temperatureDataLen].some(
            (length) => aggregateData.activityData.length !== length
          )
        ) {
          weatherDataList = aggregateData.activityData
            .map((activity) => {
              if (
                !activity.start_latlng ||
                !activity.start_latlng[0] ||
                !activity.start_latlng[1]
              ) {
                return "N/A";
              }

              return getWeather(
                activity.start_latlng[0],
                activity.start_latlng[1],
                new Date(
                  new Date(activity.start_date_local).getTime() +
                    (activity.elapsed_time * 1000) / 2
                )
              );
            })
            .map((data) => (!!data ? data : "Failed."));
        }
      } catch (e) {
        Logger.log(`Get weather details exception. ${e}`);
      }

      var formattedData = {
        [trackedColumns.Runs]: aggregateData.activities
          .map(
            (activity) =>
              `${metersToMiles(activity.distance)} mi${
                activity.type !== RUN_TYPE_KEY ? ` (${activity.type})` : ""
              }`
          )
          .join("\n"),
        [trackedColumns.Mileage]: metersToMiles(aggregateData.totalDistance),
        [trackedColumns.Temperature]: weatherDataList
          ?.map((data) => (data.temp_f ? `${data.temp_f.toFixed(1)} °F` : data))
          .join("\n"),
        [trackedColumns.Wind]: weatherDataList
          ?.map((data) => (data.wind_mph ? `${data.wind_mph} mph` : data))
          .join("\n"),
        [trackedColumns.Humidity]: weatherDataList
          ?.map((data) => (data.humidity ? `${data.humidity}%` : data))
          .join("\n"),
        [trackedColumns.MovingTime]: aggregateData.movingTime.join("\n"),
        [trackedColumns.Pace]: aggregateData.paces.join("\n"),
        [trackedColumns.HeartRate]: aggregateData.averageHeartRates.join("\n"),
        [trackedColumns.Cadence]: aggregateData.cadence.join("\n"),
        [trackedColumns.Elevation]: aggregateData.elevation.join("\n"),
        [trackedColumns.RawActivityData]: JSON.stringify(rawActivityDataList),
      };

      for (const [key, value] of Object.entries(formattedData)) {
        if (!value) {
          continue;
        }
        if (value === sheetData[key]) {
          continue;
        }

        setCellData(sheetData.rowIndex, colNameToIndex.get(key) + 1, value);
      }
    }
  );

  Logger.log(`Made a total of ${apiRequests} API requests.`);
}

function getColumnMap(useCache = false) {
  if (!useCache) {
    return getUpdatedColumnMap();
  }

  return getCachedColumnMap() ?? getUpdatedColumnMap();
}

function getCachedColumnMap() {
  const scriptProperties = PropertiesService.getScriptProperties();

  var cachedData = JSON.parse(
    scriptProperties.getProperty(`column-map-${sheet.getName}`)
  );

  if (!cachedData) {
    return;
  }

  return [new Map(cachedData[0]), new Map(cachedData[1])];
}

function getUpdatedColumnMap() {
  var columnHeaders = sheet.getRange(1, 1, 2, sheet.getLastColumn() + 1);

  var columnMaps = [
    ...columnHeaders.getValues()[0],
    ...columnHeaders.getValues()[1],
  ].reduce(
    (acc, cell, index) => {
      var [tmpColNameToIndex, tmpIndexToColName] = acc;
      if (Object.values(trackedColumns).includes(cell)) {
        tmpColNameToIndex.set(cell, index % (sheet.getLastColumn() + 1));
        tmpIndexToColName.set(index % (sheet.getLastColumn() + 1), cell);
      }

      return acc;
    },
    [new Map(), new Map()]
  );

  PropertiesService.getScriptProperties().setProperty(
    `column-map-${sheet.getName}`,
    JSON.stringify([
      Array.from(columnMaps[0].entries()),
      Array.from(columnMaps[1].entries()),
    ])
  );
  return columnMaps;
}

function logout(service) {
  var service = getStravaService(sheet.getName());
  service.reset();
}

// Weather API Functions

function getWeather(lat, lon, start) {
  var endpoint = `https://api.weatherapi.com/v1/history.json?key=${WEATHER_API_KEY}&q=${lat},${lon}&dt=${
    start.toISOString().split("T")[0]
  }`;

  var options = {
    method: "GET",
    muteHttpExceptions: true,
  };

  var response = JSON.parse(UrlFetchApp.fetch(endpoint, options));

  return response.forecast.forecastday[0].hour.find((forecaseHour) => {
    return (
      Math.abs(forecaseHour.time_epoch * 1000 - start.getTime()) <=
      31 * 60 * 1000
    );
  });
}

// Strava API Functions

function getStravaActiviesInRange(service, fromDate, toDate) {
  var endpoint = "https://www.strava.com/api/v3/athlete/activities";
  var params = `?after=${fromDate.getTime() / 1000}&before=${
    toDate.getTime() / 1000
  }&per_page=200`;

  var headers = {
    Authorization: "Bearer " + service.getAccessToken(),
  };

  var options = {
    headers: headers,
    method: "GET",
    muteHttpExceptions: true,
  };

  var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
  apiRequests++;

  if (!!response.errors) {
    throw response.message ?? JSON.stringify(response.errors);
  }

  return response;
}

function getActivityData(service, activityId) {
  Logger.log(`Get activity data ${activityId}`);
  var activityEndpoint = `https://www.strava.com/api/v3/activities/${activityId}`;

  var headers = {
    Authorization: "Bearer " + service.getAccessToken(),
  };

  var options = {
    headers: headers,
    method: "GET",
    muteHttpExceptions: true,
  };

  var res = UrlFetchApp.fetch(activityEndpoint, options);
  apiRequests++;

  if (!!res.errors) {
    throw res.message ?? JSON.stringify(res.errors);
  }

  var activityData = JSON.parse(res);

  return {
    name: activityData.name,
    distance: activityData.distance,
    moving_time: activityData.moving_time,
    elapsed_time: activityData.elapsed_time,
    type: activityData.type,
    sport_type: activityData.sport_type,
    start_date: activityData.start_date,
    start_date_local: activityData.start_date_local,
    timezone: activityData.timezone,
    start_latlng: activityData.start_latlng,
    end_latlng: activityData.end_latlng,
    average_speed: activityData.average_speed,
    max_speed: activityData.max_speed,
    average_cadence: activityData.average_cadence,
    average_temp: activityData.average_temp,
    average_heartrate: activityData.average_heartrate,
    max_heartrate: activityData.max_heartrate,
    calories: activityData.calories,
    splits: activityData.splits_standard?.map((split) => {
      return {
        elapsed_time: split.elapsed_time,
        distance: split.distance,
        elevation_difference: split.elevation_difference,
        moving_time: split.moving_time,
        average_speed: split.average_speed,
        average_grade_adjusted_speed: split.average_grade_adjusted_speed,
        average_heartrate: split.average_heartrate,
      };
    }),
    laps: activityData.laps?.map((lap) => {
      return {
        elapsed_time: lap.elapsed_time,
        start_date: lap.start_date,
        distance: lap.distance,
        total_elevation_gain: lap.total_elevation_gain,
        moving_time: lap.moving_time,
        average_speed: lap.average_speed,
        average_cadence: lap.average_cadence,
        average_heartrate: lap.average_heartrate,
        max_heartrate: lap.max_heartrate,
      };
    }),
  };
}

function getLapData(service, activityId) {
  var lapEndpoint = `https://www.strava.com/api/v3/activities/${activityId}/laps`;

  var headers = {
    Authorization: "Bearer " + service.getAccessToken(),
  };

  var options = {
    headers: headers,
    method: "GET",
    muteHttpExceptions: true,
  };

  var response = JSON.parse(UrlFetchApp.fetch(lapEndpoint, options));
  apiRequests++;

  if (!!response.errors) {
    throw response.message ?? JSON.stringify(response.errors);
  }

  return response;
}

// Utility Functions

function celsiusToFahrenheit(tempInCelsius) {
  return tempInCelsius * CELSIUS_TO_FAHRENHEIT_RATIO + 32;
}

function durationToTime(duration) {
  if (duration < 60) {
    return `${duration} sec`;
  }
  var hours = Math.trunc(duration / 3600);
  var minutes = Math.trunc((duration % 3600) / 60);
  var seconds = duration % 60;

  var durationParts = new Array();

  if (hours != 0) {
    durationParts.push(`${hours.toString().padStart(2, "0")}`);
  }

  if (minutes != 0) {
    durationParts.push(`${minutes.toString().padStart(2, "0")}`);
  }

  durationParts.push(`${seconds.toString().padStart(2, "0")}`);

  return durationParts.join(":");
}

function speedToPace(speed) {
  if (!speed) {
    return "-";
  }

  var secondPerMeter = Math.pow(speed, -1);
  var secondPerMile = secondPerMeter * 1609;
  return `${Math.trunc(secondPerMile / 60).toString()}:${Math.round(
    secondPerMile % 60
  )
    .toString()
    .padStart(2, "0")}`;
}

function metersToMiles(meters) {
  return (Math.round((meters / METERS_TO_MILE) * 100) / 100).toFixed(2);
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
