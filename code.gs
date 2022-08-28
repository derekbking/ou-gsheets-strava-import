const METERS_TO_MILE = 1609.34;
const CELSIUS_TO_FAHRENHEIT_RATIO = 9 / 5;

const WEATHER_API_KEY = "bde49a5a6b5c4e4ea69182155222508";
const MAX_LAPS_SYNC = 5;

const colors = {
  LongRun: "#99ccff",
  Workout: "#c27ba0",
  Recovery: "#d9d2e9",
  Easy: "#ccffcc",
  Off: "#ffff00",
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
  RawLapData: "RawLapData",
  Temperature: "Temperature",
  MovingTime: "Total Moving Time",
  Wind: "Wind",
  Humidity: "Humidity",
};

const startRowIndex = 4;
const ss = SpreadsheetApp.getActiveSpreadsheet();
let sheet = ss.getActiveSheet();

// Create toolbar dropdown items
function onOpen() {
  const service = getStravaService(sheet.getName());

  var ui = SpreadsheetApp.getUi();

  // Only show logout button if user is authenticated.
  ui.createMenu("Strava").addItem("Sync Current Tab", "updateSheet").addToUi();

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

function viewSplits() {
  const [colNameToIndex, indexToColName] = getColumnMap(true);
  const cell = sheet.getActiveCell();
  const splitData = JSON.parse(
    getCellData(
      cell.getRowIndex(),
      colNameToIndex.get(trackedColumns.RawLapData) + 1
    )
  );

  var html = HtmlService.createHtmlOutput(
    `
<style>
  table {
    border: 1px solid #dfdfe8;
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
<div style='display: flex; gap: 2rem; flex-direction: column; font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif;'>${splitData
      .map((activityLaps, activityIndex) => {
        return `
<div>
<h3>Activity ${(activityIndex + 1).toString()}</h3>
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
  ${activityLaps
    .map((lap, index) => {
      return `<tr><td>${(index + 1).toString()}</td><td>${metersToMiles(
        lap.distance
      )}</td><td>${durationToTime(lap.moving_time)}</td><td>${speedToPace(
        lap.average_speed
      )}</td><td>${lap.total_elevation_gain.toFixed(
        1
      )} ft</td><td>${lap.average_heartrate.toFixed(1)}</td></tr>`;
    })
    .join("\n")}
</table>
</div>`;
      })
      .join("\n")}
</div>`
  )
    .setWidth(900)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Splits");
}

function getCellData(rowIndex, colIndex) {
  var cell = sheet.getRange(rowIndex, colIndex, 1, 1);
  return cell.getValues();
}

function setCellData(rowIndex, colIndex, value) {
  var cell = sheet.getRange(rowIndex, colIndex, 1, 1);
  cell.setValues([[value]]);
}

function updateAllSheets() {
  for (const currentSheet of ss.getSheets()) {
    Logger.log("Updating sheet: " + currentSheet.getName());
    updateSheet(currentSheet);
  }
}

function updateCurrentSheet() {
  updateSheet(sheet);
}

function updateSheet(suppliedSheet) {
  if (!!suppliedSheet) {
    sheet = suppliedSheet;
  }

  const service = getStravaService(sheet.getName());

  if (!service.hasAccess()) {
    if (!suppliedSheet) {
      updateLogin();
    }
    return;
  }

  Logger.log("Getting column indexes...");
  const [colNameToIndex, indexToColName] = getColumnMap(true);

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

    // Do not parse off days.
    if (color == colors.Off) {
      Logger.log("Found off day: " + rowIndex);
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
    return Object.entries(row).some(
      ([colName, col]) => !col && colName != "rowIndex"
    );
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
  var blankLaps = Array.from(groupedActivities.entries()).filter(
    (entry, index) => {
      const [dateKey, activities] = entry;
      var sheetData = parsedDataMap.get(dateKey);

      if (!sheetData) {
        Logger.log("Couldn't find: " + dateKey);
        return -1;
      }

      return !sheetData[trackedColumns.RawLapData];
    }
  );
  var notBlankLaps = Array.from(groupedActivities.entries()).filter(
    (entry, index) => {
      const [dateKey, activities] = entry;
      var sheetData = parsedDataMap.get(dateKey);

      if (!sheetData) {
        Logger.log("Couldn't find: " + dateKey);
        return -1;
      }

      return !!sheetData[trackedColumns.RawLapData];
    }
  );
  Array.from([...blankLaps, ...notBlankLaps]).forEach((entry, index) => {
    const [dateKey, activities] = entry;
    Logger.log(`Inserting data for ${dateKey}`);
    var sheetData = parsedDataMap.get(dateKey);

    if (!sheetData) {
      Logger.log(
        `Skipped inserting Strava data for ${dateKey}. No associated row could be found on the sheet. Is the day marked as an off day?`
      );
      return;
    }

    var aggregateData = activities.reduce((acc, activity) => {
      var secondPerMeter = Math.pow(activity.average_speed, -1);
      var secondPerMile = secondPerMeter * 1609;
      var cadence = `${Math.round(activity.average_cadence * 2)} spm`;
      var elevation = `${activity.total_elevation_gain} ft`;
      var avgTemp = `${(
        Math.round(celsiusToFahrenheit(activity.average_temp) * 10) / 10
      ).toFixed(1)} °F`;
      var movingTime = `${durationToTime(activity.moving_time)}`;
      var pace = `${Math.trunc(secondPerMile / 60).toString()}:${Math.round(
        secondPerMile % 60
      )
        .toString()
        .padStart(2, "0")}`;
      var lapData;
      try {
        if (index <= MAX_LAPS_SYNC) {
          lapData = getLapData(service, activity.id);
        }
      } catch (e) {
        Logger.log(`Unable to retrieve lap data ${e}`);
      }

      var weatherData;
      try {
        var windData = sheetData[trackedColumns.Wind];
        var humidityData = sheetData[trackedColumns.Humidity];

        if (!windData || !humidityData) {
          Logger.log(`Getting weather data for ${dateKey}`);
          weatherData = getWeather(
            activity.start_latlng[0],
            activity.start_latlng[1],
            new Date(activity.start_date)
          );
        }
      } catch (e) {
        Logger.log(`Unable to retrieve weather data ${e}`);
      }

      return {
        totalDistance: activity.distance + (acc.totalDistance ?? 0),
        runs: [activity.distance, ...(acc.runs ?? [])],
        paces: [pace, ...(acc.paces ?? [])],
        cadence: [cadence, ...(acc.cadence ?? [])],
        elevation: [elevation, ...(acc.elevation ?? [])],
        avgTemp: [avgTemp, ...(acc.avgTemp ?? [])],
        movingTime: [movingTime, ...(acc.movingTime ?? [])],
        averageHeartRates: [
          `${Math.round(activity.average_heartrate)} bpm`,
          ...(acc.averageHeartRates ?? []),
        ],
        lapDataList: [
          ...(acc.lapDataList ?? []),
          ...(!!lapData ? [lapData] : []),
        ],
        weatherDataList: [
          ...(acc.weatherDataList ?? []),
          ...(!!weatherData ? [weatherData] : []),
        ],
        activityDataList: [
          // ...(acc.activityDataList ?? []),
          // ...(index <= 20 ? [getActivityData(service, activity.id)] : []),
        ],
      };
    }, {});

    var formattedData = {
      [trackedColumns.Mileage]: metersToMiles(aggregateData.totalDistance),
      [trackedColumns.Temperature]: aggregateData.avgTemp.join("\n"),
      [trackedColumns.Wind]: aggregateData.weatherDataList
        .map((data) => `${data.wind_mph} mph`)
        .join("\n"),
      [trackedColumns.Humidity]: aggregateData.weatherDataList
        .map((data) => `${data.humidity}%`)
        .join("\n"),
      [trackedColumns.MovingTime]: aggregateData.movingTime.join("\n"),
      [trackedColumns.Runs]: aggregateData.runs
        .map((run) => `${metersToMiles(run)} mi`)
        .join("\n"),
      [trackedColumns.Pace]: aggregateData.paces.join("\n"),
      [trackedColumns.HeartRate]: aggregateData.averageHeartRates.join("\n"),
      [trackedColumns.Cadence]: aggregateData.cadence.join("\n"),
      [trackedColumns.Elevation]: aggregateData.elevation.join("\n"),
      [trackedColumns.RawLapData]:
        aggregateData.lapDataList.length !== 0
          ? JSON.stringify(aggregateData.lapDataList)
          : undefined,
    };

    for (const [key, value] of Object.entries(formattedData)) {
      if (!value) {
        Logger.log(`Empty value. Skipping cell update. (${key})`);
        continue;
      }
      if (value === sheetData[key]) {
        continue;
      }

      setCellData(sheetData.rowIndex, colNameToIndex.get(key) + 1, value);
    }
  });
}

function getColumnMap(useCache = false) {
  if (!useCache) {
    return getUpdatedColumnMap();
  }

  return getCachedColumnMap() ?? getUpdatedColumnMap();
}

function getCachedColumnMap() {
  const scriptProperties = PropertiesService.getScriptProperties();

  var cachedData = JSON.parse(scriptProperties.getProperty("column-map"));

  if (!cachedData) {
    return;
  }

  return [new Map(cachedData[0]), new Map(cachedData[1])];
}

function getUpdatedColumnMap() {
  var columnHeaders = sheet.getRange(1, 1, 2, sheet.getLastColumn());

  var columnMaps = [
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

  Logger.log(JSON.stringify(Array.from(columnMaps[0].entries())));
  Logger.log(
    "Setting Column Map: " +
      JSON.stringify([
        Array.from(columnMaps[0].entries()),
        Array.from(columnMaps[1].entries()),
      ])
  );
  PropertiesService.getScriptProperties().setProperty(
    "column-map",
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
  }&per_page=100`;

  var headers = {
    Authorization: "Bearer " + service.getAccessToken(),
  };

  var options = {
    headers: headers,
    method: "GET",
    muteHttpExceptions: true,
  };

  var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));

  if (!!response.errors) {
    throw response.message ?? JSON.stringify(response.errors);
  }

  return response;
}

function getActivityData(service, activityId) {
  var lapEndpoint = `https://www.strava.com/api/v3/activities/${activityId}`;

  var headers = {
    Authorization: "Bearer " + service.getAccessToken(),
  };

  var options = {
    headers: headers,
    method: "GET",
    muteHttpExceptions: true,
  };

  if (!!response.errors) {
    throw response.message ?? JSON.stringify(response.errors);
  }

  return JSON.parse(UrlFetchApp.fetch(lapEndpoint, options));
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

  return `${Math.trunc(duration / 60)}:${(duration % 60)
    .toString()
    .padStart(2, "0")}`;
}

function speedToPace(speed) {
  var secondPerMeter = Math.pow(speed, -1);
  var secondPerMile = secondPerMeter * 1609;
  return `${Math.trunc(secondPerMile / 60).toString()}:${Math.round(
    secondPerMile % 60
  )
    .toString()
    .padStart(2, "0")}`;
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
