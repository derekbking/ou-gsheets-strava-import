# OU Strava -> Google Sheets import utility

This utility is designed to import data from [Strava API v3](https://developers.strava.com/docs/reference/) into Google Sheets.

## Requirements
* OAuth2 library (version 41)

## Notes

* Log rows with a Type column background of ![#ffffff](https://via.placeholder.com/15/ffffff/ffffff.png) `#ffffff` (white) will not be synced from Strava. White type columns are assumed to be off days. 
* Heart rate data will be added to each lap interval if the Type column is marked with the color code ![#c27ba0](https://via.placeholder.com/15/c27ba0/c27ba0.png) `#c27ba0`