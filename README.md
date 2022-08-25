# OU Strava -> Google Sheets import utility

This utility is designed to import data from [Strava API v3](https://developers.strava.com/docs/reference/) into Google Sheets.

## Requirements
* OAuth2 library (version 41)

## Notes

* Data for rows with a white ![#ffffff](https://via.placeholder.com/15/ffffff/ffffff.png) `#ffffff` Type column will not be synced. White type columns are assumed to be Off days. 
* Heart rate data will be added to each lap interval if the Type column is marked with the hex color code ![#c27ba0](https://via.placeholder.com/15/c27ba0/c27ba0.png) `#c27ba0`