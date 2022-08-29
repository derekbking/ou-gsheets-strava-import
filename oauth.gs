var CLIENT_ID = '92704';
var CLIENT_SECRET = '8aa2d7d3cd3060c239164fb01824fbd157c1b7f8';
 
// configure the service
function getStravaService(sheetName) {
  Logger.log("Service: " + sheetName);
  return OAuth2.createService(`Strava_${sheetName}`)
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('activity:read_all');
}
 
// handle the callback
function authCallback(request) {
  var sheetName = request.parameter.sheetName;

  var stravaService = getStravaService(sheetName);
  var isAuthorized = stravaService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
