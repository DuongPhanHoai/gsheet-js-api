const readline = require('readline');
const { google } = require('googleapis');
const fs = require('fs');

/**
 * @name asyncReadConsoleLine
 * @description method to read a line from console
 * @returns the text line get from console
 */
async function asyncReadConsoleLine() {
  return new Promise(function (resolve, relect) {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      resolve(code);
    });
  });
}

/**
 * @name asyncGClientGetWebOriginToken
 * @description The web token will be asked for and input from Console
 * @param {google.auth.OAuth2} oAuth2Client 
 * @param {String[]} SCOPES
 */
async function asyncGClientGetWebOriginToken(oAuth2Client, SCOPES) {
  return new Promise(async function (resolve) {
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);

    // Get Keys line from keyboard
    const tokenFromWeb = await asyncReadConsoleLine();

    resolve(tokenFromWeb);
  });
}

/**
 * @name asyncGClientGetWebToken
 * @description If there is not already token, the web token will be asked for and input from Console
 * @param {String} tokenPath the web token file path
 * @param {google.auth.OAuth2} oAuth2Client
 * @param {String[]} SCOPES
 */
async function asyncGClientGetWebToken(tokenPath, oAuth2Client, SCOPES) {
  const TOKEN_FROM_WEB = await asyncGClientGetWebOriginToken(oAuth2Client, SCOPES);
  return new Promise(function (resolve, reject) {
    // Get Token From Web
    console.log(`>>>>>> getToken TOKEN_FROM_WEB ${TOKEN_FROM_WEB}`);
    oAuth2Client.getToken(TOKEN_FROM_WEB, (err, token) => {
      if (err) reject('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(tokenPath, JSON.stringify(token), (err) => {
        if (err) reject(err);
        resolve(token);
      });
    });
  });
}

/**
 * @name asyncGSheetGet
 * @description get the Google sheet information
 * @param {google.sheets} sheets 
 * @param {String} inputSheetId 
 * @param {google.auth.OAuth2} authClient
 * @returns the response of getting spreadsheet info
 */
async function asyncGSheetGet(sheets, inputSheetId, authClient) {
  return new Promise(function (resolve, reject) {
    if (sheets) {
      var request = {
        // The spreadsheet to request.
        spreadsheetId: inputSheetId,  // TODO: Update placeholder value.

        // The ranges to retrieve from the spreadsheet.
        ranges: [],  // TODO: Update placeholder value.

        // True if grid data should be returned.
        // This parameter is ignored if a field mask was set in the request.
        includeGridData: false,  // TODO: Update placeholder value.

        auth: authClient,
      };

      sheets.spreadsheets.get(request, function (err, response) {
        if (err) {
          console.error(err);
          reject(err);
        }

        // TODO: Change code below to process the `response` object:
        resolve(response, null, 2);
      });
    }
  });
}

/**
 * @name asyncReadRange
 * @description read the range to values
 * @param {google.sheets} sheets
 * @param {String} inputSheetId the sheet ID get from web
 * @param {String} inputRange range to get values
 */
async function asyncReadRange(sheets, inputSheetId, inputRange) {
  return new Promise(function (resolve, reject) {
    sheets.spreadsheets.values.get({
      spreadsheetId: inputSheetId,
      range: inputRange,
    }, (err, res) => {
      if (err)
        reject('The API returned an error: ' + err);
      resolve(res);
    });
  });
}

/**
 * @name asyncSetValueRange
 * @description write the value to range
 * @param {google.sheets} sheets
 * @param {String} inputSheetId the sheet ID get from web
 * @param {String} value value to write
 * @param {String} writeRange Ex: 'targetResult!C10:C10'
 */
async function asyncSetValueRange(sheets, inputSheetId, value, writeRange) {
  return new Promise(function (resolve, reject) {
    if (sheets) {
      const values = [[value],];
      const body = { values: values };
      sheets.spreadsheets.values.update({
        spreadsheetId: inputSheetId,
        range: writeRange,
        valueInputOption: 'USER_ENTERED',
        resource: body
      }, (err, res) => {
        if (err)
          reject('The API returned an error: ' + err);
        resolve(res);
      });
    }
  });
}


module.exports = {
  asyncReadConsoleLine,
  asyncGClientGetWebToken,
  asyncGSheetGet,
  asyncReadRange,
  asyncSetValueRange
}