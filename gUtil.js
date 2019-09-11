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

module.exports = {
  asyncReadConsoleLine,
  asyncGClientGetWebToken,
  asyncReadRange
}