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
async function asyncGClientGetWebTokenFromConsole(oAuth2Client, SCOPES) {
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
  const TOKEN_FROM_WEB = await asyncGClientGetWebTokenFromConsole(oAuth2Client, SCOPES);
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
 * @param {String} spreadSheetID the sheet ID get from web
 * @param {String} inputRange range to get values
 */
async function asyncReadRange(sheets, spreadSheetID, inputRange) {
  return new Promise(function (resolve, reject) {
    sheets.spreadsheets.values.get({
      spreadsheetId: spreadSheetID,
      range: inputRange,
    }, (err, res) => {
      if (err)
        reject('The API returned an error: ' + err);
      resolve(res);
    });
  });
}

/**
 * @name asyncSetStringRange
 * @description write the value to range
 * @param {google.sheets} sheets
 * @param {String} spreadSheetID the sheet ID get from web
 * @param {String} value value to write
 * @param {String} writeRange Ex: 'targetResult!C10:C10'
 */
async function asyncSetStringRange(sheets, spreadSheetID, value, writeRange) {
  return new Promise(function (resolve, reject) {
    if (sheets) {
      const values = [[value],];
      const body = { values: values };
      sheets.spreadsheets.values.update({
        spreadsheetId: spreadSheetID,
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

/**
 * @name asyncSetValuesRange
 * @description write the value to range
 * @param {google.sheets} sheets
 * @param {String} spreadSheetID the sheet ID get from web
 * @param {String} values value to write
 * @param {String} writeRange Ex: 'targetResult!C10:C10'
 */
async function asyncSetValuesRange(sheets, spreadSheetID, values, writeRange) {
  return new Promise(function (resolve, reject) {
    if (sheets) {
      const body = { 'values': values };
      sheets.spreadsheets.values.update({
        spreadsheetId: spreadSheetID,
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

let bufferCombineSpreadIDWithNameID = [];
function findInBuffer(spreadSheetID, sheetName) {
  for (let i = 0; i < bufferCombineSpreadIDWithNameID.length; i++) {
    const sheetInfo = bufferCombineSpreadIDWithNameID[i];
    if (sheetInfo.spreadSheetID.localeCompare(spreadSheetID, 'en', { sensitivity: 'base' }) === 0 && sheetInfo.sheetName.localeCompare(sheetName, 'en', { sensitivity: 'base' }) === 0) {
      return sheetInfo;
    }
  }
  return null;
}
function pushToBuffer(spreadSheetID, sheetName, sheetID) {
  for (let i = 0; i < bufferCombineSpreadIDWithNameID.length; i++) {
    const sheetInfo = bufferCombineSpreadIDWithNameID[i];
    if (sheetInfo.spreadSheetID.localeCompare(spreadSheetID, 'en', { sensitivity: 'base' }) === 0 && sheetInfo.sheetName.localeCompare(sheetName, 'en', { sensitivity: 'base' })) {
      return sheetInfo;
    }
  }
  const newSheetInfo = {
    'spreadSheetID': spreadSheetID,
    'sheetName': sheetName,
    'sheetID': sheetID
  };
  bufferCombineSpreadIDWithNameID.push(newSheetInfo);
  return newSheetInfo;
}
async function asyncGetSheetIDFromName(sheets, spreadSheetID, authClient, sheetName) {
  // Send request to get sheetID
  return new Promise(function (resolve, reject) {
    const sheetInfo = findInBuffer(spreadSheetID, sheetName);
    if (sheetInfo)
      resolve(sheetInfo.sheetID);
    else if (sheets) {
      var request = {
        // The spreadsheet to request.
        spreadsheetId: spreadSheetID,  // TODO: Update placeholder value.

        // The ranges to retrieve from the spreadsheet.
        ranges: [],  // TODO: Update placeholder value.

        // True if grid data should be returned.
        // This parameter is ignored if a field mask was set in the request.
        includeGridData: false,  // TODO: Update placeholder value.

        auth: authClient,
      };

      sheets.spreadsheets.get(request, function (err, sheetInfo) {
        if (err) {
          console.error(err);
          reject(-1);
        }
        if (sheetInfo.data && sheetInfo.data.sheets && sheetInfo.data.sheets.length && sheetInfo.data.sheets.length > 0) {
          const sheetsData = sheetInfo.data.sheets;
          for (let sheetIndex = 0; sheetIndex < sheetsData.length; sheetIndex++) {
            const scanName = sheetsData[sheetIndex].properties.title;
            pushToBuffer(spreadSheetID, scanName, sheetsData[sheetIndex].properties.sheetId);
            if (scanName.localeCompare(sheetName, 'en', { sensitivity: 'base' }) === 0) {
              isheetID = sheetsData[sheetIndex].properties.sheetId;
              break;
            }
          }
        }
        resolve(isheetID);
      });
    }
    else
      reject(-1);
  });
}
async function asyncInsertColumn(sheets, spreadSheetID, authClient, sheetName, colIndex) {
  const sheetID = await asyncGetSheetIDFromName(sheets, spreadSheetID, authClient, sheetName);
  return new Promise(function (resolve, reject) {
    if (sheets && sheetID >= 0) {
      const request = {
        auth: authClient,
        spreadsheetId: spreadSheetID,
        resource: {
          requests: [
            {
              'insertDimension': {
                'range': {
                  "sheetId": sheetID,
                  "dimension": "COLUMNS",
                  "startIndex": colIndex,
                  "endIndex": (colIndex + 1)
                },
                'inheritFromBefore': false
              }
            }
          ]
        }
      };
      sheets.spreadsheets.batchUpdate(request, function (err, response) {
        if (err)
          console.error(err);
        resolve(response);
      });
    }
  });
}


module.exports = {
  asyncReadConsoleLine,
  asyncGClientGetWebToken,
  asyncReadRange,
  asyncSetStringRange,
  asyncSetValuesRange,
  asyncInsertColumn
}