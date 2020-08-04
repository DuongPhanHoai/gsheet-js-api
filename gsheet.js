const fs = require('fs');
const util = require('util');
const mkdir = util.promisify(fs.mkdir);
const { google } = require('googleapis');
const { asyncGClientGetWebToken, asyncReadRange, asyncSetStringRange, asyncSetValuesRange, asyncInsertColumn } = require('./gUtil');

const REQUEST_DURATION = 1200;// sleep to prevent limitation of requests to google free account services
const SLEEP_ON_ERROR = 10000;

let GCONF_DIR = 'gconf';
let GCONF_CREDENTIAL_FILE = "gsheet-auth.json";
let GCONF_TOKEN_PATH = 'token.json';
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];

function sleep(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}
async function sleepByStart(ms, startTime) {
  const end = new Date().getTime();
  const durationMS = end - startTime;
  if (durationMS < ms)
    await sleep(durationMS - ms);
}

async function setConf(confDir = GCONF_DIR, confCredentialFile = GCONF_CREDENTIAL_FILE, confWebToken = GCONF_TOKEN_PATH) {
  GCONF_DIR = confDir;
  GCONF_CREDENTIAL_FILE = confCredentialFile;
  GCONF_TOKEN_PATH = confWebToken;
  await getCredentials();
}

/**
 * @name getCredentials
 * @description prepare the credentials and authorization for google sheet, google sheet auth file can be get from env variable GSHEET_AUTH, web token will be in gtokens/token.json
 */
async function getCredentials() {
  // Get credential file path
  // if (process.env.GSHEET_AUTH)
  //   credentialsFilePath = process.env.GSHEET_AUTH;
  // Check if credentials file is existed
  if (!fs.existsSync(`${GCONF_DIR}/${GCONF_CREDENTIAL_FILE}`))
    return null;
  let content = await fs.readFileSync(`${GCONF_DIR}/${GCONF_CREDENTIAL_FILE}`, 'utf8');
  const CREDENTIALS = JSON.parse(content);
  const { client_secret, client_id, redirect_uris } = CREDENTIALS.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Check if Token dir is Exists
  let shouldCreateToken = false;
  if (! await fs.existsSync(GCONF_DIR)) {
    shouldCreateToken = true;
    // Create Token Directory
    await mkdir(GCONF_DIR);
  }
  if (!shouldCreateToken && (! await fs.existsSync(`${GCONF_DIR}/${GCONF_TOKEN_PATH}`)))
    shouldCreateToken = true;
  if (shouldCreateToken) {
    await asyncGClientGetWebToken(`${GCONF_DIR}/${GCONF_TOKEN_PATH}`, oAuth2Client, SCOPES); // this function also setCredentials for oAuth2Client
  }
  else {
    // const tokenFromFile = fs.readFileSync(`${TOKEN_DIR}/${TOKEN_PATH}`, 'utf8');
    const tokenFromFile = fs.readFileSync(`${GCONF_DIR}/${GCONF_TOKEN_PATH}`, 'utf8');
    oAuth2Client.setCredentials(JSON.parse(tokenFromFile));
  }
  return oAuth2Client;
}

/**
 * @class GSheet represent the google sheet object which provide ability to read / write to google sheet
 */
class GSheet {
  /**
   * @constructor
   * @param {String} spreadSheetID the spreadSheetID get from Web URL
   */
  constructor(spreadSheetID) {
    this.spreadSheetID = spreadSheetID;
  }

  /**
   * @name init
   * @description init the google sheet object which need the credential and auth, note: do not change the code 'google.sheets({ version: 'v4', auth });', it has to be auth
   */
  async init() {
    this.oAuth2Client = await getCredentials();
    let auth = this.oAuth2Client
    this.sheets = google.sheets({ version: 'v4', auth });
  }

  /**
   * @name readRange
   * @description read the range from sheet
   * @param {String} sheetName 
   * @param {String} startCol 
   * @param {Number} startRow 
   * @param {String} endCol 
   * @param {Number} endRow
   * @returns the response object which will need to dig in result rows as: const rows = res.data.values;
   */
  async readRange(sheetName, startCol, startRow, endCol, endRow) {
    const startTime = new Date().getTime();
    let readResult = null;
    let retry = false;
    try {
      if (this.sheets) {
        readResult = await asyncReadRange(this.sheets, this.spreadSheetID, `${sheetName}!${startCol}${startRow}:${endCol}${endRow}`);
        await sleepByStart(REQUEST_DURATION, startTime);
      }
    } catch (error) {
      retry = true;
      console.error(`${new Date().toUTCString()} - Retry once from readRange Error : ${error}`);
    }
    if (retry) {
      await sleep(SLEEP_ON_ERROR);
      try {
        readResult = await asyncReadRange(this.sheets, this.spreadSheetID, `${sheetName}!${startCol}${startRow}:${endCol}${endRow}`);
        await sleepByStart(REQUEST_DURATION, startTime);
      } catch (error) {
        console.error(`${new Date().toUTCString()} - Failed on readRange retry : ${error}`);
      }
    }
    return readResult;
  }

  /**
   * @name setString
   * @description  write the value to range
   * @param {String} value value to write
   * @param {String} writeRange Ex: 'targetResult!C10:C10'
   */
  async setString(value, writeRange) {
    const startTime = new Date().getTime();
    let writeResult = null;
    let retry = false;
    try {
      if (this.sheets) {
        writeResult = asyncSetStringRange(this.sheets, this.spreadSheetID, value, writeRange);
        await sleepByStart(REQUEST_DURATION, startTime);
      }
    } catch (error) {
      retry = true;
      console.error(`${new Date().toUTCString()} - Retry once from setString Error : ${error}`);
    }
    
    if (retry) {
      await sleep(SLEEP_ON_ERROR);
      try {
        writeResult = asyncSetStringRange(this.sheets, this.spreadSheetID, value, writeRange);
        await sleepByStart(REQUEST_DURATION, startTime);
      } catch (error) {
        console.error(`${new Date().toUTCString()} - Failed on setString retry : ${error}`);
      }
    }
    return writeResult;
  }

  /**
   * @name setValues
   * @description  write the value to range
   * @param {String} values value to write type of [[]]
   * @param {String} writeRange Ex: 'targetResult!C10:C10'
   */
  async setValues(values, writeRange) {
    const startTime = new Date().getTime();
    let writeResult = null;
    let retry = false;
    try {
      if (this.sheets) {
        writeResult = asyncSetValuesRange(this.sheets, this.spreadSheetID, values, writeRange);
        await sleepByStart(REQUEST_DURATION, startTime);
      }
    } catch (error) {
      retry = true;
      console.error(`${new Date().toUTCString()} - Retry once from setValues Error : ${error}`);
    }
    if (retry) {
      await sleep(SLEEP_ON_ERROR);
      try {
        writeResult = asyncSetValuesRange(this.sheets, this.spreadSheetID, values, writeRange);
        await sleepByStart(REQUEST_DURATION, startTime);
      } catch (error) {
        console.error(`${new Date().toUTCString()} - Failed on setValues retry : ${error}`);
      }
    }
    return writeResult;
  }

  /**
   * 
   * @param {Number} columnIndex 
   * @param {String} sheetName 
   */
  async insertColumn(columnIndex, sheetName) {
    const startTime = new Date().getTime();
    let runResult = null;
    let retry = false;
    try {
      if (this.sheets) {
        runResult = await asyncInsertColumn(this.sheets, this.spreadSheetID, this.oAuth2Client, sheetName, columnIndex);
        await sleepByStart(REQUEST_DURATION, startTime);
      }
    } catch (error) {
      retry = true;
      console.error(`${new Date().toUTCString()} - Retry once from insertColumn Error : ${error}`);
    }
    if (retry) {
      await sleep(SLEEP_ON_ERROR);// sleep 5s
      try {
        runResult = await asyncInsertColumn(this.sheets, this.spreadSheetID, this.oAuth2Client, sheetName, columnIndex);
        await sleepByStart(REQUEST_DURATION, startTime);
      } catch (error) {
        console.error(`${new Date().toUTCString()} - Failed on insertColumn retry : ${error}`);
      }
    }
    return runResult;
  }
}

let sheetSets = [];
/**
 * @name getGSheet
 * @description the factory will help to get the right sheet object by it spreadSheetID
 * @param {String} spreadSheetID spreadSheetID which get from URL
 */
async function getGSheet(spreadSheetID) {
  // Check if sheet is in list
  if (!sheetSets || !sheetSets[spreadSheetID]) {
    let newGSheet = new GSheet(spreadSheetID);
    await newGSheet.init();
    sheetSets[spreadSheetID] = newGSheet;
    return newGSheet;
  }
  else
    return sheetSets[spreadSheetID];
}

/**
 * @name readRange
 * @description read the range from sheet
 * @param {String} sheetName 
 * @param {String} startCol 
 * @param {Number} startRow 
 * @param {String} endCol 
 * @param {Number} endRow
 * @param {String} spreadSheetID the spreadSheetID from URL
 * @returns the response object which will need to dig in result rows as: const rows = res.data.values;
 */
async function readRange(sheetName, startCol, startRow, endCol, endRow, spreadSheetID) {
  const foundSheet = await getGSheet(spreadSheetID);
  if (foundSheet)
    return await foundSheet.readRange(sheetName, startCol, startRow, endCol, endRow);
  return null;
}

/**
 * @name setString
 * @description write the value to range
 * @param {String} value value string to write
 * @param {String} writeRange Ex: 'targetResult!C10:C10'
 * @param {String} spreadSheetID spreadSheetID which get from URL
 */
async function setString(value, writeRange, spreadSheetID) {
  const foundSheet = await getGSheet(spreadSheetID);
  if (foundSheet)
    return await foundSheet.setString(value, writeRange);
  return null;
}

/**
 * @name setValues
 * @description write the value to range
 * @param {String} values value to write
 * @param {String} writeRange Ex: 'targetResult!C10:C10'
 * @param {String} spreadSheetID spreadSheetID which get from URL
 */
async function setValues(values, writeRange, spreadSheetID) {
  const foundSheet = await getGSheet(spreadSheetID);
  if (foundSheet)
    return await foundSheet.setValues(values, writeRange);
  return null;
}

/**
 * @name insertColumn
 * @description insert a column the the colIndex
 * @param {Number} columnIndex 
 * @param {String} spreadSheetID spreadSheetID which get from URL
 */
async function insertColumn(columnIndex, sheetName, spreadSheetID) {
  const foundSheet = await getGSheet(spreadSheetID);
  if (foundSheet)
    return await foundSheet.insertColumn(columnIndex, sheetName);
  return null;
}

module.exports = {
  setConf,
  getGSheet,
  readRange,
  setString,
  setValues,
  insertColumn
}
