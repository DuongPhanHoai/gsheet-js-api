const fs = require('fs');
const util = require('util');
const mkdir = util.promisify(fs.mkdir);
const { google } = require('googleapis');
const { asyncGClientGetWebToken: asyncGClientGetWebToken, asyncReadRange, asyncSetValueRange, asyncInsertColumn } = require('./gUtil');

const TOKEN_DIR = 'gtokens';
const TOKEN_PATH = 'token.json';
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];

let credentialsFilePath = "gsheet-auth.json";

/**
 * @name getCredentials
 * @description prepare the credentials and authorization for google sheet, google sheet auth file can be get from env variable GSHEET_AUTH, web token will be in gtokens/token.json
 */
async function getCredentials() {
  // Get credential file path
  if (process.env.GSHEET_AUTH)
    credentialsFilePath = process.env.GSHEET_AUTH;
  // Check if credentials file is existed
  if (!fs.existsSync(credentialsFilePath))
    return null;
  let content = await fs.readFileSync(credentialsFilePath, 'utf8');
  const CREDENTIALS = JSON.parse(content);
  const { client_secret, client_id, redirect_uris } = CREDENTIALS.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Check if Token dir is Exists
  let shouldCreateToken = false;
  if (! await fs.existsSync(TOKEN_DIR)) {
    shouldCreateToken = true;
    // Create Token Directory
    await mkdir(TOKEN_DIR);
  }
  if (!shouldCreateToken && (! await fs.existsSync(`${TOKEN_DIR}/${TOKEN_PATH}`)))
    shouldCreateToken = true;
  if (shouldCreateToken) {
    await asyncGClientGetWebToken(`${TOKEN_DIR}/${TOKEN_PATH}`, oAuth2Client, SCOPES); // this function also setCredentials for oAuth2Client
  }
  else {
    // const tokenFromFile = fs.readFileSync(`${TOKEN_DIR}/${TOKEN_PATH}`, 'utf8');
    const tokenFromFile = fs.readFileSync(`${TOKEN_DIR}/${TOKEN_PATH}`, 'utf8');
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
    if (this.sheets) {
      const readResult = await asyncReadRange(this.sheets, this.spreadSheetID, `${sheetName}!${startCol}${startRow}:${endCol}${endRow}`);
      return readResult;
    }
    return null;
  }

  /**
   * @name setValue
   * @description  write the value to range
   * @param {String} value value to write
   * @param {String} writeRange Ex: 'targetResult!C10:C10'
   */
  async setValue(value, writeRange) {
    if (this.sheets) {
      const writeResult = asyncSetValueRange(this.sheets, this.spreadSheetID, value, writeRange);
      return writeResult;
    }
    return null;
  }

  /**
   * 
   * @param {Number} columnIndex 
   * @param {String} sheetName 
   */
  async insertColumn(columnIndex, sheetName) {
    if (this.sheets)
      return await asyncInsertColumn(this.sheets, this.spreadSheetID, this.oAuth2Client, sheetName, columnIndex);
    return null;
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
 * @name setValue
 * @description write the value to range
 * @param {String} value value to write
 * @param {String} writeRange Ex: 'targetResult!C10:C10'
 * @param {String} spreadSheetID spreadSheetID which get from URL
 */
async function setValue(value, writeRange, spreadSheetID) {
  const foundSheet = await getGSheet(spreadSheetID);
  if (foundSheet)
    return await foundSheet.setValue(value, writeRange);
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
  getGSheet,
  readRange,
  setValue,
  insertColumn
}
