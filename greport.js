const moment = require('moment');
const { readRange, setValue, insertColumn } = require('./gsheet');

const TEST_NAME_COLUMN = 'C';
const TEST_RESULT_COLUMN = 'E';
const TEST_RESULT_COLUMN_INDEX = 4;
const TEST_NAME_START_ROW = 5;
const MAX_BLANK_ROW = 5;

class GReport {

  /**
   * @constructor
   * @param {String} spreadSheetID the spreadSheetID get from Web URL
   */
  constructor(spreadSheetID) {
    this.spreadSheetID = spreadSheetID;
    this.maxRowIndex = 0;
  }

  /**
   * @name findTestByName
   * @description find test by testName in the sheet sheetName
   * @param {String} testName 
   * @param {String} sheetName 
   * @param {Boolean} allowExistingResult TRUE: find Test which has result or not; FALSE: find the test which does not has Result
   */
  async findTestByName(testName, sheetName, allowExistingResult) {
    if (testName && sheetName) {
      let blankCount = 0;
      for (let row10x = 0; row10x < 1000 && blankCount <= MAX_BLANK_ROW; row10x++) {
        const rangeValResponse = await readRange(sheetName, TEST_NAME_COLUMN, TEST_NAME_START_ROW + row10x * 10, TEST_NAME_COLUMN, TEST_NAME_START_ROW + row10x * 10 + 10, this.spreadSheetID);
        if (!rangeValResponse || !rangeValResponse.data || !rangeValResponse.data.values || !rangeValResponse.data.values.length)
          break;
        const rows = rangeValResponse.data.values;
        for (let rowIndex = 0; rowIndex < 10 && blankCount <= MAX_BLANK_ROW; rowIndex++) {
          const scanName = rows[rowIndex][0];
          if (scanName) {
            blankCount = 0;
            this.maxRowIndex = (TEST_NAME_START_ROW + row10x * 10 + rowIndex); // now it is current index
            if (scanName.localeCompare(testName, 'en', { sensitivity: 'base' }) === 0) {
              if (allowExistingResult)
                return this.maxRowIndex;
              else {
                // Get the current result to check if it is not existed
                const currentResultRes = await readRange(sheetName, TEST_RESULT_COLUMN, this.maxRowIndex, TEST_RESULT_COLUMN, this.maxRowIndex, this.spreadSheetID);
                if (!currentResultRes || !currentResultRes.data || !currentResultRes.data.values || !currentResultRes.data.values.length || currentResultRes.data.values.length === 0)
                  return this.maxRowIndex;
              }
            }
          }
          else
            blankCount++;
        }
      }
    }
    return -1;
  }

  async updateTestResultByName(testName, testResult, sheetName, overWriteResult) {
    let foundTestRow = await findTestByName(testName, sheetName, overWriteResult);
    if (foundTestRow > 0) {
      setValue(testResult, `${sheetName}!${TEST_RESULT_COLUMN}${foundTestRow}:${TEST_RESULT_COLUMN}${foundTestRow}`, this.spreadSheetID);
    }
    else {
      foundTestRow = this.maxRowIndex + 1;
      setValue(testName, `${sheetName}!${TEST_NAME_COLUMN}${foundTestRow}:${TEST_NAME_COLUMN}${foundTestRow}`, this.spreadSheetID);
      setValue(testResult, `${sheetName}!${TEST_RESULT_COLUMN}${foundTestRow}:${TEST_RESULT_COLUMN}${foundTestRow}`, this.spreadSheetID);
    }
  }
  /**
 * @name createNewResultCol
 * @description create new result column
 * @param {String} sheetName
 */
  async createNewResultCol(sheetName) {
    const insertColumnResult = await insertColumn(TEST_RESULT_COLUMN_INDEX, sheetName, this.spreadSheetID);
    if (insertColumnResult)
      await setValue(moment().format('YYYYMMDD-HHmmss'), `${sheetName}!${TEST_RESULT_COLUMN}${TEST_NAME_START_ROW - 1}:${TEST_RESULT_COLUMN}${TEST_NAME_START_ROW - 1}`, this.spreadSheetID);
  }

}

let reportSets = [];

/**
 * @name getGReport
 * @description the factory will help to get the right report object by it spreadSheetID
 * @param {String} spreadSheetID the spreadSheetID from URL
 */
async function getGReport(spreadSheetID) {
  // Check if sheet is in list
  if (!reportSets || !reportSets[spreadSheetID]) {
    let newGReport = new GReport(spreadSheetID);
    reportSets[spreadSheetID] = newGReport;
    return newGReport;
  }
  else
    return reportSets[spreadSheetID];
}


/**
 * @name findTestByName
 * @description find test by testName in the sheet sheetName
 * @param {String} testName 
 * @param {String} sheetName 
 * @param {Boolean} allowExistingResult TRUE: find Test which has result or not; FALSE: find the test which does not has Result
 * @param {String} spreadSheetID the spreadSheetID from URL
 */
async function findTestByName(testName, sheetName, allowExistingResult, spreadSheetID) {
  const foundReport = await getGReport(spreadSheetID);
  if (foundReport)
    return foundReport.findTestByName(testName, sheetName, allowExistingResult);
  return -1;
}

/**
 * @name createNewResultCol
 * @description create new result column
 * @param {String} sheetName 
 * @param {String} spreadSheetID the spreadSheetID from URL
 */
async function createNewResultCol(sheetName, spreadSheetID) {
  const foundReport = await getGReport(spreadSheetID);
  if (foundReport)
    return foundReport.createNewResultCol(sheetName);
  return -1;
}


module.exports = {
  findTestByName,
  createNewResultCol
}