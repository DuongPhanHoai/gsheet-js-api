
const { findTestByName, updateTestResultByName, createNewResultCol } = require('./greport');
const { setConf } = require('./gsheet');

module.exports = {
  setConf,
  findTestByName,
  updateTestResultByName,
  createNewResultCol
}