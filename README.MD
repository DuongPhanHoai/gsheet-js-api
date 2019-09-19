This library will support to update Automation Test report to Google Sheet, the following functions:
  - findTestByName,
  - updateTestResultByName,
  - createNewResultCol

Note: the new function which help to update the authorization function will be completed later
run the command gsheet-js-api to init google authorization with parameter as:
 - --initDir : working dir for google sheet authorization
 - --initCredentialFile : google credential file which download from google project
 - --initWebTokenfile : the file to save webtoken

 Ex: node gsheet-js-api --initDir gconf