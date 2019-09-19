// this file to run for cli
const {setConf} = require('./gsheet');

const args = process.argv;
let initDir = undefined; // let GCONF_DIR = 'gconf';
let initCredentialFile = undefined; // let GCONF_CREDENTIAL_FILE = "gsheet-auth.json";
let initWebTokenfile = undefined; // let GCONF_TOKEN_PATH = 'token.json';

// read
for (let i = 0 ; i < args.length ; i ++) {
  const currentArgv = args[i];
  if (currentArgv === '--initDir' && i < (args.length -1))
    initDir = args[i + 1];
  if (currentArgv === '--initCredentialFile' && i < (args.length -1))
    initCredentialFile = args[i + 1];
  if (currentArgv === '--initWebTokenfile' && i < (args.length -1))
    initWebTokenfile = args[i + 1];
}

setConf(initDir, initCredentialFile, initWebTokenfile);
console.log('End of init');
