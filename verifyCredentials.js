const MicrosoftGraph = require("msgraph-sdk-javascript");
const co = require('co');

// This function will be called by the platform to verify credentials
module.exports = function verifyCredentials(credentials, cb) {
  console.log('Credentials passed for verification %j', credentials);
  // Configuring MS Graph access library
  var client = MicrosoftGraph.init({
    defaultVersion: 'v1.0',
    debugLogging: true,
    authProvider: (done) => {
      done(null, cfg.oauth.access_token);
    }
  });
  // Doing verification
  var process = co(function*() {
    console.log('Fetching user information');
    var user = yield client.api('/me').get();
    console.log('Found user', user);
  });
  process.then(function () {
    console.log('Verification completed');
    cb(null, {verified: true});
  }).catch(err => {
    console.log('Error occured', err.stack || err);
    cb(null , {verified: false});
  });
};
