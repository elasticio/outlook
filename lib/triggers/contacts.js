/* eslint new-cap: [2, {"capIsNewExceptions": ["Q"]}] */
'use strict';
const messages = require('elasticio-node').messages;
const co = require('co');
const MicrosoftGraph = require("msgraph-sdk-javascript");

module.exports.process = processAction;

/**
 * This method will be called from elastic.io platform providing following data
 *
 * @param msg incoming message object that contains ``body`` with payload
 * @param cfg configuration that is account information and configuration field values
 */
function processAction(msg, cfg, snapshot) {
  const self = this;

  let client = MicrosoftGraph.init({
    defaultVersion: 'v1.0',
    debugLogging: true,
    authProvider: (done) => {
        done(null, cfg.oauth.access_token);
    }
  });

  console.log('Snapshot is %j', snapshot);

  // Should be in ISO-Date format
  snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || new Date(0).toISOString();

  // Main loop
  var process = co(function*() {
    console.log('Selecting contacts that was modified since %s', snapshot.lastModifiedDateTime);
    const contacts = yield client
      .api('/me/contacts')
      .orderby('lastModifiedDateTime asc')
      .top(1000)
      .filter('lastModifiedDateTime gt ' + snapshot.lastModifiedDateTime)
      .get();
    const values = contacts.value;
    console.log('Found %s contacts', values.length);
    if (values.length > 0) {
      var message = messages.newMessageWithBody(values);
      self.emit('data', message);
      snapshot.lastModifiedDateTime = values[values.length-1].lastModifiedDateTime;
    }
  });

  process.then(function () {
    console.log('Processing completed, new lastModifiedDateTime is ' + snapshot.lastModifiedDateTime);
    self.emit('snapshot', snapshot);
    self.emit('end');
  }).catch(err => {
    console.log('Error occured', err.stack || err);
    self.emit('error', err);
    self.emit('end');
  });

}
