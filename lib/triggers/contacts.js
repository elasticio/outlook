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
function processAction(msg, cfg) {
  var self = this;

  var client = MicrosoftGraph.init({
    defaultVersion: 'v1.0',
    debugLogging: true,
    authProvider: (done) => {
        done(null, cfg.oauth.access_token);
    }
  });

  var process = co(function*() {
    console.log('Fetching user information');
    var user = yield client.api('/me').get();
    console.log('Found user', user);
  });
  process.then(function () {
    console.log('Processing completed');
    self.emit('end');
  }).catch(err => {
    console.log('Error occured', err.stack || err);
    self.emit('error', err);
  });

}
