/* eslint new-cap: [2, {"capIsNewExceptions": ["Q"]}] */
'use strict';
const messages= require('elasticio-node').messages;
const co = require('co');
const outlook = require('node-outlook');
const bluebird = require('bluebird');

module.exports.process = processAction;

/**
 * This method will be called from elastic.io platform providing following data
 *
 * @param msg incoming message object that contains ``body`` with payload
 * @param cfg configuration that is account information and configuration field values
 */
function processAction(msg, cfg) {
  var self = this;


  var process = co(function* () {
    // Set the API endpoint to use the v2.0 endpoint
    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

    const getContacts = bluebird.promisify(outlook.contacts.getContacts, { context: outlook.contacts });

    // Set up oData parameters
    const queryParams = {
      '$select': 'DisplayName, EmailAddress',
    };

    var user = yield getContacts({
      token: cfg.oauth.access_token,
      odataParams: queryParams
    });
    console.log(user);
  });
  process.then(function() {
    console.log('Processing completed');
    self.emit('end');
  }).catch(err => {
    console.log('Error occured', err.stack || err);
    self.emit('error', err);
  });

}
