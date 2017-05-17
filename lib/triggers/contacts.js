/* eslint no-console: 0 no-invalid-this: 0*/
'use strict';
const messages = require('elasticio-node').messages;
const co = require('co');
const MicrosoftGraph = require('msgraph-sdk-javascript');
const rp = require('request-promise');
const _ = require('lodash');
const ApiClient = require('../apiClient');

/**
 * This method will be called from elastic.io platform providing following data
 *
 * @param msg incoming message object that contains ``body`` with payload
 * @param cfg configuration that is account information and configuration field values
 * @param snapshot - snapshot that stores the data between the runs
 */
function processAction(msg, cfg, snapshot) {
    console.log('Snapshot is %j', snapshot);

    // Should be in ISO-Date format
    snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || new Date(0).toISOString();

    const self = this;
    // Main loop
    return co(function* mainLoop() {
        console.log('Refreshing an OAuth Token');

        const instance = new ApiClient(cfg, this);

        const newAccessToken = yield instance.getRefreshedToken();

        const client = MicrosoftGraph.init({
            defaultVersion: 'v1.0',
            debugLogging: true,
            authProvider: (done) => {
                done(null, newAccessToken);
            }
        });

        console.log('Selecting contacts that was modified since %s', snapshot.lastModifiedDateTime);
        const contacts = yield client
            .api('/me/contacts')
            .orderby('lastModifiedDateTime asc')
            .top(900)
            .filter('lastModifiedDateTime gt ' + snapshot.lastModifiedDateTime)
            .get();
        const values = contacts.value;
        console.log('Found %s contacts', values.length);
        if (values.length > 0) {
            values.forEach((elem)=> {
                const messageBody = _.omitBy(elem, (value, key) => key.startsWith('@odata.'));
                messageBody.calendarId = cfg.calendarId;
                this.emit('data', messages.newMessageWithBody(messageBody));
            });
            let lmdate = new Date(values[values.length - 1].lastModifiedDateTime);
            // The output value has always 0 milliseconds
            // we need to set the milliseconds value to 999 in order not to see
            // the duplicate results
            lmdate.setMilliseconds(999);
            snapshot.lastModifiedDateTime = lmdate.toISOString();
        } else {
            console.log('No contacts modified since %s were found', snapshot.lastModifiedDateTime);
        }
        console.log('Processing completed, new lastModifiedDateTime is ' + snapshot.lastModifiedDateTime);
        this.emit('snapshot', snapshot);
    }.bind(this));
}

module.exports.process = processAction;
