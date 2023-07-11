/* eslint-disable no-param-reassign */
const { messages } = require('elasticio-node');
const _ = require('lodash');
const { OutlookClient } = require('../OutlookClient');

let client;

/**
 * This method will be called from elastic.io platform providing following data
 *
 * @param msg incoming message object that contains ``body`` with payload
 * @param cfg configuration that is account information and configuration field result
 * @param snapshot - snapshot that stores the data between the runs
 */
exports.process = async function process(msg, cfg, snapshot) {
  this.logger.info('Contacts trigger starting...');
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  // Should be in ISO-Date format
  snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || new Date(0).toISOString();
  this.logger.info('Selecting contacts that was modified since %s', snapshot.lastModifiedDateTime);
  const result = await client.getMyLatestContacts(snapshot.lastModifiedDateTime);
  this.logger.info('Found %s contacts', result.length);
  if (result.length > 0) {
    result.forEach((elem) => {
      const messageBody = _.omitBy(elem, (value, key) => key.startsWith('@odata.'));
      messageBody.calendarId = cfg.calendarId;
      this.emit('data', messages.newMessageWithBody(messageBody));
    });
    const lmdate = new Date(result[result.length - 1].lastModifiedDateTime);
    // The output value has always 0 milliseconds
    // we need to set the milliseconds value to 999 in order not to see
    // the duplicate results
    lmdate.setMilliseconds(999);
    snapshot.lastModifiedDateTime = lmdate.toISOString();
  } else {
    this.logger.info('No contacts modified since %s were found', snapshot.lastModifiedDateTime);
  }
  this.logger.info(`Processing completed, new lastModifiedDateTime is ${snapshot.lastModifiedDateTime}`);
  this.emit('snapshot', snapshot);
};
