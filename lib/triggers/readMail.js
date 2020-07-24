/* eslint-disable no-param-reassign */
const { messages } = require('elasticio-node');
const { Client } = require('../Client');
const { getFolders } = require('../utils/selectViewModels');

exports.process = async function process(msg, cfg, snapshot) {
  this.logger.trace('Incoming configuration: %j', cfg);
  this.logger.trace('Incoming message: %j', msg);
  this.logger.trace('Incoming snapshot: %j', snapshot);
  const { folderId, startTime } = cfg;
  snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || startTime || new Date(0).toISOString();
  this.logger.info(`Getting message from folder with id: ${folderId} starting from: ${snapshot.lastModifiedDateTime}`);
  const client = new Client(this, cfg);
  const results = await client.getLatestMessages(folderId, snapshot.lastModifiedDateTime);
  this.logger.trace('Results: %j', results);
  if (results.length > 0) {
    this.logger.info('New Mails found, going to emit...');
    // eslint-disable-next-line no-restricted-syntax
    for (const result of results) {
      this.logger.trace('Emit Mail: %j', result);
      // eslint-disable-next-line no-await-in-loop
      await this.emit('data', messages.newMessageWithBody(result));
    }
    const lmdate = new Date(results[results.length - 1].lastModifiedDateTime);
    // The output value has always 0 milliseconds
    // we need to set the milliseconds value to 999 in order not to see
    // the duplicate results
    lmdate.setMilliseconds(999);
    snapshot.lastModifiedDateTime = lmdate.toISOString();
    this.logger.info(`Processing completed, new lastModifiedDateTime is ${snapshot.lastModifiedDateTime}`);
    this.emit('snapshot', snapshot);
  } else {
    this.logger.info('No contacts modified since %s were found', snapshot.lastModifiedDateTime);
  }
  this.emit('end');
};

exports.getFolders = getFolders;
