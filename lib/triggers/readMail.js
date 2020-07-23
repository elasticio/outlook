/* eslint-disable no-param-reassign */
const { messages } = require('elasticio-node');
const { Client } = require('../Client');

exports.process = async function process(msg, cfg, snapshot) {
  this.logger.trace('Incoming configuration: %j', cfg);
  this.logger.trace('Incoming message: %j', msg);
  this.logger.trace('Incoming snapshot: %j', snapshot);
  const { folderId } = cfg;
  const client = new Client(this, cfg);
  this.logger.info(`Getting message from folder with id: ${folderId}`);
  snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || new Date(0).toISOString();
  const result = await client.getLatestMessages(folderId, snapshot.lastModifiedDateTime);
  this.logger.trace('Result: %j', result);
  if (result.length > 0) {
    await this.emit('data', messages.newMessageWithBody(result));
    const lmdate = new Date(result[result.length - 1].lastModifiedDateTime);
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

exports.getFolders = async function getFolders(cfg) {
  const client = new Client(this, cfg);
  this.logger.info('Getting list of folders...');
  const items = await client.getMailFolders();
  this.logger.trace('Found folders: %j', items);
  this.logger.info('Processing list of folders...');
  const result = {};
  items.forEach((item) => {
    result[item.id] = item.displayName;
  });
  this.logger.trace('Result: %j', result);
  return result;
};
