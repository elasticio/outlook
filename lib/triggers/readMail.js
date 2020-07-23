const { messages } = require('elasticio-node');
const { Client } = require('../Client');

exports.process = async function process(msg, cfg, snapshot) {
  this.logger.trace('Incoming configuration: %j', cfg);
  this.logger.trace('Incoming message: %j', msg);
  this.logger.trace('Incoming snapshot: %j', snapshot);
  const { folderId } = cfg;
  const client = new Client(this.logger, cfg);
  this.logger.info(`Getting message from folder with id: ${folderId}`);
  const result = await client.getMessages(folderId);
  this.logger.trace('Result: %j', result);
  return messages.newMessageWithBody(result);
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
