const { messages } = require('elasticio-node');
const { Client } = require('../Client');

exports.process = async function process(msg, cfg, snapshot) {
  this.logger.trace('Incoming configuration: %j', cfg);
  this.logger.trace('Incoming message: %j', msg);
  this.logger.trace('Incoming snapshot: %j', snapshot);
  const client = new Client(this.logger, cfg);
  const result = await client.getMyMessages();
  this.logger.info('Result: %j', result);
  return messages.newMessageWithBody(result);
};
