const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');

let client;

exports.process = async function process(msg, cfg) {
  this.logger.info('Send Mail action started');
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  await client.sendMail(msg);
  this.logger.info('Send Mail action successfully executed');
  return messages.newMessageWithBody(msg.body);
};
