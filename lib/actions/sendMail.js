const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');

exports.process = async function process(msg, cfg) {
  this.logger.info('Send Mail action started');
  const client = new OutlookClient(this, cfg);
  const result = await client.sendMail(msg);
  this.logger.info(`${JSON.stringify(result, null, 2)}`);
  this.logger.info('Send Mail action successfully executed');
  return messages.newMessageWithBody(result);
};
