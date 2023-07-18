const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');

let client;

exports.process = async function process(msg, cfg) {
  this.logger.info('Check Availability action started');
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  const currentTime = msg.body.time || (new Date()).toISOString();
  this.logger.info('Checking for available slots in the calendar');
  const events = await client.getMyLatestEvents(currentTime);
  this.logger.info(`The time range specified is ${events.length ? 'busy' : 'free'}`);
  const message = messages.newMessageWithBody({
    available: !events.length,
  });
  await this.emit('data', message);
  this.logger.info('Check Availability action successfully executed');
};
