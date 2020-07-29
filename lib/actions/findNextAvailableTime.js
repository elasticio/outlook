const { messages } = require('elasticio-node');
const { Client } = require('../Client');

exports.process = async function process(msg, cfg) {
  const client = new Client(this, cfg);
  let currentTime = msg.body.time || (new Date()).toISOString();
  const events = await client.getMyLatestEvents(currentTime);
  if (events.length) {
    this.logger.info('Defined time is busy, calculating the new available time...');
    currentTime = events.pop().end.dateTime;
  }
  return messages.newMessageWithBody({
    time: currentTime,
    subject: msg.body.subject,
  });
};
