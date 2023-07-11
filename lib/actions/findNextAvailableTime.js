const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');

let client;

exports.process = async function process(msg, cfg) {
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
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
