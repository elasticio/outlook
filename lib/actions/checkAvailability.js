const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');

let client;

exports.process = async function process(msg, cfg) {
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  const currentTime = msg.body.time || (new Date()).toISOString();
  const events = await client.getMyLatestEvents(currentTime);
  return messages.newMessageWithBody({
    available: !events.length,
  });
};
