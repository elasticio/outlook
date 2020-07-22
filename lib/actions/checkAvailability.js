const { messages } = require('elasticio-node');
const { Client } = require('../Client');

exports.process = async function process(msg, cfg) {
  const client = new Client(this, cfg);
  const currentTime = msg.body.time || (new Date()).toISOString();
  const events = await client.getMyEvents(currentTime);
  return messages.newMessageWithBody({
    available: !events.length,
  });
};
