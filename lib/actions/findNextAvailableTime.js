const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const ApiClient = require('../apiClient');

module.exports.process = processAction;

async function processAction(msg, cfg) {
  try {
    const client = ApiClient(cfg);
    let time = msg.body.time || (new Date()).toISOString();
    let events = await client.get(`/me/events?$filter=start/dateTime le '${time}' 
          and end/dateTime ge '${time}'`);
    if (events.value.length) {
      time = events.value.pop().end.dateTime
    }
    this.emit('data', messages.newMessageWithBody({
      time: time,
      subject: msg.body.subject
    }));
  } catch (e) {
    console.log(e.error);
    this.emit('error', e);
  } finally {
    this.emit('end');
  }
}