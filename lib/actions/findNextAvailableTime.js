const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const ApiClient = require('../apiClient');

module.exports.process = processAction;

async function processAction(msg, cfg) {
    const client = ApiClient(cfg);
    let time = msg.body.time || (new Date()).toISOString();
    const events = await client.get(`/me/events?$filter=start/dateTime le '${time}' 
      and end/dateTime ge '${time}'`);
    if (events.value.length) {
        time = events.value.pop().end.dateTime;
    }

    return messages.newMessageWithBody({
        time,
        subject: msg.body.subject
    });
}
