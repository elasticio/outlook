const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const ApiClient = require('../apiClient');

module.exports.process = processAction;

async function processAction(msg, cfg) {
    const client = ApiClient(cfg);
    const currentTime = msg.body.time || (new Date()).toISOString();
    const events = await client.get(`/me/events?$filter=start/dateTime le '${currentTime}' 
      and end/dateTime ge '${currentTime}'`);

    return messages.newMessageWithBody({
        available: !events.value.length
    });
}
