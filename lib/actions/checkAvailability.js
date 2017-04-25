const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const ApiClient = require('../apiClient');

module.exports.process = processAction;

async function processAction(msg, cfg) {
    try {
        const client = ApiClient(cfg);
        let currentTime = msg.body.time || (new Date()).toISOString();
        let events = await client.get(`/me/events?$filter=start/dateTime le '${currentTime}' 
          and end/dateTime ge '${currentTime}'`);
        this.emit('data', messages.newMessageWithBody({
            available: !events.value.length
        }));
    } catch (e) {
        console.log(e.error);
        this.emit('error', e);
    } finally {
        this.emit('end');
    }
}
