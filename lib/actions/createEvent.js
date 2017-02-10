'use strict';
const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const _ = require('lodash');

const ApiClient = require('../apiClient');
const processEventData = require('../processEventDataHelper').processEventData;

function getCalendars(cfg, cb) {

    function processData(items) {
        console.log('Processing calendar data ...');
        let result = {};
        _.forEach(items.value, function setItem(item) {
            result[item.id] = item.name;
        });
        return result;
    }


    console.log('Getting calendar data ...');
    const instance = new ApiClient(cfg);
    return instance
    .get('/me/calendars')
    .then(processData)
    .nodeify(cb);
}


function processAction(msg, cfg) {

    const self = this;
    const calendarId = cfg.calendarId;
    console.log(`About to create event in calendar with id=${calendarId}`);

    const apiCall = `/me/calendars/${calendarId}/events`;

    const instance = new ApiClient(cfg, self);

    function createEvent(postRequestBody) {
        console.log('Creating Event with properties:');
        console.log(JSON.stringify(postRequestBody, null, 2));
        return instance.post(apiCall, postRequestBody);
    }

    function emitData(data) {
        const id = data.id;
        console.log('Successfully created event with ID=' + id);
        const messageBody = _.omitBy(data, (value, key) => key.startsWith('@odata.'));
        messageBody.calendarId = cfg.calendarId;
        console.log('Emitting data ...');
        console.log('%s', messageBody);
        self.emit('data', messages.newMessageWithBody(messageBody));
    }

    function emitError(e) {
        console.log('Oops! Error occurred');
        self.emit('error', e);
    }

    function emitEnd() {
        console.log('Finished execution!');
        self.emit('end');
    }


    let promise = processEventData(cfg, msg.body)
    .then(createEvent)
    .then(emitData)
    .fail(emitError);

    return promise.finally(emitEnd);

}

module.exports.process = processAction;
module.exports.getCalendars = getCalendars;

