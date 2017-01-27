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
    const apiCall = `/me/calendars/${cfg.calendarId}/events`;

    const instance = new ApiClient(cfg);

    function createEvent(postRequestBody) {
        console.log('Creating Event with properties:');
        console.log(JSON.stringify(postRequestBody, null, 2));
        return instance.post(apiCall, postRequestBody);
    }

    function emitData(data) {
        const id = data.id;
        data.calendarId = cfg.calendarId;
        console.log('Successfully created event with ID=' + id);
        console.log('Emitting data ...');
        console.log('---DATA---');
        console.log('%j', data);
        console.log('---MSG---');
        console.log('%j', messages.newMessageWithBody(data));
        let messageBody = {
            id,
            calendarId: cfg.calendarId
        };
        console.log('Emitting data ...');
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

