'use strict';
const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const _ = require('lodash');

const Helper = require('../helper');
const baseComponentFunction = require('../baseComponentFunction');
const processEvent = require('../processEventDataHelper').processEvent;

function getCalendars(cfg, cb) {

  function processData(items) {
    console.log("Processing calendar data ...");
    let result = {};
    _.forEach(items.value, function setItem(item) {
      result[item.id] = item.name;
    });
    return result;
  }

  console.log("Getting calendar data ...");
  const instance = new Helper(cfg, 'Calendars.Read');
  return instance
    .get('/me/calendars')
    .then(processData)
    .nodeify(cb);
}



function processAction(msg, cfg) {

  const postRequestBody = processEvent(cfg, msg.body);
  const self = this;

  function emitData(data) {
    console.log("Emitting data ...");

    let messageBody = {
      id: data.id,
      calendarId : cfg.calendarId
    }
    self.emit('data', messages.newMessageWithBody(messageBody));
  }

  console.log("Creating Event with properties: ");
  console.log(postRequestBody);

  const apiCall = `/me/calendars/${cfg.calendarId}/events`;
  baseComponentFunction(self, apiCall, postRequestBody, emitData, cfg, 'Calendars.ReadWrite');

}

module.exports.process = processAction;
module.exports.getCalendars = getCalendars;

