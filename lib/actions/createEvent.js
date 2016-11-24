'use strict';
const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const _ = require('lodash');
const moment = require('moment');
const Helper = require('../helper');
const baseComponentFunction = require('../baseComponentFunction');

function getCalendars(cfg, cb) {

  function processData(items) {
    let result = {};
    _.forEach(items.value, function setItem(item) {
      result[item.id] = item.name;
    });
    return result;
  }

  const instance = new Helper(cfg, 'Calendars.Read');
  return instance
    .get('/me/calendars')
    .then(processData)
    .nodeify(cb);
}

function processEventData(cfg, messageBody) {
  var result = _.cloneDeep(messageBody);

  if (cfg.importance) {
    result.importance = cfg.importance;
  };
  if (cfg.sensitivity) {
    result.sensitivity = cfg.sensitivity;
  };
  if (cfg.showAs) {
    result.showAs = cfg.showAs;
  };
  if (cfg.isAllDay) {
    result.isAllDay = cfg.isAllDay;
  };
  if (cfg.bodyContentType) {
    result.body.contentType = cfg.bodyContentType;
  };
  result.end.timeZone = cfg.timeZone;
  result.start.timeZone = cfg.timeZone;

  let allDay = (result.isAllDay) && (result.isAllDay === 'true');

  if (allDay) {
    let FORMAT = 'YYYY-MM-DD';
    result.start.dateTime = moment(result.start.dateTime.trim()).format(FORMAT);
    // The last day given as parameter is exclusive.
    // If the user wants to add all day event for 21-23 for ex,
    // outlook expects something like: 21 (00:00 AM) - 24 (00:00 AM)
    result.end.dateTime = moment(result.end.dateTime.trim()).add(1, 'days').format(FORMAT);
  } else {
    let FORMAT = 'YYYY-MM-DDTHH:mm:ss';
    result.start.dateTime = moment(result.start.dateTime.trim()).format(FORMAT);
    result.end.dateTime = moment(result.end.dateTime.trim()).format(FORMAT);
  }

  return result;
}

function processAction(msg, cfg) {

  const postRequestBody = processEventData(cfg, msg.body);
  const self = this;

  function emitData(data) {
    let messageBody = {
      id: data.id
    }
    self.emit('data', messages.newMessageWithBody(messageBody));
  }

  const apiCall = `/me/calendars/${cfg.calendarId}/events`;
  baseComponentFunction(self, apiCall, postRequestBody, emitData, cfg, 'Calendars.ReadWrite');

}

module.exports.process = processAction;
module.exports.getCalendars = getCalendars;

