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
    _.forEach(items, function setItem(item) {
      result[item.id] = item.name;
    });
    return result;
  }

  function onSuccess(data) {
    cb(null, processData(data.value));
  }

  function onError(err) {
    console.log('Oops! Error occurred');
    console.log(err.message);
    cb(err);
  }

  const instance = new Helper(cfg, 'Calendars.Read');
  return instance.get('/me/calendars').then(onSuccess).fail(onError).done();

}

function processEventData(messageBody) {
  console.log('Processing input data...');
  let postRequestBody = _.cloneDeep(messageBody);
  let allDay = (!postRequestBody.isAllDay) && (postRequestBody.isAllDay === '1');

  if (allDay) {
    let FORMAT = 'YYYY-MM-DD';
    postRequestBody.isAllDay = 'true';//the data mapper will map this as numeric - setting back as string
    postRequestBody.start.dateTime = moment(postRequestBody.start.dateTime.trim()).format(FORMAT);
    // The last day given as parameter is exclusive.
    // If the user wants to add all day event for 21-23 for ex,
    // outlook expects something like: 21 (00:00 AM) - 24 (00:00 AM)
    postRequestBody.end.dateTime = moment(postRequestBody.end.dateTime.trim()).add(1, 'days').format(FORMAT);
  } else {
    let FORMAT = 'YYYY-MM-DDTHH:mm:ss';
    postRequestBody.isAllDay = 'false';//the data mapper will map this as numeric - setting back as string
    postRequestBody.start.dateTime = moment(postRequestBody.start.dateTime.trim()).format(FORMAT);
    postRequestBody.end.dateTime = moment(postRequestBody.end.dateTime.trim()).format(FORMAT);
  }

  return postRequestBody;
}

function processAction(msg, cfg) {

  const postRequestBody = processEventData(msg.body);
  const self = this;

  function emitData(data) {
    self.emit('data', messages.newMessageWithBody(data));
  }

  const apiCall = `/me/calendars/${cfg.calendarId}/events`;
  baseComponentFunction(self, apiCall, postRequestBody, emitData, cfg, 'Calendars.ReadWrite');
}

module.exports.process = processAction;
module.exports.getCalendars = getCalendars;
