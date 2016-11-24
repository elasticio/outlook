'use strict';
const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const _ = require('lodash');
const moment = require('moment');
const Helper = require('../helper');
const enumData = require('../data/createEventEnumData.json');
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

function getEnumInfo(enumName, cb) {

  function processData(enumName) {
    let data = enumData[enumName];
    if (!data) {
      throw new Error(enumName + ' info not found!');
    }
    let result = {};
    _.forEach(data, function setItem(item) {
      result[item] = item;
    });
    return result;
  }

  return cb(null, processData(enumName));
}

function getImportanceEnum(cfg, cb) {
  return getEnumInfo("importance", cb);
}
function getTimeZoneEnum(cfg, cb) {
  return getEnumInfo("timeZone", cb);
}
function getShowAsEnum(cfg, cb) {
  return getEnumInfo("showAs", cb);
}
function getSensitivityEnum(cfg, cb) {
  return getEnumInfo("sensitivity", cb);
}

function processEventData(messageBody) {
  console.log('Processing input data...');
  let postRequestBody = _.cloneDeep(messageBody);
  let allDay = (postRequestBody.isAllDay) && (postRequestBody.isAllDay === '1');

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
  console.log("Event Importance:" + cfg.importance);
  console.log("Event ShowAs:" + cfg.showAs);
  console.log("Event Sensitivity:" + cfg.sensitivity);
  console.log("Event TimeZone:" + cfg.timeZone);
  console.log("Event AllDay:" + cfg.allDay);

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
module.exports.getShowAsEnum = getShowAsEnum;
module.exports.getImportanceEnum = getImportanceEnum;
module.exports.getSensitivityEnum = getSensitivityEnum;
module.exports.getTimeZoneEnum = getTimeZoneEnum;
