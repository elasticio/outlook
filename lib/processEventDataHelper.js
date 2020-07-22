/* eslint no-console: 0 */

const moment = require('moment-timezone');
const _ = require('lodash');
const Q = require('q');

function checkRequiredFields(cfg, messageBody) {
  console.log('Checking required fields ...');
  if (!messageBody.start) {
    throw new Error('Start Time missing! This field is required!');
  }
  if (!messageBody.end) {
    throw new Error('End Time missing! This field is required!');
  }
  if (!cfg.timeZone) {
    throw new Error('Time Zone missing! This field is required!');
  }
  if (!cfg.calendarId) {
    throw new Error('Calendar ID missing! This field is required!');
  }
  if ((cfg.bodyContentType) && (!messageBody.body)) {
    throw new Error('Body Type provided, but Body Content is missing!');
  }
}

function formatDate(date, timeZone, format) {
  // Valid date string
  if (moment(date).isValid()) {
    // Valid ISO_8601 string
    if ((moment(date, moment.ISO_8601).isValid())) {
      console.log(`Date ${date} has a known ISO-8601 format.`);
      return moment.tz(date, timeZone).format(format);
    }
    // non ISO_8601 string
    throw new Error(`non ISO-8601 date formats are currently not supported: ${date}.`);
  }
  // Valid millisecond value
  if (moment(date, 'x').isValid()) {
    return moment(date, 'x').tz(timeZone).format(format);
  }
  throw new Error(`Invalid date ${date}.`);
}

function buildPostBody(cfg, messageBody) {
  checkRequiredFields(cfg, messageBody);

  const result = _.cloneDeep(messageBody);

  if (cfg.importance) {
    result.importance = cfg.importance;
  }

  if (cfg.sensitivity) {
    result.sensitivity = cfg.sensitivity;
  }

  if (cfg.showAs) {
    result.showAs = cfg.showAs;
  }

  if (cfg.bodyContentType) {
    result.body.contentType = cfg.bodyContentType;
  }

  result.end.timeZone = cfg.timeZone;
  result.start.timeZone = cfg.timeZone;

  if ((cfg.isAllDay) && (cfg.isAllDay === 'true')) {
    result.isAllDay = cfg.isAllDay;
    // The last day given as parameter is exclusive.
    // If the user wants to add all day event for 21-23 for ex,
    // outlook expects something like: 21 (00:00 AM) - 24 (00:00 AM)
    const FORMAT = 'YYYY-MM-DD';
    result.start.dateTime = formatDate(result.start.dateTime.trim(), cfg.timeZone, FORMAT);
    result.end.dateTime = formatDate(result.end.dateTime.trim(), cfg.timeZone, FORMAT);
    result.end.dateTime = moment(result.end.dateTime).add(1, 'days').format(FORMAT);
  } else {
    const FORMAT = 'YYYY-MM-DDTHH:mm:ss';
    result.start.dateTime = formatDate(result.start.dateTime.trim(), cfg.timeZone, FORMAT);
    result.end.dateTime = formatDate(result.end.dateTime.trim(), cfg.timeZone, FORMAT);
  }

  return result;
}

function processEventData(cfg, messageBody) {
  console.log('Processing event data ...');
  const defer = Q.defer();
  try {
    const result = buildPostBody(cfg, messageBody);
    defer.resolve(result);
  } catch (err) {
    defer.reject(err);
  }
  return defer.promise;
}

module.exports.processEventData = processEventData;
module.exports.formatDate = formatDate;
module.exports.checkRequiredFields = checkRequiredFields;
