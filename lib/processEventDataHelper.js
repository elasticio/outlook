'use strict';
const moment = require('moment');
const _ = require('lodash');


function formatDate(date, format) {

  if (moment(date).isValid()) {
    return moment(date).format(format);
  }

  if (moment(date, 'x').isValid()) {
    return moment(date, 'x').format(format);
  }

  throw new Error("Invalid date " + date);
}

function processEventData(cfg, messageBody) {

   console.log("Processing event data ...");

  let result = _.cloneDeep(messageBody);

  if (cfg.importance) {
    result.importance = cfg.importance;
  };

  if (cfg.sensitivity) {
    result.sensitivity = cfg.sensitivity;
  };

  if (cfg.showAs) {
    result.showAs = cfg.showAs;
  };

  if (cfg.bodyContentType) {
    result.body.contentType = cfg.bodyContentType;
  };

  let FORMAT = 'YYYY-MM-DDTHH:mm:ss';
  if ((cfg.isAllDay) && (cfg.isAllDay === 'true')) {
    result.isAllDay = cfg.isAllDay;
    // The last day given as parameter is exclusive.
    // If the user wants to add all day event for 21-23 for ex,
    // outlook expects something like: 21 (00:00 AM) - 24 (00:00 AM)
    FORMAT = 'YYYY-MM-DD';
    result.start.dateTime = formatDate(result.start.dateTime.trim(), FORMAT);
    result.end.dateTime = formatDate(moment(result.end.dateTime).add(1,'days'), FORMAT);
  } else {
    result.end.dateTime = formatDate(result.end.dateTime.trim(), FORMAT);
    result.start.dateTime = formatDate(result.start.dateTime.trim(), FORMAT);
  }

  result.end.timeZone = cfg.timeZone;
  result.start.timeZone = cfg.timeZone;

  return result;
}


module.exports.processEvent = processEventData;
