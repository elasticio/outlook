const { messages } = require('elasticio-node');
const _ = require('lodash');

const { processEventData } = require('../processEventDataHelper');
const { Client } = require('../Client');

exports.process = async function process(msg, cfg) {
  const client = new Client(this, cfg);
  const { calendarId } = cfg;
  this.logger.info('About to create event in specified calendar...');
  const postRequestBody = await processEventData(this.logger, cfg, msg.body);
  this.logger.debug('postRequestBody is created');
  try {
    const data = await client.createEvent(calendarId, postRequestBody);
    this.logger.info('Successfully created event');
    const messageBody = _.omitBy(data, (value, key) => key.startsWith('@odata.'));
    messageBody.calendarId = cfg.calendarId;
    this.logger.info('Emitting data ...');
    await this.emit('data', messages.newMessageWithBody(messageBody));
  } catch (e) {
    this.logger.error('Oops! Error occurred');
    this.emit('error', e);
  }
  this.emit('end');
};

exports.getCalendars = async function getCalendars(cfg) {
  const client = new Client(this, cfg);
  this.logger.info('Getting calendar data ...');
  const items = await client.getMyCalendars();
  this.logger.debug('Found %s calendar(s)', items.length);
  this.logger.info('Processing calendar data ...');
  const result = {};
  items.forEach((item) => {
    result[item.id] = item.name;
  });
  return result;
};
