const { messages } = require('elasticio-node');
const _ = require('lodash');

const { processEventData } = require('../processEventDataHelper');
const { OutlookClient } = require('../OutlookClient');

let client;

exports.process = async function process(msg, cfg) {
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
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
    await this.emit('error', e);
  }
};

exports.getCalendars = async function getCalendars(cfg) {
  if (!client) client = new OutlookClient(this, cfg);
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
