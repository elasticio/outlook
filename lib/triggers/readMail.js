/* eslint-disable no-param-reassign */
const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');
const { getFolders } = require('../utils/selectViewModels');

let client;

exports.process = async function process(msg, cfg, snapshot) {
  this.logger.info('Poll for New Mail trigger starting...');
  const {
    folderId, startTime, pollOnlyUnreadMail, emitBehavior = 'emitIndividually',
  } = cfg;
  snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || startTime || new Date(0).toISOString();
  this.logger.info(`Getting message from specified folder, starting from: ${snapshot.lastModifiedDateTime}, pollOnlyUnreadMail ${pollOnlyUnreadMail || false}`);
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  const results = await client.getLatestMessages(folderId, snapshot.lastModifiedDateTime, pollOnlyUnreadMail);
  this.logger.debug('Results are received');
  if (results.length > 0) {
    this.logger.info(`New Mails found, going to emit, Emit Behavior: ${emitBehavior}...`);
    switch (emitBehavior) {
      case 'emitIndividually':
        this.logger.info('Starting to emit individually found mails...');
        // eslint-disable-next-line no-restricted-syntax
        for (const result of results) {
          this.logger.debug('Emit Mail...');
          // eslint-disable-next-line no-await-in-loop
          await this.emit('data', messages.newMessageWithBody(result));
        }
        break;
      case 'emitAll':
        this.logger.info('Starting to emit all found mails');
        await this.emit('data', messages.newMessageWithBody({ results }));
        break;
      default:
        throw new Error('Unsupported Emit Behavior');
    }

    const lmdate = new Date(results[results.length - 1].lastModifiedDateTime);
    // The output value has always 0 milliseconds
    // we need to set the milliseconds value to 999 in order not to see
    // the duplicate results
    lmdate.setMilliseconds(999);
    snapshot.lastModifiedDateTime = lmdate.toISOString();
    this.logger.info(`Processing completed, new lastModifiedDateTime is ${snapshot.lastModifiedDateTime}`);
    await this.emit('snapshot', snapshot);
  } else {
    this.logger.info('No contacts modified since %s were found', snapshot.lastModifiedDateTime);
  }
};

exports.getFolders = getFolders;
