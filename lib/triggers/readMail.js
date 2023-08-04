/* eslint-disable no-param-reassign, no-restricted-syntax, no-loop-func */
const { messages } = require('elasticio-node');
const commons = require('@elastic.io/component-commons-library');
const { OutlookClient } = require('../OutlookClient');
const { getFolders } = require('../utils/selectViewModels');
const {
  getUserAgent,
} = require('../utils/utils');

/**
 * @type {OutlookClient}
 */
let client;

exports.process = async function process(msg, cfg, snapshot) {
  this.logger.info('Poll for New Mail trigger starting...');
  const {
    folderId, startTime, pollOnlyUnreadMail, emitBehavior = 'emitIndividually', getAttachment,
  } = cfg;
  snapshot.lastModifiedDateTime = snapshot.lastModifiedDateTime || startTime || new Date(0).toISOString();
  this.logger.info(`Getting message from specified folder, starting from: ${snapshot.lastModifiedDateTime}, pollOnlyUnreadMail ${pollOnlyUnreadMail || false}`);
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  const results = await client.getLatestMessages(folderId, snapshot.lastModifiedDateTime, pollOnlyUnreadMail);
  this.logger.debug('Results are received');
  if (results.length > 0) {
    this.logger.info(`New Mails found, going to emit, Emit Behavior: ${emitBehavior}...`);
    if (getAttachment) {
      const mailsWithAttachments = results.filter((mail) => mail.hasAttachments);
      if (mailsWithAttachments.length > 0) {
        this.logger.info(`${mailsWithAttachments.length} mails with attachments found`);
        const attachmentsProcessor = new commons.AttachmentProcessor(getUserAgent(), msg.id);
        for (const mail of mailsWithAttachments) {
          const mailAttachments = await client.getAttachments(folderId, mail.id);
          this.logger.info(`${mailAttachments.length} attachments in mail found, going to download`);
          for (const mailAttachment of mailAttachments) {
            const getAttachmentFunction = async () => (await client.downloadAttachment(folderId, mail.id, mailAttachment.id)).data;
            const attachmentId = await attachmentsProcessor.uploadAttachment(getAttachmentFunction);
            const attachmentUrl = attachmentsProcessor.getMaesterAttachmentUrlById(attachmentId);
            const attachment = {
              name: mailAttachment.name,
              contentType: mailAttachment.contentType,
              size: mailAttachment.size,
              url: attachmentUrl,
            };
            if (mail.attachments) {
              mail.attachments.push(attachment);
            } else {
              mail.attachments = [attachment];
            }
          }
        }
      } else {
        this.logger.info('Mails with attachments not found');
      }
    }
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
