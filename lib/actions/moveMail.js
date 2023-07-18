const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');
const { getFolders } = require('../utils/selectViewModels');

let client;

exports.process = async function process(msg, cfg) {
  this.logger.info('Move Mail action started');
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  const { originalMailFolders, destinationFolder } = cfg;
  const { messageId } = msg.body;
  let destinationId;
  let currentFolder;
  if (destinationFolder) {
    destinationId = destinationFolder;
    this.logger.info('Destination Folder is specified. Looking for the mail folder by ID...');
    currentFolder = await client.getMailFolderById(destinationId);
  } else {
    this.logger.info('Destination Folder is not specified. Looking for deleteditems...');
    const deleteditemsFolder = await client.getDeletedItemsFolder();
    this.logger.info('deleteditemsFolder found');
    currentFolder = deleteditemsFolder;
    destinationId = deleteditemsFolder.id;
  }
  this.logger.info('Moving the message...');
  const result = await client.moveMessage(originalMailFolders, messageId, destinationId);
  result.currentFolder = currentFolder;
  this.logger.info('The message has been successfully moved. Move Mail action finished');
  return messages.newMessageWithBody(result);
};

exports.getFolders = getFolders;
