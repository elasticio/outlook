const { messages } = require('elasticio-node');
const { OutlookClient } = require('../OutlookClient');
const { getFolders } = require('../utils/selectViewModels');

let client;

exports.process = async function process(msg, cfg) {
  if (!client) client = new OutlookClient(this, cfg);
  client.setLogger(this.logger);
  const { originalMailFolders, destinationFolder } = cfg;
  const { messageId } = msg.body;
  let destinationId;
  let currentFolder;
  if (destinationFolder) {
    destinationId = destinationFolder;
    currentFolder = await client.getMailFolderById(destinationId);
  } else {
    const deleteditemsFolder = await client.getDeletedItemsFolder();
    this.logger.trace('deleteditemsFolder found');
    currentFolder = deleteditemsFolder;
    destinationId = deleteditemsFolder.id;
  }
  const result = await client.moveMessage(originalMailFolders, messageId, destinationId);
  result.currentFolder = currentFolder;
  return messages.newMessageWithBody(result);
};

exports.getFolders = getFolders;
