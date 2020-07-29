const { messages } = require('elasticio-node');
const { Client } = require('../Client');
const { getFolders } = require('../utils/selectViewModels');

exports.process = async function process(msg, cfg) {
  const client = new Client(this, cfg);
  const { originalMailFolders, destinationFolder } = cfg;
  const { messageId } = msg.body;
  let destinationId;
  let currentFolder;
  if (destinationFolder) {
    destinationId = destinationFolder;
    currentFolder = await client.getMailFolderById(destinationId);
  } else {
    const deleteditemsFolder = await client.getDeleteditemsFolder();
    this.logger.trace('Found deleteditemsFolder: %j', deleteditemsFolder);
    currentFolder = deleteditemsFolder;
    destinationId = deleteditemsFolder.id;
  }
  const result = await client.moveMessage(originalMailFolders, messageId, destinationId);
  result.currentFolder = currentFolder;
  return messages.newMessageWithBody(result);
};

exports.getFolders = getFolders;
