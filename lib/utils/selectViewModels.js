/* eslint-disable no-restricted-syntax,no-param-reassign,no-await-in-loop */
const { Client } = require('../Client');

async function getChildFolders(client, parentId, result) {
  const childResults = await client.listChildFolders(parentId);
  for (const childResult of childResults) {
    result[childResult.id] = childResult.displayName;
    if (childResult.childFolderCount > 0) {
      await getChildFolders(client, childResult.id, result);
    }
  }
}

exports.getFolders = async function getFolders(cfg) {
  const client = new Client(this, cfg);
  this.logger.info('Getting list of folders...');
  const items = await client.getMailFolders();
  this.logger.trace('Found folders: %j', items);
  this.logger.info('Processing list of folders...');
  const result = {};
  for (const item of items) {
    result[item.id] = item.displayName;
    if (item.childFolderCount > 0) {
      await getChildFolders(client, item.id, result);
    }
  }
  this.logger.trace('Result: %j', result);
  return result;
};
