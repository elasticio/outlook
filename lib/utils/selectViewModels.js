const { Client } = require('../Client');

exports.getFolders = async function getFolders(cfg) {
  const client = new Client(this, cfg);
  this.logger.info('Getting list of folders...');
  const items = await client.getMailFolders();
  this.logger.trace('Found folders: %j', items);
  this.logger.info('Processing list of folders...');
  const result = {};
  items.forEach((item) => {
    result[item.id] = item.displayName;
  });
  this.logger.trace('Result: %j', result);
  return result;
};
