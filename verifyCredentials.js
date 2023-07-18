const { OutlookClient } = require('./lib/OutlookClient');

module.exports = async function verifyCredentials(cfg) {
  this.logger.info('Verify Credentials started');
  const client = new OutlookClient(this, cfg);
  await client.getUserInfo();
  this.logger.info('User information is retrieved, credentials are valid');
  return { verified: true };
};
