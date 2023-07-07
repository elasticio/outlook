const { Client } = require('./lib/OutlookClient');

module.exports = async function verifyCredentials(credentials) {
  this.logger.info('Verify Credentials started');
  const client = new Client(this, credentials);
  try {
    await client.getUserInfo();
    this.logger.info('User information is retrieved, credentials are valid');
  } catch (e) {
    this.logger.error('Cannot retrieve user information. Credentials are invalid, please check them');
    throw e;
  }
  return { verified: true };
};
