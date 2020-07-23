const { Client } = require('./lib/Client');

module.exports = async function verifyCredentials(credentials) {
  this.logger.trace('Credentials passed for verification %j', credentials);
  const client = new Client(this, credentials);
  try {
    const userInfo = await client.getUserInfo();
    this.logger.info('Found user: %j', userInfo);
  } catch (e) {
    this.logger.error('Cannot retrieve user information. Credentials are invalid, please check them');
    throw e;
  }
  return { verified: true };
};
