const { Client } = require('./lib/Client');

module.exports = async function verifyCredentials(credentials) {
  this.logger.info('Credentials passed for verification %j', credentials);
  const client = new Client(this, credentials);
  try {
    const userInfo = await client.getUserInfo();
    this.logger.info('Found user: %j', userInfo);
  } catch (e) {
    this.logger.error('Cannot retrieve user information. Credentials is invalid, please check it');
    throw e;
  }
  return { verified: true };
};
