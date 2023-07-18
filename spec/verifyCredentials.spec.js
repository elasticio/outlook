const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
require('./commons');

const { expect } = chai;
const logger = Logger.getLogger();
const sinon = require('sinon');
const verifyCredentials = require('../verifyCredentials');
const { OutlookClient } = require('../lib/OutlookClient');

const credentials = require('./data/configuration.new.in.json');

describe('Verify Credentials', () => {
  let self;
  beforeEach(() => {
    self = {
      logger,
    };
  });
  afterEach(() => {
    sinon.restore();
  });

  it('should succeed verify credentials', async () => {
    sinon.stub(OutlookClient.prototype, 'getUserInfo').callsFake(() => {});
    const result = await verifyCredentials.call(self, credentials);
    expect(result).to.eql({ verified: true });
  });

  it('should fail verify credentials', async () => {
    sinon.stub(OutlookClient.prototype, 'getUserInfo').throws(new Error('Unauthorized'));
    try {
      await verifyCredentials.call(self, credentials);
    } catch (error) {
      expect(error.message).to.eql('Unauthorized');
    }
  });
});
