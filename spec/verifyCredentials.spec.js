const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');

const { expect } = chai;
const logger = Logger.getLogger();
const verifyCredentials = require('../verifyCredentials');

const credentials = require('./data/configuration.new.in.json');

describe('Verify Credentials', () => {
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = '/me';

  let self;
  beforeEach(() => {
    self = {
      logger,
    };
  });

  it('should succeed verify credentials', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
      .reply(200, {
        displayName: 'PS Team',
        surname: 'Team',
        givenName: 'PS',
        id: '7161fc17af0d3ce4',
      });
    const result = await verifyCredentials.call(self, credentials);
    expect(result).to.eql({ verified: true });
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
  });

  it('should fail verify credentials', (done) => {
    credentials.oauth2.tokenExpiryTime = (new Date('2995-12-17T03:24:00')).toISOString();

    const scope1 = nock(microsoftGraphUri).get(microsoftGraphApi)
      .reply(401);
    verifyCredentials.call(self, credentials)
      .then(() => done.fail(new Error('Error is expected')))
      .catch((err) => {
        expect(err.message).to.contains('Error in making request');
        expect(scope1.isDone()).to.eql(true);
        done();
      });
  });
});
