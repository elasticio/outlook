const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');
const sinon = require('sinon');

const { expect } = chai;
const logger = Logger.getLogger();
const { getCalendars } = require('../../lib/actions/createEvent');

const configuration = require('../data/configuration.new.in.json');
const jsonOut = require('../data/getCalendars_test.out.json');

const cfgString = JSON.stringify(configuration);
describe('Outlook Get Calendars', () => {
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = '/me/calendars';

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    self = {
      emit: sinon.spy(),
      logger,
    };
  });

  it('should return calendar info on success get request', (done) => {
    nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    nock(microsoftGraphUri).get(microsoftGraphApi)
      .reply(200, jsonOut);

    function checkResults(data) {
      expect(data).to.eql({
        'AAMkAGI2TGuLAAA=': 'Calendar',
      });
    }

    getCalendars.call(self, cfg)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('should return errors on refresh token failure ', (done) => {
    nock(refreshTokenUri).post(refreshTokenApi)
      .reply(401);

    function checkError(err) {
      expect(err.message).to.contains('Error in authentication.  Status code: 401');
    }

    getCalendars.call(self, cfg)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });
});
