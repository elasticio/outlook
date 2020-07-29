const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');
const sinon = require('sinon');

const { expect } = chai;
const logger = Logger.getLogger();
const _ = require('lodash');
const trigger = require('../../lib/triggers/contacts');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOut = {
  // eslint-disable-next-line global-require
  value: require('../data/getContact_test.out.json'),
};

describe('Outlook Contacts', () => {
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = '/me/contacts';

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    self = {
      emit: sinon.spy(),
      logger,
    };
  });

  it('should emit data without @oauth property - case: http 200', (done) => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
      .query({
        $orderby: 'lastModifiedDateTime asc',
        $top: 900,
        $filter: 'lastModifiedDateTime gt 1970-01-01T00:00:00.000Z',
      })
      .reply(200, jsonOut);

    function checkResults() {
      const { callCount, args } = self.emit;
      expect(callCount).to.eql(4);
      expect(args[0][0]).to.eql('updateKeys');
      expect(args[1][0]).to.eql('data');
      expect(args[1][1].body).to.eql(Object.assign(
        _.omit(jsonOut.value[0], '@odata.etag'),
        {
          calendarId: cfg.calendarId,
        },
      ));
      expect(args[2][0]).to.eql('data');
      expect(args[2][1].body).to.eql(Object.assign(
        _.omit(jsonOut.value[1], '@odata.etag'),
        {
          calendarId: cfg.calendarId,
        },
      ));
      expect(scope1.isDone()).to.eql(true);
      expect(scope2.isDone()).to.eql(true);
    }

    trigger.process.call(self, {}, cfg, {})
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('should emit error on unsuccessful refresh token request', (done) => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(401);

    trigger.process.call(self, {}, cfg, {})
      .then(() => done.fail(new Error('Error is expected')))
      .catch((err) => {
        expect(err.message).to.contains('Error in authentication.  Status code: 401');
        expect(scope1.isDone()).to.eql(true);
        done();
      });
  });
});
