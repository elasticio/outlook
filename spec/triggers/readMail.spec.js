const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');
const sinon = require('sinon');

const { expect } = chai;
const logger = Logger.getLogger();
const trigger = require('../../lib/triggers/readMail');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOut = require('../data/readMail_test.out.json');

describe('Outlook Read Mail', () => {
  const folderId = 'folderId';
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = `/me/mailFolders/${folderId}/messages`;

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    cfg.folderId = folderId;
    self = {
      emit: sinon.spy(),
      logger,
    };
  });

  it('should emit data', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
      .query({
        $orderby: 'lastModifiedDateTime asc',
        $top: 1000,
        $filter: 'lastModifiedDateTime gt 1970-01-01T00:00:00.000Z',
      })
      .reply(200, jsonOut);

    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(5);
    expect(args[0][0]).to.eql('updateKeys');
    expect(args[1][0]).to.eql('data');
    expect(args[1][1].body).to.eql(jsonOut.value[0]);
    expect(args[2][0]).to.eql('data');
    expect(args[2][1].body).to.eql(jsonOut.value[1]);
    expect(args[3][0]).to.eql('snapshot');
    expect(args[3][1]).to.eql({
      lastModifiedDateTime: '2020-07-20T11:44:53.999Z',
    });
    expect(args[4][0]).to.eql('end');
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
  });

  it('should poll only unread message', async () => {
    cfg.pollOnlyUnreadMail = true;
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
      .query({
        $orderby: 'lastModifiedDateTime asc',
        $top: 1000,
        $filter: 'lastModifiedDateTime gt 1970-01-01T00:00:00.000Z and isRead eq false',
      })
      .reply(200, jsonOut);

    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(5);
    expect(args[0][0]).to.eql('updateKeys');
    expect(args[1][0]).to.eql('data');
    expect(args[1][1].body).to.eql(jsonOut.value[0]);
    expect(args[2][0]).to.eql('data');
    expect(args[2][1].body).to.eql(jsonOut.value[1]);
    expect(args[3][0]).to.eql('snapshot');
    expect(args[3][1]).to.eql({
      lastModifiedDateTime: '2020-07-20T11:44:53.999Z',
    });
    expect(args[4][0]).to.eql('end');
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
  });

  it('should poll only unread message and emit all', async () => {
    cfg.pollOnlyUnreadMail = true;
    cfg.emitBehavior = 'emitAll';
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
      .query({
        $orderby: 'lastModifiedDateTime asc',
        $top: 1000,
        $filter: 'lastModifiedDateTime gt 1970-01-01T00:00:00.000Z and isRead eq false',
      })
      .reply(200, jsonOut);

    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(4);
    expect(args[0][0]).to.eql('updateKeys');
    expect(args[1][0]).to.eql('data');
    expect(args[1][1].body).to.eql({ results: jsonOut.value });
    expect(args[2][0]).to.eql('snapshot');
    expect(args[2][1]).to.eql({
      lastModifiedDateTime: '2020-07-20T11:44:53.999Z',
    });
    expect(args[3][0]).to.eql('end');
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
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
