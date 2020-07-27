const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');
const sinon = require('sinon');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/moveMail');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOut = require('../data/moveMail_test.out.json');

describe('Outlook Move Mail', () => {
  const originalMailFolders = 'originalId';
  const destinationFolder = 'destinationId';
  const messageId = 'messageId';
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const msg = {
    body: {
      messageId,
    },
  };

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    cfg.originalMailFolders = originalMailFolders;
    self = {
      emit: sinon.spy(),
      logger,
    };
  });

  it('should move message to specified folder', async () => {
    cfg.destinationFolder = destinationFolder;
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri)
      .post(`/me/mailFolders/${originalMailFolders}/messages/${messageId}/move`,
        {
          destinationId: destinationFolder,
        })
      .reply(200, jsonOut);

    const scope3 = nock(microsoftGraphUri)
      .get(`/me/mailFolders/${destinationFolder}`)
      .reply(200, { id: destinationFolder });

    const result = await action.process.call(self, msg, cfg, {});
    const expectedReult = jsonOut;
    expectedReult.currentFolder = {
      id: destinationFolder,
    };
    expect(result.body).to.eql(expectedReult);
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
    expect(scope3.isDone()).to.eql(true);
  });

  it('should move message to Deleted Items folder', async () => {
    const deletedItemsFolderId = 'deletedItemsFolderId';
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri)
      .get('/me/mailFolders/deleteditems')
      .reply(200, { id: deletedItemsFolderId });

    const scope3 = nock(microsoftGraphUri)
      .post(`/me/mailFolders/${originalMailFolders}/messages/${messageId}/move`,
        {
          destinationId: deletedItemsFolderId,
        })
      .reply(200, jsonOut);

    const result = await action.process.call(self, msg, cfg, {});
    const expectedReult = jsonOut;
    expectedReult.currentFolder = {
      id: 'deletedItemsFolderId',
    };
    expect(result.body).to.eql(expectedReult);
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
    expect(scope3.isDone()).to.eql(true);
  });
});
