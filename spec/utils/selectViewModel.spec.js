const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');
const sinon = require('sinon');

const { expect } = chai;
const logger = Logger.getLogger();
const { getFolders } = require('../../lib/utils/selectViewModels');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOutMailFolders = require('../data/mailFolders_test.out.json');

describe('Outlook Read Mail', () => {
  const folderId = 'folderId';
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphMailFolders = '/me/mailFolders?$top=100';

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

  it('getFolder test', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphMailFolders)
      .reply(200, jsonOutMailFolders);

    const result = await getFolders.call(self, cfg);
    expect(result).to.eql({
      1: 'Drafts',
      2: 'Inbox',
    });
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
  });

  it('getFolder test with child', async () => {
    const parentFolder = {
      id: 1,
      displayName: 'Level 1',
      childFolderCount: 1,
    };
    const childFolder = {
      id: 2,
      displayName: 'Level 2',
      childFolderCount: 1,
    };
    const childChildFolder = {
      id: 3,
      displayName: 'Level 3',
      childFolderCount: 0,
    };
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).get(microsoftGraphMailFolders)
      .reply(200, { value: [parentFolder] });

    const scope3 = nock(microsoftGraphUri).get(`/me/mailFolders/${parentFolder.id}/childFolders?$top=100`)
      .reply(200, { value: [childFolder] });

    const scope4 = nock(microsoftGraphUri).get(`/me/mailFolders/${childFolder.id}/childFolders?$top=100`)
      .reply(200, { value: [childChildFolder] });

    const result = await getFolders.call(self, cfg);
    expect(result).to.eql({
      1: 'Level 1',
      2: 'Level 2',
      3: 'Level 3',
    });
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
    expect(scope3.isDone()).to.eql(true);
    expect(scope4.isDone()).to.eql(true);
  });
});
