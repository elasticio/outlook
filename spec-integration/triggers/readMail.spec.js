const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('dotenv').config();

const { expect } = chai;
const logger = Logger.getLogger();
const trigger = require('../../lib/triggers/readMail');

describe('Outlook Read Mail', () => {
  let self;
  let cfg;
  beforeEach(() => {
    self = {
      emit: sinon.spy(),
      logger,
    };
    cfg = {
      oauth2: {
        token_type: 'Bearer',
        scope: 'openid User.Read Contacts.Read profile Calendars.ReadWrite Mail.ReadWrite Files.ReadWrite.All',
        expires_in: 3600,
        ext_expires_in: 3600,
        tokenExpiryTime: (new Date('2995-12-17T03:24:00')).toISOString(),
        client_id: process.env.OAUTH_CLIENT_ID,
        client_secret: process.env.OAUTH_CLIENT_SECRET,
        access_token: process.env.ACCESS_TOKEN,
      },
      folderId: process.env.FOLDER_ID,
    };
  });

  it('getFolder test', async () => {
    const result = await trigger.getFolders.call(self, cfg);
    expect(result).to.not.eql({});
  });

  it('process test', async () => {
    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(args[0][0]).to.be.eql('data');
    expect(args[0][1].body.id).to.not.eql(0);
    expect(args[callCount - 2][0]).to.be.eql('snapshot');
    expect(args[callCount - 1][0]).to.be.eql('end');
  });
});
