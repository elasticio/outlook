const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('dotenv').config();

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/moveMail');

describe('Outlook Move Mail', () => {
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
    const result = await action.getFolders.call(self, cfg);
    expect(result).to.not.eql({});
  });

  it('process test', async () => {
    cfg.originalMailFolders = 'AQMkADAwATM0MDAAMS0yMgA3MC1mZDkyLTAwAi0wMAoALgAAA14NR8zJ7GxJnz_JauTU1uQBALdeJHYCtWNEt1wFU6jUUYgAAAIBDAAAAA==';
    const msg = {
      body: {
        messageId: 'AQMkADAwATM0MDAAMS0yMgA3MC1mZDkyLTAwAi0wMAoARgAAA14NR8zJ7GxJnz_JauTU1uQHALdeJHYCtWNEt1wFU6jUUYgAAAIBDAAAALdeJHYCtWNEt1wFU6jUUYgAAABiIhOGAAAA',
      },
    };
    const result = await action.process.call(self, msg, cfg);
    expect(result.body.id).to.eql(msg.body.messageId);
  });
});
