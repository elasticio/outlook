const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('dotenv').config();

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/checkAvailability');

describe('Outlook Check Availability', () => {
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

  it('Process Action', async () => {
    const msg = {
      body: {
        time: '2023-07-20T10:00:00.0000000',
      },
    };
    const result = await action.process.call(self, msg, cfg);
    expect(result.body).to.have.property('available');
  });
});
